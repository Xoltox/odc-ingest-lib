using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text;
using Microsoft.Extensions.Logging;
using OdcIngest.Extractors;
using OdcIngest.Interfaces;
using OdcIngest.Storage;
using OdcIngest.Structures;

[assembly: InternalsVisibleTo("OdcIngestLib.UnitTests")]
[assembly: InternalsVisibleTo("OdcIngestLib.IntegrationTests")]

namespace OdcIngest.Implementations
{
    public class OdcIngestLibImpl : IOdcIngestLib
    {
        /// <summary>Total JSON-escaped bytes of all Markdown returned by one call.</summary>
        public const int MaxOutputBytes = 4_000_000;
        /// <summary>A single page larger than the budget is cut here, then "[truncated]" is appended.</summary>
        public const int MaxSinglePageBytes = 3_900_000;
        public const int MaxPagesPerCall = 50;
        public const int MaxRowsPerCall = 500;
        /// <summary>Time guard: no further page is added after this many seconds (platform limit is 95).</summary>
        public const int TimeBudgetSeconds = 70;

        private readonly ILogger _logger;
        private readonly Func<StorageConnection, IObjectSource> _sourceFactory;

        public OdcIngestLibImpl(ILogger logger)
        {
            _logger = logger;
            _sourceFactory = c => new S3ObjectSource(c, logger);
        }

        internal OdcIngestLibImpl(ILogger logger, Func<StorageConnection, IObjectSource> sourceFactory)
        {
            _logger = logger;
            _sourceFactory = sourceFactory;
        }

        /// <summary>Overridable by tests only.</summary>
        internal TimeSpan TimeBudget { get; set; } = TimeSpan.FromSeconds(TimeBudgetSeconds);

        /// <summary>Hard deadline enforced inside every extractor (overridable by tests only).</summary>
        internal TimeSpan HardTimeBudget { get; set; } = DocContext.DefaultHardLimit;

        /// <summary>Context of the most recent call (tests read DocumentOpens).</summary>
        internal DocContext? LastContext { get; private set; }

        private DocContext NewContext() => LastContext = new DocContext(HardTimeBudget);

        /// <summary>Conservative size of a string once JSON-escaped (non-ASCII, quotes and angle brackets counted as 6-byte unicode escapes).</summary>
        internal static long EscapedBytes(string s)
        {
            long n = 0;
            foreach (var c in s) n += EscLen(c);
            return n;
        }

        private static int EscLen(char c)
        {
            if (c == '\n' || c == '\r' || c == '\t' || c == '\b' || c == '\f') return 2;
            if (c < 0x20 || c >= 0x7F || c == '"' || c == '\\' || c == '<' || c == '>' || c == '&' || c == '\'' || c == '+' || c == '`') return 6;
            return 1;
        }

        public int GetPageCount(StorageConnection connection, string objectKey)
        {
            var sw = Stopwatch.StartNew();
            return Guarded(() =>
            {
                var ctx = NewContext();
                var path = _sourceFactory(connection).GetLocalPath(objectKey);
                int count;
                using (var session = FormatDetector.CreateExtractor(path).Open(path, ctx)) count = session.PageCount;
                _logger.LogInformation("OdcIngest: GetPageCount key={ObjectKey} pages={Pages} elapsedMs={ElapsedMs}", objectKey, count, sw.ElapsedMilliseconds);
                return count;
            });
        }

        public IEnumerable<PageRecord> ExtractPages(StorageConnection connection, string objectKey, int fromPage, int toPage)
        {
            var sw = Stopwatch.StartNew();
            return Guarded(() =>
            {
                var ctx = NewContext();
                var path = _sourceFactory(connection).GetLocalPath(objectKey);
                using var session = FormatDetector.CreateExtractor(path).Open(path, ctx);
                int count = session.PageCount;

                if (fromPage < 1 || fromPage > count) throw new InvalidOperationException("OdcIngest: page out of range");
                if (toPage < fromPage) throw new InvalidOperationException("OdcIngest: page out of range");
                int last = Math.Min(Math.Min(toPage, fromPage + MaxPagesPerCall - 1), count);

                var result = new List<PageRecord>();
                long total = 0;
                using (var e = session.Pages(fromPage, last).GetEnumerator())
                {
                    while (true)
                    {
                        if (result.Count > 0 && sw.Elapsed >= TimeBudget) break;
                        if (!e.MoveNext()) break;
                        var page = LimitPage(e.Current);
                        long bytes = EscapedBytes(page.Markdown);
                        if (result.Count > 0 && total + bytes > MaxOutputBytes) break;
                        result.Add(page);
                        total += bytes;
                    }
                }
                _logger.LogInformation(
                    "OdcIngest: ExtractPages key={ObjectKey} from={From} to={To} returned={Returned} bytes={Bytes} elapsedMs={ElapsedMs}",
                    objectKey, fromPage, last, result.Count, total, sw.ElapsedMilliseconds);
                return (IEnumerable<PageRecord>)result;
            });
        }

        public IEnumerable<QuestionnaireRow> ExtractQuestionnaireRows(StorageConnection connection, string objectKey, string? sheetName = null, int fromRow = 1, int toRow = 500)
        {
            var sw = Stopwatch.StartNew();
            return Guarded(() =>
            {
                var ctx = NewContext();
                var path = _sourceFactory(connection).GetLocalPath(objectKey);
                if (FormatDetector.Detect(path) != DocFormat.Xlsx) throw new InvalidOperationException("OdcIngest: not a workbook");
                if (fromRow < 1) fromRow = 1;
                if (toRow < fromRow) throw new InvalidOperationException("OdcIngest: row out of range");

                var result = new List<QuestionnaireRow>();
                long total = 0;
                foreach (var row in new XlsxExtractor().ExtractRows(path, sheetName, fromRow, toRow, ctx))
                {
                    long bytes = RowBytes(row);
                    if (result.Count > 0 && (result.Count >= MaxRowsPerCall || total + bytes > MaxOutputBytes)) break;
                    result.Add(row);
                    total += bytes;
                    if (sw.Elapsed >= TimeBudget) break;
                }
                _logger.LogInformation(
                    "OdcIngest: ExtractQuestionnaireRows key={ObjectKey} rows={Rows} bytes={Bytes} elapsedMs={ElapsedMs}",
                    objectKey, result.Count, total, sw.ElapsedMilliseconds);
                return (IEnumerable<QuestionnaireRow>)result;
            });
        }

        private static long RowBytes(QuestionnaireRow r) =>
            EscapedBytes(r.SheetName ?? "") + EscapedBytes(r.RefId ?? "") + EscapedBytes(r.Section ?? "") +
            EscapedBytes(r.Question ?? "") + EscapedBytes(r.Answer ?? "") + EscapedBytes(r.ColumnsJson ?? "") + 200;

        internal static PageRecord LimitPage(PageRecord page)
        {
            var md = page.Markdown ?? string.Empty;
            if (EscapedBytes(md) <= MaxSinglePageBytes) return page;
            long used = 0;
            int cut = 0;
            while (cut < md.Length && used + EscLen(md[cut]) <= MaxSinglePageBytes) { used += EscLen(md[cut]); cut++; }
            if (cut > 0 && char.IsHighSurrogate(md[cut - 1])) cut--;
            var text = md.Substring(0, cut) + "\n[truncated]";
            page.Markdown = text;
            page.CharCount = text.Count(c => !char.IsWhiteSpace(c));
            return page;
        }

        /// <summary>Wraps format-library failures as OdcIngest: errors that never carry document text.</summary>
        private T Guarded<T>(Func<T> work)
        {
            try
            {
                return work();
            }
            catch (InvalidOperationException ex) when (ex.Message.StartsWith("OdcIngest:", StringComparison.Ordinal))
            {
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogError("OdcIngest: failure type={ExceptionType}", ex.GetType().Name);
                throw new InvalidOperationException("OdcIngest: could not read document (" + ex.GetType().Name + ")");
            }
        }
    }
}
