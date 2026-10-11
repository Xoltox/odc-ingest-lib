using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;
using OdcIngest.Implementations;
using OdcIngest.Storage;
using OdcIngest.Structures;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Fonts.Standard14Fonts;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Writer;

namespace OdcIngest.UnitTests
{
    public sealed record KeyEntry(string Ref, string SourceDoc, int? Page, string Location, string? ChangeType, string Text);

    public sealed class ListLogger : ILogger
    {
        public List<string> Lines { get; } = new();
        public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;
        public bool IsEnabled(LogLevel logLevel) => true;
        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception, Func<TState, Exception?, string> formatter)
            => Lines.Add(formatter(state, exception));
    }

    public static class Fixtures
    {
        public static readonly string[] PackNames = { "pack-a-arvant", "pack-b-kessler", "pack-c-arvant-amended" };

        public static string RepoRoot { get; } = FindRoot();

        private static string FindRoot()
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null)
            {
                if (Directory.Exists(Path.Combine(dir.FullName, "tests", "fixtures", "rfq-packs"))) return dir.FullName;
                dir = dir.Parent;
            }
            throw new InvalidOperationException("tests/fixtures/rfq-packs not found above " + AppContext.BaseDirectory);
        }

        public static string PackDir(string pack) => Path.Combine(RepoRoot, "tests", "fixtures", "rfq-packs", pack);

        public static string DocPath(string pack, string sourceDoc)
        {
            var file = sourceDoc switch
            {
                "tech-spec" => "tech-spec.pdf",
                "commercial-terms" => "commercial-terms.docx",
                "questionnaire" => "questionnaire.xlsx",
                _ => throw new ArgumentException(sourceDoc)
            };
            return Path.Combine(PackDir(pack), file);
        }

        /// <summary>Key entries for a source doc; entries with change_type "removed" are skipped.</summary>
        public static List<KeyEntry> Key(string pack, string sourceDoc)
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(Path.Combine(PackDir(pack), "answer-key.json")));
            var list = new List<KeyEntry>();
            foreach (var r in doc.RootElement.GetProperty("requirements").EnumerateArray())
            {
                if (r.GetProperty("source_doc").GetString() != sourceDoc) continue;
                string? ct = r.TryGetProperty("change_type", out var c) ? c.GetString() : null;
                if (ct == "removed") continue;
                int? page = r.GetProperty("page").ValueKind == JsonValueKind.Number ? r.GetProperty("page").GetInt32() : null;
                list.Add(new KeyEntry(r.GetProperty("ref").GetString()!, sourceDoc, page, r.GetProperty("location").GetString()!, ct, r.GetProperty("text").GetString()!));
            }
            return list;
        }

        public static OdcIngestLibImpl NewLib(ILogger? logger = null) =>
            new(logger ?? new ListLogger(), _ => new FileObjectSource());

        public static StorageConnection Conn => default;

        /// <summary>Whole-document batched extraction following the caller loop from the spec.</summary>
        public static List<PageRecord> ExtractAll(OdcIngestLibImpl lib, string path)
        {
            int count = lib.GetPageCount(Conn, path);
            var all = new List<PageRecord>();
            int from = 1;
            while (from <= count)
            {
                var batch = lib.ExtractPages(Conn, path, from, from + 49).ToList();
                Assert.NotEmpty(batch);
                all.AddRange(batch);
                from = batch[^1].PageNumber + 1;
            }
            return all;
        }

        public static Regex RefRegex(string reference, bool allowPipe) =>
            new("(?m)^[" + (allowPipe ? "|" : "") + @"#\s*-]*" + Regex.Escape(reference) + @"(?:[\s:]|\.(?!\d))");

        // ---- generated PDFs ----

        public static string TempFile(string ext) => Path.Combine(Path.GetTempPath(), "odcing-ut-" + Guid.NewGuid().ToString("N") + ext);

        public static string MakePdf(int pages, Func<int, IEnumerable<string>> linesOfPage, double fontSize = 10, double lineStep = 12)
        {
            var b = new PdfDocumentBuilder();
            var font = b.AddStandard14Font(Standard14Font.Helvetica);
            for (int p = 1; p <= pages; p++)
            {
                var page = b.AddPage(PageSize.A4);
                double y = 820;
                foreach (var line in linesOfPage(p))
                {
                    page.AddText(line, fontSize, new PdfPoint(10, y), font);
                    y -= lineStep;
                    if (y < 5) break;
                }
            }
            var path = TempFile(".pdf");
            File.WriteAllBytes(path, b.Build());
            return path;
        }

        public static string MakeScannedPdf()
        {
            var b = new PdfDocumentBuilder();
            var font = b.AddStandard14Font(Standard14Font.Helvetica);
            var p1 = b.AddPage(PageSize.A4);
            p1.AddText("Page one carries a real text layer with enough characters.", 10, new PdfPoint(20, 800), font);
            var p2 = b.AddPage(PageSize.A4);
            p2.DrawRectangle(new PdfPoint(50, 50), 200, 200, 2);
            var path = TempFile(".pdf");
            File.WriteAllBytes(path, b.Build());
            return path;
        }
    }
}
