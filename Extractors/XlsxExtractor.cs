using System.Text;
using System.Text.Json;
using ClosedXML.Excel;
using RFPIngestLib.Structures;

namespace RFPIngestLib.Extractors
{
    /// <summary>
    /// Sparse-grid XLSX reader. Only cells that hold content are read (the declared used range is never iterated),
    /// capped at 100,000 rows and 200 columns. Hidden rows and columns are excluded and counted. Merged cells carry
    /// the top-left value across the merged range (ranges over 10,000 cells are not propagated). Formula cells
    /// use the cached value stored in the file (empty when the file stores none).
    /// </summary>
    public sealed class XlsxExtractor : IDocumentExtractor
    {
        public const int MaxSheetMarkdownBytes = 1_000_000;
        public const int HeaderSearchRows = 10;
        public const int RowWindow = 500;
        private const int MaxMergePropagation = 10_000;

        private static readonly string[] RefAliases = { "q-id", "id", "item", "no", "#", "ref" };
        private static readonly string[] SectionAliases = { "section", "category", "group" };
        private static readonly string[] QuestionAliases = { "question", "requirement", "description", "query", "clause" };
        private static readonly string[] AnswerAliases = { "answer", "response", "comment", "vendor response" };

        private static XLWorkbook OpenWorkbook(string path, DocContext ctx)
        {
            ctx.DocumentOpens++;
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var wb = new XLWorkbook(fs);
            ctx.Check();
            return wb;
        }

        private static List<IXLWorksheet> VisibleSheets(XLWorkbook wb) =>
            wb.Worksheets.Where(w => w.Visibility == XLWorksheetVisibility.Visible).ToList();

        public IDocumentSession Open(string path, DocContext ctx) => new Session(OpenWorkbook(path, ctx), ctx);

        private sealed class Session : IDocumentSession
        {
            private readonly XLWorkbook _wb;
            private readonly List<IXLWorksheet> _sheets;
            private readonly DocContext _ctx;
            public int PageCount => _sheets.Count;

            public Session(XLWorkbook wb, DocContext ctx) { _wb = wb; _sheets = VisibleSheets(wb); _ctx = ctx; }

            public IEnumerable<PageRecord> Pages(int fromPage, int toPage)
            {
                for (int n = fromPage; n <= toPage && n <= _sheets.Count; n++)
                {
                    _ctx.Check();
                    yield return RenderSheet(n, _sheets[n - 1], _ctx);
                }
            }

            public void Dispose() => _wb.Dispose();
        }

        // ---- sparse grid ----

        internal sealed class SheetGrid
        {
            public readonly SortedDictionary<int, SortedDictionary<int, string>> Rows = new();
            public readonly HashSet<int> HiddenCols = new();
            public int MinCol, LastCol, LastRow;
            public bool Truncated;
            public int HiddenRowCount, HiddenColCount;

            public string Get(int r, int c) => Rows.TryGetValue(r, out var row) && row.TryGetValue(c, out var v) ? v : string.Empty;
            public List<int> Columns() => MinCol == 0 ? new List<int>() : Enumerable.Range(MinCol, LastCol - MinCol + 1).Where(c => !HiddenCols.Contains(c)).ToList();
        }

        private static string CellString(IXLCell cell)
        {
            try
            {
                if (cell.HasFormula && cell.NeedsRecalculation) return string.Empty; // never evaluate formulas: only values stored in the file are used
                return (cell.GetFormattedString() ?? string.Empty).Trim();
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        internal static SheetGrid LoadGrid(IXLWorksheet ws, DocContext ctx)
        {
            ctx.Check();
            var g = new SheetGrid();
            int n = 0;
            foreach (var cell in ws.CellsUsed(XLCellsUsedOptions.Contents))
            {
                if ((++n & 1023) == 0) ctx.Check();
                int r = cell.Address.RowNumber, c = cell.Address.ColumnNumber;
                if (r > Limits.MaxScanRows || c > Limits.MaxScanColumns) { g.Truncated = true; continue; }
                var s = CellString(cell);
                if (s.Length == 0) continue;
                if (!g.Rows.TryGetValue(r, out var row)) g.Rows[r] = row = new SortedDictionary<int, string>();
                row[c] = s;
            }

            // merged ranges: the top-left value is carried across the range
            foreach (var mr in ws.MergedRanges)
            {
                ctx.Check();
                int r1 = mr.RangeAddress.FirstAddress.RowNumber, r2 = Math.Min(mr.RangeAddress.LastAddress.RowNumber, Limits.MaxScanRows);
                int c1 = mr.RangeAddress.FirstAddress.ColumnNumber, c2 = Math.Min(mr.RangeAddress.LastAddress.ColumnNumber, Limits.MaxScanColumns);
                var v = g.Get(r1, c1);
                if (v.Length == 0 || (long)(r2 - r1 + 1) * (c2 - c1 + 1) > MaxMergePropagation) continue;
                for (int r = r1; r <= r2; r++)
                    for (int c = c1; c <= c2; c++)
                    {
                        if (g.Get(r, c).Length > 0) continue;
                        if (!g.Rows.TryGetValue(r, out var row)) g.Rows[r] = row = new SortedDictionary<int, string>();
                        row[c] = v;
                    }
            }

            // hidden rows and columns: excluded, counted
            foreach (var r in g.Rows.Keys.ToList())
            {
                if ((++n & 255) == 0) ctx.Check();
                if (ws.Row(r).IsHidden) { g.Rows.Remove(r); g.HiddenRowCount++; }
            }
            var usedCols = g.Rows.Values.SelectMany(x => x.Keys).Distinct().ToList();
            foreach (var c in usedCols)
                if (ws.Column(c).IsHidden)
                {
                    g.HiddenCols.Add(c);
                    g.HiddenColCount++;
                }
            if (g.HiddenCols.Count > 0)
                foreach (var row in g.Rows.Values) foreach (var c in g.HiddenCols) row.Remove(c);
            foreach (var r in g.Rows.Where(kv => kv.Value.Count == 0).Select(kv => kv.Key).ToList()) g.Rows.Remove(r);

            // real last non-empty row and column
            if (g.Rows.Count > 0)
            {
                g.LastRow = g.Rows.Keys.Max();
                var cols = g.Rows.Values.SelectMany(x => x.Keys);
                g.MinCol = cols.Min();
                g.LastCol = cols.Max();
            }
            return g;
        }

        private static string Escape(string s) =>
            s.Replace("|", "\\|").Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");

        private static PageRecord RenderSheet(int number, IXLWorksheet ws, DocContext ctx)
        {
            var g = LoadGrid(ws, ctx);
            var sb = new StringBuilder();
            sb.Append("## Sheet: ").Append(ws.Name);
            bool any = g.Rows.Count > 0;
            if (any)
            {
                sb.Append("\n\n");
                var cols = g.Columns();
                long tableBytes = 0;
                int written = 0, idx = 0, total = g.Rows.Count;
                foreach (var kv in g.Rows)
                {
                    if ((idx & 255) == 0) ctx.Check();
                    var cells = cols.Select(c => Escape(kv.Value.TryGetValue(c, out var v) ? v : string.Empty));
                    var line = "| " + string.Join(" | ", cells) + " |";
                    long lineBytes = Encoding.UTF8.GetByteCount(line) + 1;
                    if (written > 0 && (tableBytes + lineBytes > MaxSheetMarkdownBytes || sb.Length > Limits.RenderCapChars))
                    {
                        sb.Append("\n[truncated: ").Append(total - idx).Append(" more rows, use ExtractQuestionnaireRows]");
                        break;
                    }
                    tableBytes += lineBytes;
                    if (written > 0) sb.Append('\n');
                    sb.Append(line);
                    if (written == 0) sb.Append('\n').Append('|').Append(string.Concat(Enumerable.Repeat("---|", cols.Count)));
                    written++;
                    idx++;
                }
            }
            if (g.Truncated) sb.Append("\n[scan limit: rows beyond ").Append(Limits.MaxScanRows).Append(" or columns beyond ").Append(Limits.MaxScanColumns).Append(" were ignored]");
            if (g.HiddenRowCount > 0 || g.HiddenColCount > 0)
                sb.Append("\n[hidden content skipped: ").Append(g.HiddenRowCount).Append(" rows, ").Append(g.HiddenColCount).Append(" columns]");
            return PageUtil.Make(number, sb.ToString(), any);
        }

        // ---- Questionnaire rows ----

        /// <summary>
        /// Lazily yields rows. Window: toRow is clamped to fromRow+499; when that window holds no row but later rows exist,
        /// the window rolls forward to the next row that has content, so an empty result always means the sheet is finished.
        /// More than one visible sheet requires a sheet name (E_SHEET_REQUIRED).
        /// </summary>
        public IEnumerable<QuestionnaireRow> ExtractRows(string path, string? sheetName, int fromRow, int toRow, DocContext ctx)
        {
            using var wb = OpenWorkbook(path, ctx);
            var sheets = VisibleSheets(wb);
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                var match = sheets.Where(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase)).ToList();
                if (match.Count == 0) throw new InvalidOperationException("RFPIngest: sheet not found");
                sheets = match;
            }
            else if (sheets.Count > 1)
            {
                throw new InvalidOperationException("RFPIngest: E_SHEET_REQUIRED workbook has " + sheets.Count + " visible sheets; pass sheetName");
            }
            foreach (var ws in sheets)
                foreach (var row in SheetRows(ws, fromRow, toRow, ctx))
                    yield return row;
        }

        private static IEnumerable<QuestionnaireRow> SheetRows(IXLWorksheet ws, int fromRow, int toRow, DocContext ctx)
        {
            var g = LoadGrid(ws, ctx);
            if (g.Rows.Count == 0) yield break;

            int header = -1;
            foreach (var kv in g.Rows)
            {
                if (kv.Key > HeaderSearchRows) break;
                if (kv.Value.Count >= 2) { header = kv.Key; break; }
            }
            if (header < 0) yield break;

            var colNums = g.Columns();
            var headers = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var c in colNums)
            {
                var h = g.Get(header, c);
                if (h.Length == 0) h = "Col" + XLHelper.GetColumnLetterFromNumber(c);
                var unique = h;
                int k = 2;
                while (!seen.Add(unique)) unique = h + "_" + k++;
                headers.Add(unique);
            }

            int refCol = FindExact(headers, RefAliases);
            int secCol = FindExact(headers, SectionAliases);
            int qCol = FindExact(headers, QuestionAliases);
            int aCol = FindContains(headers, AnswerAliases);

            if (qCol < 0)
            {
                double best = -1;
                for (int i = 0; i < headers.Count; i++)
                {
                    long total = 0; int cnt = 0;
                    foreach (var kv in g.Rows)
                    {
                        if (kv.Key <= header) continue;
                        if (kv.Value.TryGetValue(colNums[i], out var v)) { total += v.Length; cnt++; }
                    }
                    double avg = cnt == 0 ? 0 : (double)total / cnt;
                    if (avg > best) { best = avg; qCol = i; }
                }
            }

            int last = Math.Min(toRow, Limits.MaxScanRows);
            int first = Math.Max(header + 1, fromRow);
            int windowEnd = (int)Math.Min((long)last, (long)fromRow + RowWindow - 1);
            int emitted = 0, scanned = 0;

            foreach (var kv in g.Rows)
            {
                if ((++scanned & 255) == 0) ctx.Check();
                int r = kv.Key;
                if (r < first) continue;
                if (r > last) break;
                if (r > windowEnd)
                {
                    if (emitted > 0) break;
                    windowEnd = (int)Math.Min((long)last, (long)r + RowWindow - 1);
                }
                var vals = new string[headers.Count];
                for (int i = 0; i < headers.Count; i++) vals[i] = g.Get(r, colNums[i]);
                var question = qCol >= 0 ? vals[qCol] : string.Empty;
                if (question.Length == 0) continue;

                var json = new Dictionary<string, string>();
                for (int i = 0; i < headers.Count; i++) json[headers[i]] = vals[i];
                emitted++;
                yield return new QuestionnaireRow
                {
                    SheetName = ws.Name,
                    RowNumber = r,
                    RefId = refCol >= 0 ? vals[refCol] : string.Empty,
                    Section = secCol >= 0 ? vals[secCol] : string.Empty,
                    Question = question,
                    Answer = aCol >= 0 ? vals[aCol] : string.Empty,
                    ColumnsJson = JsonSerializer.Serialize(json),
                    HiddenRowsSkipped = g.HiddenRowCount,
                    HiddenColumnsSkipped = g.HiddenColCount,
                    SheetScanTruncated = g.Truncated
                };
            }
        }

        private static int FindExact(List<string> headers, string[] aliases)
        {
            foreach (var a in aliases)
                for (int i = 0; i < headers.Count; i++)
                    if (string.Equals(headers[i], a, StringComparison.OrdinalIgnoreCase)) return i;
            return -1;
        }

        /// <summary>Answer headers such as "Supplier response" are matched by alias containment, alias priority first.</summary>
        private static int FindContains(List<string> headers, string[] aliases)
        {
            foreach (var a in aliases)
                for (int i = 0; i < headers.Count; i++)
                    if (headers[i].Contains(a, StringComparison.OrdinalIgnoreCase)) return i;
            return -1;
        }
    }
}
