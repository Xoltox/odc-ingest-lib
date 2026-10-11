using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OdcIngest.Structures;

namespace OdcIngest.Extractors
{
    /// <summary>
    /// Body content only: headers, footers, footnotes and endnotes are NOT extracted (documented exclusion).
    /// Nested tables are flattened into their parent cell once (direct-child paragraphs, recursion for tables).
    /// </summary>
    public sealed class DocxExtractor : IDocumentExtractor
    {
        public const int CharsPerPage = 3000;

        public IDocumentSession Open(string path, DocContext ctx)
        {
            ctx.DocumentOpens++;
            var doc = WordprocessingDocument.Open(path, false);
            try
            {
                var pages = Paginate(doc, ctx);
                return new Session(doc, pages, ctx);
            }
            catch
            {
                doc.Dispose();
                throw;
            }
        }

        private sealed class Session : IDocumentSession
        {
            private readonly WordprocessingDocument _doc;
            private readonly List<List<OpenXmlElement>> _pages;
            private readonly DocContext _ctx;
            public int PageCount => _pages.Count;

            public Session(WordprocessingDocument doc, List<List<OpenXmlElement>> pages, DocContext ctx) { _doc = doc; _pages = pages; _ctx = ctx; }

            public IEnumerable<PageRecord> Pages(int fromPage, int toPage)
            {
                for (int n = fromPage; n <= toPage && n <= _pages.Count; n++)
                {
                    _ctx.Check();
                    yield return Render(n, _pages[n - 1]);
                }
            }

            public void Dispose() => _doc.Dispose();
        }

        /// <summary>The single pagination function behind the page count and the page extraction.</summary>
        internal static List<List<OpenXmlElement>> Paginate(WordprocessingDocument doc, DocContext ctx)
        {
            var body = doc.MainDocumentPart?.Document?.Body;
            var pages = new List<List<OpenXmlElement>>();
            if (body == null) { pages.Add(new List<OpenXmlElement>()); return pages; }

            bool explicitBreaks = body.Descendants<LastRenderedPageBreak>().Any()
                                  || body.Descendants<Break>().Any(b => b.Type != null && b.Type.Value == BreakValues.Page);

            var cur = new List<OpenXmlElement>();
            int chars = 0;
            void Close()
            {
                if (cur.Any(e => e.InnerText.Trim().Length > 0)) pages.Add(cur);
                cur = new List<OpenXmlElement>();
                chars = 0;
            }

            int seen = 0;
            foreach (var el in body.ChildElements)
            {
                if ((++seen & 255) == 0) ctx.Check();
                if (el is Paragraph p)
                {
                    if (explicitBreaks && p.Descendants<LastRenderedPageBreak>().Any()) Close();
                    cur.Add(p);
                    chars += p.InnerText.Length;
                    bool pageBreak = explicitBreaks && p.Descendants<Break>().Any(b => b.Type != null && b.Type.Value == BreakValues.Page);
                    bool sectionBreak = p.ParagraphProperties?.SectionProperties != null;
                    if (pageBreak || sectionBreak || (!explicitBreaks && chars >= CharsPerPage)) Close();
                }
                else if (el is Table t)
                {
                    cur.Add(t);
                    chars += t.InnerText.Length;
                    if (!explicitBreaks && chars >= CharsPerPage) Close();
                }
            }
            Close();
            if (pages.Count == 0) pages.Add(new List<OpenXmlElement>());
            return pages;
        }

        private static PageRecord Render(int number, List<OpenXmlElement> elements)
        {
            var sb = new StringBuilder();
            bool prevList = false;
            bool any = false;
            foreach (var el in elements)
            {
                if (sb.Length > Limits.RenderCapChars) { sb.Append(Limits.TruncatedLine); break; }
                if (el is Paragraph p)
                {
                    var text = ParagraphText(p);
                    if (text.Length == 0) continue;
                    any = true;
                    var style = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? string.Empty;
                    string line;
                    bool isList = false;
                    int level = HeadingLevel(style);
                    if (level > 0) line = new string('#', level) + " " + text;
                    else if (string.Equals(style, "Title", StringComparison.OrdinalIgnoreCase)) line = "# " + text;
                    else if (p.ParagraphProperties?.NumberingProperties != null || style.StartsWith("List", StringComparison.OrdinalIgnoreCase))
                    {
                        line = "- " + text;
                        isList = true;
                    }
                    else line = text;

                    if (sb.Length > 0) sb.Append(isList && prevList ? "\n" : "\n\n");
                    sb.Append(line);
                    prevList = isList;
                }
                else if (el is Table t)
                {
                    var md = RenderTable(t, out bool hasCell);
                    if (md.Length == 0) continue;
                    if (hasCell) any = true;
                    if (sb.Length > 0) sb.Append("\n\n");
                    sb.Append(md);
                    prevList = false;
                }
            }
            return PageUtil.Make(number, sb.ToString(), any);
        }

        private static int HeadingLevel(string style)
        {
            if (style.Length == 8 && style.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) && style[7] >= '1' && style[7] <= '6')
                return style[7] - '0';
            return 0;
        }

        private static string ParagraphText(Paragraph p)
        {
            var sb = new StringBuilder();
            foreach (var d in p.Descendants())
            {
                switch (d)
                {
                    case Text t: sb.Append(t.Text); break;
                    case TabChar: sb.Append(' '); break;
                    case Break b when b.Type == null || b.Type.Value != BreakValues.Page: sb.Append(' '); break;
                    case CarriageReturn: sb.Append(' '); break;
                }
            }
            return sb.ToString().Trim();
        }

        /// <summary>Direct-child paragraphs, then nested tables recursively (each nested table flattened once).</summary>
        private static string CellText(TableCell c)
        {
            var parts = new List<string>();
            foreach (var child in c.ChildElements)
            {
                if (child is Paragraph p)
                {
                    var s = ParagraphText(p);
                    if (s.Length > 0) parts.Add(s);
                }
                else if (child is Table nested)
                {
                    var inner = nested.Elements<TableRow>().SelectMany(r => r.Elements<TableCell>()).Select(CellText).Where(s => s.Length > 0);
                    var joined = string.Join("; ", inner);
                    if (joined.Length > 0) parts.Add(joined);
                }
            }
            return string.Join(" ", parts).Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
        }

        private static string RenderTable(Table t, out bool hasCell)
        {
            hasCell = false;
            var sb = new StringBuilder();
            bool first = true;
            foreach (var row in t.Elements<TableRow>())
            {
                if (sb.Length > Limits.RenderCapChars) { sb.Append(Limits.TruncatedLine); break; }
                var cells = row.Elements<TableCell>().Select(CellText).ToList();
                if (cells.Count == 0) continue;
                if (cells.Any(c => c.Length > 0)) hasCell = true;
                if (sb.Length > 0) sb.Append('\n');
                sb.Append("| ").Append(string.Join(" | ", cells)).Append(" |");
                if (first)
                {
                    sb.Append('\n').Append("|").Append(string.Concat(Enumerable.Repeat("---|", cells.Count)));
                    first = false;
                }
            }
            return sb.ToString();
        }
    }
}
