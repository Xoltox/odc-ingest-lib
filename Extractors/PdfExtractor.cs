using System.Text;
using System.Text.RegularExpressions;
using RFPIngestLib.Structures;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Exceptions;

namespace RFPIngestLib.Extractors
{
    public sealed class PdfExtractor : IDocumentExtractor
    {
        private static readonly Regex ClauseHeading = new(@"^\d+(\.\d+)+\s+\p{Lu}", RegexOptions.Compiled);

        public IDocumentSession Open(string path, DocContext ctx)
        {
            try
            {
                ctx.DocumentOpens++;
                var doc = ctx.Run(() => PdfDocument.Open(path));
                int count;
                try { count = doc.NumberOfPages; }
                catch (Exception) { doc.Dispose(); throw new InvalidOperationException("RFPIngest: E_PDF_CORRUPT unreadable page tree"); }
                return new Session(doc, count, ctx);
            }
            catch (InvalidOperationException ex) when (ex.Message.StartsWith("RFPIngest:", StringComparison.Ordinal))
            {
                throw;
            }
            catch (PdfDocumentEncryptedException)
            {
                throw new InvalidOperationException("RFPIngest: E_PDF_ENCRYPTED password-protected PDF is not supported");
            }
            catch (Exception)
            {
                throw new InvalidOperationException("RFPIngest: E_PDF_CORRUPT unreadable PDF");
            }
        }

        private sealed class Session : IDocumentSession
        {
            private readonly PdfDocument _doc;
            private readonly DocContext _ctx;
            public int PageCount { get; }

            public Session(PdfDocument doc, int count, DocContext ctx) { _doc = doc; PageCount = count; _ctx = ctx; }

            public IEnumerable<PageRecord> Pages(int fromPage, int toPage)
            {
                for (int n = fromPage; n <= toPage && n <= PageCount; n++)
                {
                    _ctx.Check();
                    yield return Build(n);
                }
            }

            private PageRecord Build(int n)
            {
                try
                {
                    return _ctx.Run(() => BuildPage(n, _doc.GetPage(n)));
                }
                catch (InvalidOperationException ex) when (ex.Message.StartsWith("RFPIngest:", StringComparison.Ordinal))
                {
                    throw;
                }
                catch (Exception)
                {
                    throw new InvalidOperationException("RFPIngest: E_PDF_CORRUPT unreadable page");
                }
            }

            public void Dispose() => _doc.Dispose();
        }

        private sealed class Line
        {
            public StringBuilder Text = new();
            public double MaxSize;
        }

        private static PageRecord BuildPage(int number, Page page)
        {
            var lines = new List<Line>();
            var sizes = new List<double>();

            Line? cur = null;
            double lineY = 0;
            Letter? prev = null;
            long chars = 0;
            foreach (var l in page.Letters)
            {
                var v = l.Value;
                if (string.IsNullOrEmpty(v)) continue;
                if (chars > Limits.RenderCapChars) break;
                double size = l.PointSize > 0 ? l.PointSize : 1;
                double y = l.StartBaseLine.Y;
                bool isSpace = v.Trim().Length == 0;

                if (cur == null || Math.Abs(y - lineY) > 0.5 * size)
                {
                    if (isSpace) { prev = l; continue; }
                    cur = new Line();
                    lines.Add(cur);
                    lineY = y;
                }
                else if (!isSpace && prev != null && cur.Text.Length > 0)
                {
                    double gap = l.StartBaseLine.X - prev.EndBaseLine.X;
                    char last = cur.Text[cur.Text.Length - 1];
                    if (last != ' ' && gap > 0.2 * size) cur.Text.Append(' ');
                }

                if (isSpace)
                {
                    if (cur.Text.Length > 0 && cur.Text[cur.Text.Length - 1] != ' ') cur.Text.Append(' ');
                }
                else
                {
                    cur.Text.Append(v);
                    chars += v.Length;
                    if (size > cur.MaxSize) cur.MaxSize = size;
                    sizes.Add(size);
                }
                prev = l;
            }

            double median = 0;
            if (sizes.Count > 0)
            {
                sizes.Sort();
                median = sizes[sizes.Count / 2];
            }

            var rendered = new List<(string text, double size)>();
            foreach (var ln in lines)
            {
                var t = ln.Text.ToString().Trim();
                if (t.Length == 0) continue;
                if (rendered.Count > 0)
                {
                    var (pt, ps) = rendered[rendered.Count - 1];
                    if (pt.Length >= 2 && pt[^1] == '-' && char.IsLetter(pt[^2]) && char.IsLower(t[0]))
                    {
                        rendered[rendered.Count - 1] = (pt.Substring(0, pt.Length - 1) + t, ps);
                        continue;
                    }
                }
                rendered.Add((t, ln.MaxSize));
            }

            var sb = new StringBuilder();
            int alnum = 0;
            foreach (var (text, size) in rendered)
            {
                if (sb.Length > Limits.RenderCapChars) { sb.Append(Limits.TruncatedLine); break; }
                bool heading = (text.Length < 120 && median > 0 && size >= 1.25 * median) || ClauseHeading.IsMatch(text);
                if (heading) sb.Append("## ");
                sb.Append(text).Append('\n');
                foreach (var ch in text) if (char.IsLetterOrDigit(ch)) alnum++;
            }

            bool hasText = alnum >= 20;
            return PageUtil.Make(number, sb.ToString().TrimEnd('\n'), hasText);
        }
    }
}
