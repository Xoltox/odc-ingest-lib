using System.Diagnostics;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OutSystems.ExternalLibraries.SDK;
using RFPIngestLib.Implementations;
using RFPIngestLib.Interfaces;
using RFPIngestLib.Structures;

namespace RFPIngestLib.UnitTests
{
    public class GeneratedTests
    {
        private static string ShortPdf(int pages) =>
            Fixtures.MakePdf(pages, p => new[] { $"Page marker {p} with some extra words to pass the text layer threshold." });

        // Block 4 test 4: range table (10 rows)
        [Theory]
        [InlineData(0, 5, "throw")]
        [InlineData(-3, 5, "throw")]
        [InlineData(13, 15, "throw")]       // from > count (12 pages)
        [InlineData(5, 4, "throw")]         // to < from
        [InlineData(11, 500, "2:11-12")]    // to > count clamped
        [InlineData(3, 3, "1:3-3")]         // single page
        [InlineData(12, 12, "1:12-12")]     // last page only
        [InlineData(1, 12, "12:1-12")]      // whole document
        [InlineData(1, 1, "1:1-1")]
        [InlineData(2, 2000, "11:2-12")]
        public void Test4_RangeTable(int from, int to, string expect)
        {
            var path = ShortPdf(12);
            try
            {
                var lib = Fixtures.NewLib();
                if (expect == "throw")
                {
                    var ex = Assert.Throws<InvalidOperationException>(() => lib.ExtractPages(Fixtures.Conn, path, from, to));
                    Assert.Equal("RFPIngest: page out of range", ex.Message);
                    return;
                }
                var pages = lib.ExtractPages(Fixtures.Conn, path, from, to).ToList();
                var parts = expect.Split(':');
                Assert.Equal(int.Parse(parts[0]), pages.Count);
                var span = parts[1].Split('-');
                Assert.Equal(int.Parse(span[0]), pages[0].PageNumber);
                Assert.Equal(int.Parse(span[1]), pages[^1].PageNumber);
            }
            finally { File.Delete(path); }
        }

        [Fact]
        public void Test4b_FiftyPageClamp()
        {
            var path = ShortPdf(120);
            try
            {
                var pages = Fixtures.NewLib().ExtractPages(Fixtures.Conn, path, 1, 120).ToList();
                Assert.Equal(50, pages.Count);
                Assert.Equal(50, pages[^1].PageNumber);
                var next = Fixtures.NewLib().ExtractPages(Fixtures.Conn, path, 51, 120).ToList();
                Assert.Equal(51, next[0].PageNumber);
                Assert.Equal(100, next[^1].PageNumber);
            }
            finally { File.Delete(path); }
        }

        // Block 4 test 5
        [Fact]
        public void Test5_OutputBudgetAndContinuation()
        {
            // 60 pages, ~200 KB of text each: 400 lines x 500 chars at font size 1.
            var path = Fixtures.MakePdf(60, p =>
            {
                var lines = new List<string>();
                for (int i = 0; i < 400; i++)
                    lines.Add((i == 0 ? $"PAGE{p} " : "") + string.Concat(Enumerable.Repeat("lorem ipsum dolor ", 28)).Substring(0, 500));
                return lines;
            }, fontSize: 1, lineStep: 2);
            try
            {
                var lib = Fixtures.NewLib();
                var first = lib.ExtractPages(Fixtures.Conn, path, 1, 50).ToList();
                Assert.InRange(first.Count, 1, 49);
                long bytes = first.Sum(p => Encoding.UTF8.GetByteCount(p.Markdown));
                Assert.True(bytes <= 4_000_000, "bytes " + bytes);
                Assert.True(bytes > 150_000 * first.Count, "pages should be large, got " + bytes);

                var second = lib.ExtractPages(Fixtures.Conn, path, first[^1].PageNumber + 1, first[^1].PageNumber + 50).ToList();
                Assert.NotEmpty(second);
                Assert.Equal(first[^1].PageNumber + 1, second[0].PageNumber);
                for (int i = 1; i < second.Count; i++) Assert.Equal(second[i - 1].PageNumber + 1, second[i].PageNumber);
                Assert.True(second.Sum(p => Encoding.UTF8.GetByteCount(p.Markdown)) <= 4_000_000);
            }
            finally { File.Delete(path); }
        }

        // Block 4 test 6
        [Fact]
        public void Test6_Scale400Pages()
        {
            var path = Fixtures.MakePdf(400, p => Enumerable.Range(1, 40).Select(i => $"Page {p} line {i} some descriptive text for the scale test"));
            try
            {
                var lib = Fixtures.NewLib();
                Assert.Equal(400, lib.GetPageCount(Fixtures.Conn, path));
                var sw = Stopwatch.StartNew();
                var pages = lib.ExtractPages(Fixtures.Conn, path, 201, 250).ToList();
                sw.Stop();
                Assert.Equal(50, pages.Count);
                Assert.True(sw.Elapsed < TimeSpan.FromSeconds(20), "took " + sw.Elapsed);
            }
            finally { File.Delete(path); }
        }

        // Block 4 test 7
        [Fact]
        public void Test7_ScannedPageHasNoTextLayer()
        {
            var path = Fixtures.MakeScannedPdf();
            try
            {
                var pages = Fixtures.NewLib().ExtractPages(Fixtures.Conn, path, 1, 2).ToList();
                Assert.True(pages[0].HasTextLayer);
                Assert.False(pages[1].HasTextLayer);
                Assert.Equal(string.Empty, pages[1].Markdown);
                Assert.Equal(0, pages[1].CharCount);
            }
            finally { File.Delete(path); }
        }

        private static string MakeDocx()
        {
            var path = Fixtures.TempFile(".docx");
            using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                var body = new Body();
                main.Document = new Document(body);
                Paragraph P(string text, string? style = null, bool list = false)
                {
                    var pp = new ParagraphProperties();
                    if (style != null) pp.ParagraphStyleId = new ParagraphStyleId { Val = style };
                    if (list) pp.NumberingProperties = new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 });
                    return new Paragraph(pp, new Run(new Text(text)));
                }
                body.Append(P("Scope heading", "Heading1"));
                body.Append(P("First list item", list: true));
                body.Append(P("Second list item", list: true));
                var table = new Table();
                foreach (var row in new[] { new[] { "H1", "H2", "H3" }, new[] { "a|b", "c", "d" } })
                {
                    var tr = new TableRow();
                    foreach (var c in row) tr.Append(new TableCell(new Paragraph(new Run(new Text(c)))));
                    table.Append(tr);
                }
                body.Append(table);
                var sect = new Paragraph(new ParagraphProperties(new SectionProperties()), new Run(new Text("End of section one")));
                body.Append(sect);
                body.Append(P("Second section heading", "Heading2"));
                body.Append(P("Second section body text"));
                main.Document.Save();
            }
            return path;
        }

        // Block 4 test 8
        [Fact]
        public void Test8_DocxHeadingListTableAndSectionPages()
        {
            var path = MakeDocx();
            try
            {
                var lib = Fixtures.NewLib();
                Assert.Equal(2, lib.GetPageCount(Fixtures.Conn, path));
                var pages = lib.ExtractPages(Fixtures.Conn, path, 1, 2).ToList();
                var md = pages[0].Markdown;
                Assert.Contains("# Scope heading", md);
                Assert.Contains("- First list item", md);
                Assert.Contains("|---|---|---|", md);
                Assert.Contains("a\\|b", md);
                Assert.DoesNotContain("Second section", md);
                Assert.Contains("## Second section heading", pages[1].Markdown);
            }
            finally { File.Delete(path); }
        }

        // Block 4 test 9
        [Fact]
        public void Test9a_PdfRenamedDocxIsReadAsPdf()
        {
            var src = Fixtures.DocPath(Fixtures.PackNames[0], "tech-spec");
            var renamed = Fixtures.TempFile(".docx");
            File.Copy(src, renamed);
            try
            {
                var lib = Fixtures.NewLib();
                Assert.Equal(lib.GetPageCount(Fixtures.Conn, src), lib.GetPageCount(Fixtures.Conn, renamed));
                Assert.True(lib.ExtractPages(Fixtures.Conn, renamed, 1, 1).First().HasTextLayer);
            }
            finally { File.Delete(renamed); }
        }

        [Fact]
        public void Test9b_TextFileIsUnsupported()
        {
            var txt = Fixtures.TempFile(".txt");
            File.WriteAllText(txt, "plain text, not a document");
            try
            {
                var ex = Assert.Throws<InvalidOperationException>(() => Fixtures.NewLib().GetPageCount(Fixtures.Conn, txt));
                Assert.Equal("RFPIngest: unsupported format", ex.Message);
            }
            finally { File.Delete(txt); }
        }

        // Block 4 test 10
        [Fact]
        public void Test10_InterfaceUsesOnlySupportedOdcTypes()
        {
            var iface = typeof(IRFPIngestLib);
            Assert.NotNull(iface.GetCustomAttribute<OSInterfaceAttribute>());
            var simple = new HashSet<Type> { typeof(string), typeof(int), typeof(long), typeof(bool), typeof(DateTime), typeof(decimal), typeof(float), typeof(double) };

            bool Supported(Type t, bool allowNullable)
            {
                if (t == typeof(void)) return true;
                if (t.IsByRef) t = t.GetElementType()!;
                var nn = Nullable.GetUnderlyingType(t);
                if (nn != null) return allowNullable && simple.Contains(nn);
                if (simple.Contains(t)) return true;
                if (t.IsValueType && !t.IsPrimitive && !t.IsEnum && t.GetCustomAttribute<OSStructureAttribute>() != null) return true;
                if (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(IEnumerable<>)) return Supported(t.GetGenericArguments()[0], false);
                return false;
            }

            var methods = iface.GetMethods();
            Assert.Equal(new[] { "ExtractPages", "ExtractQuestionnaireRows", "GetPageCount" }, methods.Select(m => m.Name).OrderBy(n => n));
            foreach (var m in methods)
            {
                Assert.False(typeof(Task).IsAssignableFrom(m.ReturnType), m.Name + " returns Task");
                Assert.NotEqual(typeof(byte[]), m.ReturnType);
                Assert.True(Supported(m.ReturnType, false), m.Name + " return type");
                foreach (var p in m.GetParameters())
                {
                    Assert.NotEqual(typeof(byte[]), p.ParameterType);
                    Assert.True(Supported(p.ParameterType, true), m.Name + "." + p.Name);
                }
            }
            foreach (var st in new[] { typeof(StorageConnection), typeof(PageRecord), typeof(QuestionnaireRow) })
                foreach (var prop in st.GetProperties())
                    Assert.True(Supported(prop.PropertyType, false), st.Name + "." + prop.Name);

            var q = iface.GetMethod("ExtractQuestionnaireRows")!.GetParameters().First(p => p.Name == "sheetName");
            Assert.Equal(typeof(string), q.ParameterType);
            Assert.True(q.HasDefaultValue && q.DefaultValue == null);
        }

        [Fact]
        public void LimitPage_TruncatesOversizedPage()
        {
            var big = new PageRecord { PageNumber = 1, Markdown = new string('x', 4_100_000), HasTextLayer = true, CharCount = 4_100_000 };
            var r = RFPIngestLibImpl.LimitPage(big);
            Assert.EndsWith("\n[truncated]", r.Markdown);
            Assert.True(Encoding.UTF8.GetByteCount(r.Markdown) <= 3_900_020);
        }

        [Fact]
        public void TimeBudget_StopsAfterAtLeastOnePage()
        {
            var path = ShortPdf(10);
            try
            {
                var lib = Fixtures.NewLib();
                lib.TimeBudget = TimeSpan.Zero;
                var pages = lib.ExtractPages(Fixtures.Conn, path, 1, 10).ToList();
                Assert.Single(pages);
                Assert.Equal(70, RFPIngestLibImpl.TimeBudgetSeconds);
            }
            finally { File.Delete(path); }
        }

        [Fact]
        public void MissingObject_ThrowsRfpIngestError()
        {
            var ex = Assert.Throws<InvalidOperationException>(() => Fixtures.NewLib().GetPageCount(Fixtures.Conn, Path.Combine(Path.GetTempPath(), "does-not-exist.pdf")));
            Assert.StartsWith("RFPIngest:", ex.Message);
        }
    }
}
