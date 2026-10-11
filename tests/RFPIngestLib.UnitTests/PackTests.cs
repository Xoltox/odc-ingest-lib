using System.Text;
using System.Text.Json;
using ClosedXML.Excel;
using RFPIngestLib.Structures;

namespace RFPIngestLib.UnitTests
{
    public class PackTests
    {
        private readonly ITestOutputHelper _out;
        public PackTests(ITestOutputHelper output) { _out = output; }

        public static IEnumerable<TheoryDataRow<string>> Packs() => Fixtures.PackNames.Select(p => new TheoryDataRow<string>(p));

        // Block 4 test 1
        [Theory]
        [MemberData(nameof(Packs))]
        public void Test1_EveryKeyRefAppearsAtLineStart(string pack)
        {
            var lib = Fixtures.NewLib();
            int hits = 0, total = 0;
            var misses = new List<string>();
            foreach (var sourceDoc in new[] { "tech-spec", "commercial-terms", "questionnaire" })
            {
                var path = Fixtures.DocPath(pack, sourceDoc);
                var pages = Fixtures.ExtractAll(lib, path);
                var rx = sourceDoc == "questionnaire" ? "pipe" : "plain";
                int dh = 0;
                var key = Fixtures.Key(pack, sourceDoc);
                foreach (var k in key)
                {
                    total++;
                    var regex = Fixtures.RefRegex(k.Ref, allowPipe: sourceDoc == "questionnaire");
                    var found = pages.Where(p => regex.IsMatch(p.Markdown)).Select(p => p.PageNumber).ToList();
                    if (found.Count == 0) { misses.Add($"{sourceDoc}:{k.Ref}"); continue; }
                    if (sourceDoc == "tech-spec" && k.Page.HasValue && !found.Contains(k.Page.Value))
                    {
                        misses.Add($"{sourceDoc}:{k.Ref} wrong page (key {k.Page}, found {string.Join(",", found)})");
                        continue;
                    }
                    hits++; dh++;
                }
                _out.WriteLine($"{pack} {sourceDoc}: {dh}/{key.Count} ({rx})");
            }
            _out.WriteLine($"{pack}: {hits}/{total}");
            Assert.True(misses.Count == 0, "Missing: " + string.Join("; ", misses));
        }

        // Block 4 test 2
        [Theory]
        [MemberData(nameof(Packs))]
        public void Test2_QuestionnaireRowsMatchKey(string pack)
        {
            var path = Fixtures.DocPath(pack, "questionnaire");
            var lib = Fixtures.NewLib();
            int visible;
            using (var wb = new XLWorkbook(path)) visible = wb.Worksheets.Count(w => w.Visibility == XLWorksheetVisibility.Visible);
            Assert.Equal(visible, lib.GetPageCount(Fixtures.Conn, path));

            var key = Fixtures.Key(pack, "questionnaire");
            var rows = lib.ExtractQuestionnaireRows(Fixtures.Conn, path, null, 1, 500).ToList();
            Assert.Equal(key.Count, rows.Count);
            Assert.All(rows, r => Assert.False(string.IsNullOrWhiteSpace(r.Question)));
            for (int i = 1; i < rows.Count; i++) Assert.True(rows[i].RowNumber > rows[i - 1].RowNumber);
            Assert.DoesNotContain(rows, r => r.RowNumber == 1);

            using var wbRef = new XLWorkbook(path);
            var wsRef = wbRef.Worksheet(1);
            var byRef = rows.ToDictionary(r => r.RefId);
            int textMatches = 0;
            foreach (var k in key)
            {
                Assert.True(byRef.ContainsKey(k.Ref), "missing " + k.Ref);
                var row = byRef[k.Ref];
                using var cols = JsonDocument.Parse(row.ColumnsJson);
                Assert.True(cols.RootElement.TryGetProperty("Clause ref", out var cr));
                // Some general questions (Q-24..Q-26) legitimately carry no clause ref.
                var expected = wsRef.Cell(row.RowNumber, 2).GetString().Trim();
                Assert.Equal(expected, cr.GetString());
                if (expected.Length > 0) Assert.Matches(@"^\d+(\.\d+)*$", expected);
                // location is Questions!C<row>
                Assert.Equal(int.Parse(k.Location.Split('!')[1].Substring(1)), row.RowNumber);
                if (row.Question == k.Text) textMatches++;
            }
            _out.WriteLine($"{pack}: rows {rows.Count}/{key.Count}, question text equals key {textMatches}/{key.Count}");
            Assert.Equal(key.Count, textMatches);
        }

        // Block 4 test 3
        [Theory]
        [MemberData(nameof(Packs))]
        public void Test3_PageCountAgreesWithExtraction(string pack)
        {
            var lib = Fixtures.NewLib();
            foreach (var sourceDoc in new[] { "tech-spec", "commercial-terms", "questionnaire" })
            {
                var path = Fixtures.DocPath(pack, sourceDoc);
                var all = Fixtures.ExtractAll(lib, path);
                Assert.Equal(lib.GetPageCount(Fixtures.Conn, path), all.Max(p => p.PageNumber));
                Assert.Equal(Enumerable.Range(1, all.Count), all.Select(p => p.PageNumber));
                _out.WriteLine($"{pack} {sourceDoc}: {all.Count} pages");
            }
        }

        [Fact]
        public void Logs_NeverCarryDocumentText()
        {
            var log = new ListLogger();
            var lib = Fixtures.NewLib(log);
            var pack = Fixtures.PackNames[0];
            foreach (var d in new[] { "tech-spec", "commercial-terms", "questionnaire" }) Fixtures.ExtractAll(lib, Fixtures.DocPath(pack, d));
            lib.ExtractQuestionnaireRows(Fixtures.Conn, Fixtures.DocPath(pack, "questionnaire")).ToList();
            Assert.NotEmpty(log.Lines);
            var all = string.Join("\n", log.Lines);
            foreach (var d in new[] { "tech-spec", "commercial-terms", "questionnaire" })
                foreach (var k in Fixtures.Key(pack, d).Take(10))
                    Assert.DoesNotContain(k.Text.Substring(0, 25), all);
            Assert.DoesNotContain("CT-1 ", all);
        }
    }
}
