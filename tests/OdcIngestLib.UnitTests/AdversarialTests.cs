using System.IO.Compression;
using System.Text;
using System.Text.Json;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OdcIngest.Extractors;
using OdcIngest.Implementations;
using OdcIngest.Storage;
using OdcIngest.Structures;

namespace OdcIngest.UnitTests
{
    /// <summary>All fixtures are generated in-test and deleted afterwards; nothing large is committed.</summary>
    public class AdversarialTests
    {
        private static InvalidOperationException Fails(Action a)
        {
            var ex = Assert.Throws<InvalidOperationException>(a);
            Assert.StartsWith("OdcIngest:", ex.Message);
            return ex;
        }

        // ---- 1. zip bomb ----

        private static string MakeZip(params (string name, long size)[] entries)
        {
            var path = Fixtures.TempFile(".docx");
            using (var fs = File.Create(path))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                var chunk = new byte[1024 * 1024];
                foreach (var (name, size) in entries)
                {
                    using var s = zip.CreateEntry(name, CompressionLevel.Fastest).Open();
                    for (long w = 0; w < size; w += chunk.Length) s.Write(chunk, 0, (int)Math.Min(chunk.Length, size - w));
                }
            }
            return path;
        }

        [Fact]
        public void ZipBomb_SingleEntryOver100MB()
        {
            var p = MakeZip(("word/document.xml", 101L * 1024 * 1024));
            try { Assert.Contains("E_ZIP_BOMB entry", Fails(() => Fixtures.NewLib().GetPageCount(Fixtures.Conn, p)).Message); }
            finally { File.Delete(p); }
        }

        [Fact]
        public void ZipBomb_TotalOver200MB()
        {
            var p = MakeZip(("word/document.xml", 90L << 20), ("word/a.bin", 90L << 20), ("word/b.bin", 30L << 20));
            try { Assert.Contains("E_ZIP_BOMB total", Fails(() => Fixtures.NewLib().GetPageCount(Fixtures.Conn, p)).Message); }
            finally { File.Delete(p); }
        }

        [Fact]
        public void ZipBomb_RatioOver100To1()
        {
            var p = MakeZip(("word/document.xml", 30L << 20));
            try { Assert.Contains("E_ZIP_BOMB compression ratio", Fails(() => Fixtures.NewLib().GetPageCount(Fixtures.Conn, p)).Message); }
            finally { File.Delete(p); }
        }

        // ---- 2. hard deadline, scan caps, trim to real last cell ----

        [Fact]
        public void HardDeadline_StopsPdfAndXlsx()
        {
            var lib = Fixtures.NewLib();
            lib.HardTimeBudget = TimeSpan.Zero;
            Assert.Contains("E_TIMEOUT", Fails(() => lib.ExtractPages(Fixtures.Conn, Fixtures.DocPath(Fixtures.PackNames[0], "tech-spec"), 1, 3)).Message);
            Assert.Contains("E_TIMEOUT", Fails(() => lib.ExtractQuestionnaireRows(Fixtures.Conn, Fixtures.DocPath(Fixtures.PackNames[0], "questionnaire")).ToList()).Message);
            Assert.Contains("E_TIMEOUT", Fails(() => lib.ExtractPages(Fixtures.Conn, Fixtures.DocPath(Fixtures.PackNames[0], "commercial-terms"), 1, 1)).Message);
        }

        [Fact]
        public void DocContext_Run_AbandonsWorkPastDeadline()
        {
            var ctx = new DocContext(TimeSpan.FromMilliseconds(150));
            Assert.Contains("E_TIMEOUT", Fails(() => ctx.Run(() => { Thread.Sleep(3000); return 1; })).Message);
            Assert.Equal(7, new DocContext(TimeSpan.FromSeconds(5)).Run(() => 7));
        }

        [Fact]
        public void HugeUsedRange_IsCappedAndNotIterated()
        {
            var p = Fixtures.TempFile(".xlsx");
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Questions");
                ws.Cell("A1").Value = "Q-ID"; ws.Cell("B1").Value = "Question";
                ws.Cell("A2").Value = "Q-1"; ws.Cell("B2").Value = "First question";
                ws.Cell(2, 300).Value = "beyond column cap";
                ws.Cell(1_000_000, 1).Value = "far away";
                wb.SaveAs(p);
            }
            try
            {
                var lib = Fixtures.NewLib();
                var sw = System.Diagnostics.Stopwatch.StartNew();
                var rows = lib.ExtractQuestionnaireRows(Fixtures.Conn, p).ToList();
                var page = lib.ExtractPages(Fixtures.Conn, p, 1, 1).Single();
                sw.Stop();
                Assert.True(sw.Elapsed < TimeSpan.FromSeconds(20), sw.Elapsed.ToString());
                Assert.Single(rows);
                Assert.Equal(2, rows[0].RowNumber);
                Assert.True(rows[0].SheetScanTruncated);
                Assert.Contains("[scan limit", page.Markdown);
                Assert.DoesNotContain("far away", page.Markdown);
                Assert.DoesNotContain("beyond column cap", page.Markdown);
            }
            finally { File.Delete(p); }
        }

        // ---- 3. row windows and sheet selection ----

        private static string MakeRowsBook(int lastRow, int[]? onlyRows = null, string[]? extraSheets = null)
        {
            var p = Fixtures.TempFile(".xlsx");
            using var wb = new XLWorkbook();
            foreach (var name in new[] { "Questions" }.Concat(extraSheets ?? Array.Empty<string>()))
            {
                var ws = wb.AddWorksheet(name);
                ws.Cell("A1").Value = "Q-ID"; ws.Cell("B1").Value = "Question";
                foreach (var r in onlyRows ?? Enumerable.Range(2, lastRow - 1).ToArray())
                {
                    ws.Cell(r, 1).Value = "Q-" + r; ws.Cell(r, 2).Value = "Question " + r;
                }
            }
            wb.SaveAs(p);
            return p;
        }

        [Fact]
        public void RowWindow_ClampedTo500_AndResumable()
        {
            var p = MakeRowsBook(1200);
            try
            {
                var lib = Fixtures.NewLib();
                var all = new List<QuestionnaireRow>();
                int from = 1;
                for (int guard = 0; guard < 10; guard++)
                {
                    var batch = lib.ExtractQuestionnaireRows(Fixtures.Conn, p, null, from, 5000).ToList();
                    if (batch.Count == 0) break;
                    Assert.InRange(batch.Count, 1, 500);
                    Assert.True(batch[^1].RowNumber <= from + 499);
                    all.AddRange(batch);
                    from = batch[^1].RowNumber + 1;
                }
                Assert.Equal(1199, all.Count);
                Assert.Equal(all.Select(r => r.RowNumber).Distinct().Count(), all.Count);
            }
            finally { File.Delete(p); }
        }

        [Fact]
        public void RowWindow_RollsForwardOverEmptyGap()
        {
            var p = MakeRowsBook(0, new[] { 2, 3, 1200 });
            try
            {
                var rows = Fixtures.NewLib().ExtractQuestionnaireRows(Fixtures.Conn, p, null, 11, 5000).ToList();
                Assert.Single(rows);
                Assert.Equal(1200, rows[0].RowNumber);
                Assert.Empty(Fixtures.NewLib().ExtractQuestionnaireRows(Fixtures.Conn, p, null, 1201, 5000));
            }
            finally { File.Delete(p); }
        }

        [Fact]
        public void MultipleVisibleSheets_RequireSheetName()
        {
            var p = MakeRowsBook(5, null, new[] { "Second" });
            try
            {
                var lib = Fixtures.NewLib();
                Assert.Contains("E_SHEET_REQUIRED", Fails(() => lib.ExtractQuestionnaireRows(Fixtures.Conn, p).ToList()).Message);
                var rows = lib.ExtractQuestionnaireRows(Fixtures.Conn, p, "Second").ToList();
                Assert.Equal(4, rows.Count);
                Assert.All(rows, r => Assert.Equal("Second", r.SheetName));
            }
            finally { File.Delete(p); }
        }

        // ---- 4. output size counted as JSON-escaped ----

        [Fact]
        public void JsonExpansion_BudgetIsEscapedSize()
        {
            // 200 KB of text per page, half of it double quotes (6 bytes each once JSON-escaped).
            var path = Fixtures.MakePdf(20, p => Enumerable.Range(0, 400).Select(_ => string.Concat(Enumerable.Repeat("a\"", 250))), fontSize: 1, lineStep: 2);
            try
            {
                var lib = Fixtures.NewLib();
                var pages = lib.ExtractPages(Fixtures.Conn, path, 1, 20).ToList();
                Assert.InRange(pages.Count, 1, 7);
                Assert.True(pages.Sum(x => OdcIngestLibImpl.EscapedBytes(x.Markdown)) <= 4_000_000);
                var json = JsonSerializer.Serialize(pages);
                Assert.True(json.Length < 5_500_000, "json " + json.Length);
            }
            finally { File.Delete(path); }
        }

        [Fact]
        public void RenderCap_IsAppliedWhileBuilding()
        {
            // one line of 6 MB of text must not be rendered in full
            var big = new PageRecord { PageNumber = 1, Markdown = new string('é', 1_000_000), HasTextLayer = true };
            var r = OdcIngestLibImpl.LimitPage(big);
            Assert.True(OdcIngestLibImpl.EscapedBytes(r.Markdown) <= 3_900_020);
            Assert.Equal(3_900_000 / 6, r.Markdown.Length - "\n[truncated]".Length);
        }

        // ---- 5. hidden content, merged cells, formulas ----

        [Fact]
        public void HiddenRowsAndColumns_AreSkippedAndCounted()
        {
            var p = Fixtures.TempFile(".xlsx");
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Questions");
                ws.Cell("A1").Value = "Q-ID"; ws.Cell("B1").Value = "Question"; ws.Cell("D1").Value = "Notes";
                for (int r = 2; r <= 4; r++) { ws.Cell(r, 1).Value = "Q-" + r; ws.Cell(r, 2).Value = "Question " + r; ws.Cell(r, 4).Value = "secret note"; }
                ws.Row(3).Hide();
                ws.Column(4).Hide();
                wb.SaveAs(p);
            }
            try
            {
                var lib = Fixtures.NewLib();
                var rows = lib.ExtractQuestionnaireRows(Fixtures.Conn, p).ToList();
                Assert.Equal(new[] { 2, 4 }, rows.Select(r => r.RowNumber));
                Assert.All(rows, r => { Assert.Equal(1, r.HiddenRowsSkipped); Assert.Equal(1, r.HiddenColumnsSkipped); Assert.DoesNotContain("Notes", r.ColumnsJson); });
                var page = lib.ExtractPages(Fixtures.Conn, p, 1, 1).Single();
                Assert.DoesNotContain("secret note", page.Markdown);
                Assert.Contains("[hidden content skipped: 1 rows, 1 columns]", page.Markdown);
            }
            finally { File.Delete(p); }
        }

        [Fact]
        public void MergedCells_PropagateTopLeft_AndFormulasUseCachedValue()
        {
            var p = Fixtures.TempFile(".xlsx");
            using (var fs = File.Create(p))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                void Add(string name, string text) { using var w = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false)); w.Write(text); }
                const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                Add("[Content_Types].xml", "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>");
                Add("_rels/.rels", "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
                Add("xl/workbook.xml", "<?xml version=\"1.0\"?><workbook xmlns=\"" + ns + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Questions\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                Add("xl/_rels/workbook.xml.rels", "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
                string S(string r, string t) => "<c r=\"" + r + "\" t=\"inlineStr\"><is><t>" + t + "</t></is></c>";
                Add("xl/worksheets/sheet1.xml", "<?xml version=\"1.0\"?><worksheet xmlns=\"" + ns + "\"><sheetData>"
                    + "<row r=\"1\">" + S("A1", "Section") + S("B1", "Q-ID") + S("C1", "Question") + S("D1", "Answer") + "</row>"
                    + "<row r=\"2\">" + S("A2", "SecA") + S("B2", "Q-1") + S("C2", "First") + "<c r=\"D2\"><f>1+2</f><v>3</v></c></row>"
                    + "<row r=\"3\">" + S("B3", "Q-2") + S("C3", "Second") + "<c r=\"D3\"><f>1+2</f></c></row>"
                    + "</sheetData><mergeCells count=\"1\"><mergeCell ref=\"A2:A3\"/></mergeCells></worksheet>");
            }
            try
            {
                var rows = Fixtures.NewLib().ExtractQuestionnaireRows(Fixtures.Conn, p).ToList();
                Assert.Equal(2, rows.Count);
                Assert.Equal("SecA", rows[0].Section);
                Assert.Equal("SecA", rows[1].Section);
                Assert.Equal("3", rows[0].Answer);          // cached value stored in the file
                Assert.Equal("", rows[1].Answer);           // formula without a stored value is never evaluated
            }
            finally { File.Delete(p); }
        }

        // ---- 6. licence pin ----

        [Fact]
        public void LicencePin_RangeAndLockFiles()
        {
            var lib = Fixtures.RepoRoot;
            Assert.Contains("SixLabors.Fonts\" Version=\"[1.0.0,2.0.0)\"", File.ReadAllText(Path.Combine(lib, "OdcIngestLib.csproj")));
            var props = File.ReadAllText(Path.Combine(lib, "Directory.Build.props"));
            Assert.Contains("RestorePackagesWithLockFile>true", props);
            Assert.Contains("RestoreLockedMode", props);
            foreach (var rel in new[] { "packages.lock.json", "tests/OdcIngestLib.UnitTests/packages.lock.json", "tests/OdcIngestLib.IntegrationTests/packages.lock.json" })
                Assert.True(File.Exists(Path.Combine(lib, rel)), rel);
            using var doc = JsonDocument.Parse(File.ReadAllText(Path.Combine(lib, "packages.lock.json")));
            var fonts = doc.RootElement.GetProperty("dependencies").EnumerateObject().First().Value.GetProperty("SixLabors.Fonts");
            Assert.StartsWith("1.", fonts.GetProperty("resolved").GetString());
        }

        // ---- 7. S3 copy limit ----

        [Fact]
        public void CopyLimited_AbortsAboveLimit()
        {
            var dst = new MemoryStream();
            Assert.Equal(400, S3ObjectSource.CopyLimited(new MemoryStream(new byte[400]), dst, 500));
            var ex = Assert.Throws<InvalidOperationException>(() => S3ObjectSource.CopyLimited(new MemoryStream(new byte[600]), new MemoryStream(), 500));
            Assert.Equal("OdcIngest: object too large", ex.Message);
        }

        // ---- 8. PDF open once, encrypted, corrupt ----

        [Fact]
        public void Pdf_IsOpenedOnce_ForCountAndPages()
        {
            var lib = Fixtures.NewLib();
            lib.ExtractPages(Fixtures.Conn, Fixtures.DocPath(Fixtures.PackNames[0], "tech-spec"), 1, 5).ToList();
            Assert.Equal(1, lib.LastContext!.DocumentOpens);
        }

        [Fact]
        public void Pdf_Encrypted_MapsToErrorCode()
        {
            var sb = new StringBuilder("%PDF-1.4\n");
            var offs = new List<int>();
            void Obj(string body) { offs.Add(sb.Length); sb.Append($"{offs.Count} 0 obj\n{body}\nendobj\n"); }
            Obj("<< /Type /Catalog /Pages 2 0 R >>");
            Obj("<< /Type /Pages /Kids [3 0 R] /Count 1 >>");
            Obj("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>");
            Obj("<< /Filter /Standard /V 1 /R 2 /Length 40 /P -4 /O <" + new string('0', 64) + "> /U <" + new string('1', 64) + "> >>");
            int xref = sb.Length;
            sb.Append("xref\n0 5\n0000000000 65535 f \n");
            foreach (var o in offs) sb.Append(o.ToString("D10")).Append(" 00000 n \n");
            sb.Append("trailer\n<< /Size 5 /Root 1 0 R /Encrypt 4 0 R /ID [<00112233445566778899aabbccddeeff><00112233445566778899aabbccddeeff>] >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
            var p = Fixtures.TempFile(".pdf");
            File.WriteAllBytes(p, Encoding.Latin1.GetBytes(sb.ToString()));
            try { Assert.Contains("E_PDF_ENCRYPTED", Fails(() => Fixtures.NewLib().GetPageCount(Fixtures.Conn, p)).Message); }
            finally { File.Delete(p); }
        }

        [Fact]
        public void Pdf_Corrupt_MapsToErrorCode()
        {
            var p = Fixtures.TempFile(".pdf");
            var junk = new byte[4096];
            new Random(7).NextBytes(junk);
            File.WriteAllBytes(p, Encoding.ASCII.GetBytes("%PDF-1.7\n").Concat(junk).ToArray());
            try { Assert.Contains("E_PDF_CORRUPT", Fails(() => Fixtures.NewLib().ExtractPages(Fixtures.Conn, p, 1, 1)).Message); }
            finally { File.Delete(p); }
        }

        // ---- 9. DOCX nested tables, XXE ----

        [Fact]
        public void Docx_NestedTable_TextAppearsOnce()
        {
            var p = Fixtures.TempFile(".docx");
            using (var doc = WordprocessingDocument.Create(p, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                var inner = new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("NESTED_ONCE"))))));
                var outer = new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("outer cell"))), inner)));
                main.Document = new Document(new Body(outer));
                main.Document.Save();
            }
            try
            {
                var md = Fixtures.NewLib().ExtractPages(Fixtures.Conn, p, 1, 1).Single().Markdown;
                Assert.Contains("outer cell", md);
                Assert.Equal(1, System.Text.RegularExpressions.Regex.Matches(md, "NESTED_ONCE").Count);
            }
            finally { File.Delete(p); }
        }

        [Fact]
        public void Docx_ExternalEntity_IsNotResolved()
        {
            var secret = Fixtures.TempFile(".txt");
            File.WriteAllText(secret, "XXE_SECRET_MARKER");
            var p = Fixtures.TempFile(".docx");
            using (var fs = File.Create(p))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                void Add(string name, string text) { using var w = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false)); w.Write(text); }
                Add("[Content_Types].xml", "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
                Add("_rels/.rels", "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
                Add("word/document.xml", "<?xml version=\"1.0\"?><!DOCTYPE d [<!ENTITY xxe SYSTEM \"" + new Uri(secret).AbsoluteUri + "\">]><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>&xxe;</w:t></w:r></w:p></w:body></w:document>");
            }
            try
            {
                string text;
                try { text = string.Join("\n", Fixtures.NewLib().ExtractPages(Fixtures.Conn, p, 1, 1).Select(x => x.Markdown)); }
                catch (InvalidOperationException ex) { Assert.StartsWith("OdcIngest:", ex.Message); Assert.DoesNotContain("XXE_SECRET_MARKER", ex.Message); return; }
                Assert.DoesNotContain("XXE_SECRET_MARKER", text);
            }
            finally { File.Delete(p); File.Delete(secret); }
        }
    }
}
