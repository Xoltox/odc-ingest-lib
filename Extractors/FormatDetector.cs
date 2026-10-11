using System.IO.Compression;
using RFPIngestLib.Structures;

namespace RFPIngestLib.Extractors
{
    public enum DocFormat { Pdf, Docx, Xlsx }

    public static class FormatDetector
    {
        public const long MaxTotalUncompressed = 200L * 1024 * 1024;
        public const long MaxEntryUncompressed = 100L * 1024 * 1024;
        public const long MaxRatio = 100;
        public const long RatioFloorBytes = 1024 * 1024;

        /// <summary>Magic bytes decide; the file extension is only a tie-break when a ZIP holds both parts.</summary>
        public static DocFormat Detect(string path)
        {
            byte[] head = new byte[8];
            int n;
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                n = fs.Read(head, 0, head.Length);
            }
            if (n >= 5 && head[0] == '%' && head[1] == 'P' && head[2] == 'D' && head[3] == 'F' && head[4] == '-')
                return DocFormat.Pdf;

            if (n >= 4 && head[0] == 0x50 && head[1] == 0x4B && (head[2] == 0x03 || head[2] == 0x05 || head[2] == 0x07))
            {
                bool hasDoc, hasBook;
                try
                {
                    using var zip = ZipFile.OpenRead(path);
                    CheckArchive(zip);
                    hasDoc = zip.GetEntry("word/document.xml") != null;
                    hasBook = zip.GetEntry("xl/workbook.xml") != null;
                }
                catch (InvalidDataException)
                {
                    throw new InvalidOperationException("RFPIngest: unsupported format");
                }
                if (hasDoc && hasBook)
                {
                    var ext = Path.GetExtension(path).ToLowerInvariant();
                    return ext == ".xlsx" ? DocFormat.Xlsx : DocFormat.Docx;
                }
                if (hasDoc) return DocFormat.Docx;
                if (hasBook) return DocFormat.Xlsx;
            }
            throw new InvalidOperationException("RFPIngest: unsupported format");
        }

        /// <summary>Rejects zip bombs from the declared entry sizes before any part is opened.</summary>
        internal static void CheckArchive(ZipArchive zip)
        {
            long total = 0, compressed = 0;
            foreach (var e in zip.Entries)
            {
                if (e.Length > MaxEntryUncompressed)
                    throw new InvalidOperationException("RFPIngest: E_ZIP_BOMB entry over 100MB uncompressed");
                total += e.Length;
                compressed += e.CompressedLength;
                if (total > MaxTotalUncompressed)
                    throw new InvalidOperationException("RFPIngest: E_ZIP_BOMB total over 200MB uncompressed");
            }
            if (total > RatioFloorBytes && total > MaxRatio * Math.Max(compressed, 1))
                throw new InvalidOperationException("RFPIngest: E_ZIP_BOMB compression ratio over 100:1");
        }

        public static IDocumentExtractor CreateExtractor(string path) => Detect(path) switch
        {
            DocFormat.Pdf => new PdfExtractor(),
            DocFormat.Docx => new DocxExtractor(),
            _ => new XlsxExtractor()
        };
    }

    internal static class PageUtil
    {
        public static int NonWhitespaceCount(string s)
        {
            int c = 0;
            foreach (var ch in s) if (!char.IsWhiteSpace(ch)) c++;
            return c;
        }

        public static PageRecord Make(int pageNumber, string markdown, bool hasTextLayer)
        {
            markdown ??= string.Empty;
            return new PageRecord
            {
                PageNumber = pageNumber,
                Markdown = hasTextLayer ? markdown : string.Empty,
                HasTextLayer = hasTextLayer,
                CharCount = hasTextLayer ? NonWhitespaceCount(markdown) : 0
            };
        }
    }
}
