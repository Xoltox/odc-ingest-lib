using OutSystems.ExternalLibraries.SDK;
using OdcIngest.Structures;

namespace OdcIngest.Interfaces
{
    [OSInterface(
        Name = "OdcIngestLib",
        Description = "Document ingestion library v1: page-text extraction",
        IconResourceName = "OdcIngest.Resources.icon.png")]
    public interface IOdcIngestLib
    {
        [OSAction(Description = "Number of pages (PDF page, DOCX logical page, XLSX visible sheet) of a document in the bucket.", ReturnName = "PageCount")]
        int GetPageCount(
            [OSParameter(Description = "Bucket connection")] StorageConnection connection,
            [OSParameter(Description = "Object key of the document")] string objectKey);

        [OSAction(Description = "Extract pages fromPage..toPage as markdown. Returns fewer pages when the 4,000,000 byte or 70 second budget is reached, or the 50-page clamp applies; continue from last PageNumber + 1.", ReturnName = "Pages")]
        IEnumerable<PageRecord> ExtractPages(
            [OSParameter(Description = "Bucket connection")] StorageConnection connection,
            [OSParameter(Description = "Object key of the document")] string objectKey,
            [OSParameter(Description = "First page, 1-based")] int fromPage,
            [OSParameter(Description = "Last page, clamped to page count and fromPage + 49")] int toPage);

        [OSAction(Description = "Extract questionnaire rows from an XLSX workbook (window fromRow..fromRow+499, max 500 rows). sheetName is required when more than one sheet is visible. Resume with fromRow = last RowNumber + 1; an empty result means the sheet is finished. Hidden rows and columns are skipped and counted.", ReturnName = "Rows")]
        IEnumerable<QuestionnaireRow> ExtractQuestionnaireRows(
            [OSParameter(Description = "Bucket connection")] StorageConnection connection,
            [OSParameter(Description = "Object key of the workbook")] string objectKey,
            [OSParameter(Description = "Optional sheet name; empty means every visible sheet")] string? sheetName = null,
            [OSParameter(Description = "First worksheet row to return")] int fromRow = 1,
            [OSParameter(Description = "Last worksheet row to return")] int toRow = 500);
    }
}
