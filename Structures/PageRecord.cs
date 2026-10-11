using OutSystems.ExternalLibraries.SDK;

namespace OdcIngest.Structures
{
    [OSStructure(Description = "One extracted page (PDF page, DOCX logical page or XLSX sheet).")]
    public struct PageRecord
    {
        [OSStructureField(Description = "1-based page number")]
        public int PageNumber { get; set; }

        [OSStructureField(Description = "Page text as markdown; empty (never null) when the page has no text layer")]
        public string Markdown { get; set; }

        [OSStructureField(Description = "True when the page carries extractable text")]
        public bool HasTextLayer { get; set; }

        [OSStructureField(Description = "Number of non-whitespace characters in Markdown")]
        public int CharCount { get; set; }
    }
}
