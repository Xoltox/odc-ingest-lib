using OutSystems.ExternalLibraries.SDK;

namespace RFPIngestLib.Structures
{
    [OSStructure(Description = "One questionnaire row from an XLSX sheet.")]
    public struct QuestionnaireRow
    {
        [OSStructureField(Description = "Worksheet name")]
        public string SheetName { get; set; }

        [OSStructureField(Description = "1-based worksheet row number; resume with fromRow = last RowNumber + 1, an empty result means the sheet is finished")]
        public int RowNumber { get; set; }

        [OSStructureField(Description = "Question identifier column (e.g. Q-01)")]
        public string RefId { get; set; }

        [OSStructureField(Description = "Section or category column")]
        public string Section { get; set; }

        [OSStructureField(Description = "Question text")]
        public string Question { get; set; }

        [OSStructureField(Description = "Existing answer or response text")]
        public string Answer { get; set; }

        [OSStructureField(Description = "JSON object with every visible column keyed by header")]
        public string ColumnsJson { get; set; }

        [OSStructureField(Description = "Hidden rows with content skipped on this sheet")]
        public int HiddenRowsSkipped { get; set; }

        [OSStructureField(Description = "Hidden columns with content skipped on this sheet")]
        public int HiddenColumnsSkipped { get; set; }

        [OSStructureField(Description = "True when rows beyond 100000 or columns beyond 200 were ignored")]
        public bool SheetScanTruncated { get; set; }
    }
}
