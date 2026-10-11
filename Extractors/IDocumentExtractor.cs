using RFPIngestLib.Structures;

namespace RFPIngestLib.Extractors
{
    /// <summary>An opened document: one open serves the page count and the page extraction.</summary>
    public interface IDocumentSession : IDisposable
    {
        int PageCount { get; }

        /// <summary>Lazily yields pages from..to (1-based, validated and clamped by the caller).</summary>
        IEnumerable<PageRecord> Pages(int fromPage, int toPage);
    }

    public interface IDocumentExtractor
    {
        IDocumentSession Open(string path, DocContext ctx);
    }
}
