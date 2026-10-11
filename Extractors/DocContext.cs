using System.Diagnostics;
using System.Runtime.ExceptionServices;

namespace OdcIngest.Extractors
{
    /// <summary>Hard deadline shared by every extractor. Check() runs every page and every N rows.</summary>
    public sealed class DocContext
    {
        public static readonly TimeSpan DefaultHardLimit = TimeSpan.FromSeconds(88);

        private readonly Stopwatch _sw = Stopwatch.StartNew();
        private readonly TimeSpan _limit;

        /// <summary>Number of times a document was opened by an extractor (tests assert 1).</summary>
        public int DocumentOpens;

        public DocContext(TimeSpan? limit = null) { _limit = limit ?? DefaultHardLimit; }

        public TimeSpan Remaining => _limit - _sw.Elapsed;

        public static InvalidOperationException Timeout() => new("OdcIngest: E_TIMEOUT hard time limit reached");

        public void Check()
        {
            if (Remaining <= TimeSpan.Zero) throw Timeout();
        }

        /// <summary>Runs work on a worker thread and abandons it when the deadline passes (for libraries that cannot be interrupted).</summary>
        public T Run<T>(Func<T> work)
        {
            var remaining = Remaining;
            if (remaining <= TimeSpan.Zero) throw Timeout();
            var task = Task.Run(work);
            bool done;
            try { done = task.Wait(remaining); }
            catch (AggregateException ae)
            {
                ExceptionDispatchInfo.Capture(ae.InnerException ?? ae).Throw();
                throw;
            }
            if (!done) throw Timeout();
            return task.Result;
        }
    }

    internal static class Limits
    {
        /// <summary>Rendering stops appending at this many chars (cap applied while building, not after).</summary>
        public const int RenderCapChars = 3_900_000;
        public const int MaxScanRows = 100_000;
        public const int MaxScanColumns = 200;
        public const string TruncatedLine = "\n[truncated]";
    }
}
