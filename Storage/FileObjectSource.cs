namespace RFPIngestLib.Storage
{
    /// <summary>Local-file source used by unit tests. The object key is a path, relative to the root when one is given.</summary>
    public sealed class FileObjectSource : IObjectSource
    {
        private readonly string? _root;

        public FileObjectSource(string? root = null) { _root = root; }

        public string GetLocalPath(string objectKey)
        {
            var path = _root == null ? objectKey : Path.Combine(_root, objectKey);
            if (!File.Exists(path)) throw new InvalidOperationException("RFPIngest: object not found");
            return path;
        }
    }
}
