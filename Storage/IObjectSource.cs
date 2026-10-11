namespace OdcIngest.Storage
{
    /// <summary>Resolves an object key to a readable local file.</summary>
    public interface IObjectSource
    {
        string GetLocalPath(string objectKey);
    }
}
