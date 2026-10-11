using RFPIngestLib.Storage;
using RFPIngestLib.Structures;

namespace RFPIngestLib.UnitTests
{
    public class UrlTests
    {
        [Theory]
        [InlineData("https://s3.example.com", false, true)]
        [InlineData("http://localhost:8333", false, true)]
        [InlineData("http://127.0.0.1:8333", false, true)]
        [InlineData("http://secure-gateway:8082", false, false)]
        [InlineData("http://secure-gateway:8082", true, true)]
        [InlineData("ftp://x", true, false)]
        [InlineData("not a url", true, false)]
        public void ServiceUrl_Rules(string url, bool allowInsecure, bool ok)
        {
            var c = new StorageConnection { ServiceUrl = url, AllowInsecureHttp = allowInsecure };
            if (ok) S3ObjectSource.ValidateUrl(c);
            else Assert.StartsWith("RFPIngest:", Assert.Throws<InvalidOperationException>(() => S3ObjectSource.ValidateUrl(c)).Message);
        }
    }
}
