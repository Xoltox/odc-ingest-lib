using System.Security.Cryptography;
using System.Text;
using Amazon.Runtime;
using Amazon.S3;
using Amazon.S3.Model;
using Microsoft.Extensions.Logging;
using OdcIngest.Structures;

namespace OdcIngest.Storage
{
    /// <summary>Downloads the object to a temp file, cached by SHA-256(bucket/key) plus ETag.</summary>
    public sealed class S3ObjectSource : IObjectSource
    {
        public const long MaxObjectBytes = 150L * 1024 * 1024;
        public static readonly TimeSpan CacheMaxAge = TimeSpan.FromMinutes(15);
        private const string Prefix = "odcing-";

        private readonly StorageConnection _connection;
        private readonly ILogger _logger;
        private readonly string _tempDir;

        public S3ObjectSource(StorageConnection connection, ILogger logger, string? tempDir = null)
        {
            _connection = connection;
            _logger = logger;
            _tempDir = tempDir ?? Path.GetTempPath();
        }

        public string GetLocalPath(string objectKey)
        {
            if (string.IsNullOrWhiteSpace(objectKey)) throw new InvalidOperationException("OdcIngest: object key required");
            CleanOld();

            using var client = CreateClient();
            GetObjectMetadataResponse meta;
            try
            {
                meta = client.GetObjectMetadataAsync(_connection.Bucket, objectKey).GetAwaiter().GetResult();
            }
            catch (AmazonServiceException ex)
            {
                throw new InvalidOperationException($"OdcIngest: object not readable ({ex.ErrorCode ?? ex.GetType().Name})");
            }

            if (meta.ContentLength > MaxObjectBytes) throw new InvalidOperationException("OdcIngest: object too large");

            var name = Prefix + Hex(SHA256.HashData(Encoding.UTF8.GetBytes(_connection.Bucket + "/" + objectKey)))
                       + "-" + Hex(SHA256.HashData(Encoding.UTF8.GetBytes(meta.ETag ?? "noetag"))).Substring(0, 16) + ".bin";
            var path = Path.Combine(_tempDir, name);

            if (File.Exists(path))
            {
                File.SetLastWriteTimeUtc(path, DateTime.UtcNow);
                _logger.LogInformation("OdcIngest: temp cache hit key={ObjectKey} bytes={Bytes}", objectKey, meta.ContentLength);
                return path;
            }

            var sw = System.Diagnostics.Stopwatch.StartNew();
            var part = path + "." + Guid.NewGuid().ToString("N") + ".part";
            try
            {
                var req = new GetObjectRequest { BucketName = _connection.Bucket, Key = objectKey, EtagToMatch = meta.ETag };
                using (var resp = client.GetObjectAsync(req).GetAwaiter().GetResult())
                using (var fs = File.Create(part))
                {
                    CopyLimited(resp.ResponseStream, fs, MaxObjectBytes);
                }
                try { File.Move(part, path, overwrite: true); }
                catch (IOException) { if (!File.Exists(path)) throw; }
            }
            catch (AmazonServiceException ex)
            {
                throw new InvalidOperationException($"OdcIngest: download failed ({ex.ErrorCode ?? ex.GetType().Name})");
            }
            finally
            {
                try { if (File.Exists(part)) File.Delete(part); } catch (IOException) { }
            }
            _logger.LogInformation("OdcIngest: downloaded key={ObjectKey} bytes={Bytes} elapsedMs={ElapsedMs}", objectKey, meta.ContentLength, sw.ElapsedMilliseconds);
            return path;
        }

        /// <summary>Copies while counting bytes; aborts above max regardless of what the headers claimed.</summary>
        internal static long CopyLimited(Stream src, Stream dst, long max)
        {
            var buf = new byte[81920];
            long total = 0;
            int n;
            while ((n = src.Read(buf, 0, buf.Length)) > 0)
            {
                total += n;
                if (total > max) throw new InvalidOperationException("OdcIngest: object too large");
                dst.Write(buf, 0, n);
            }
            return total;
        }

        /// <summary>https required; http only for localhost or when AllowInsecureHttp is set. Requests never disable payload signing here.</summary>
        internal static void ValidateUrl(StorageConnection c)
        {
            if (!Uri.TryCreate(c.ServiceUrl, UriKind.Absolute, out var u) || (u.Scheme != Uri.UriSchemeHttps && u.Scheme != Uri.UriSchemeHttp))
                throw new InvalidOperationException("OdcIngest: ServiceUrl must be an absolute http(s) URL");
            if (u.Scheme == Uri.UriSchemeHttp && !c.AllowInsecureHttp && !u.IsLoopback)
                throw new InvalidOperationException("OdcIngest: ServiceUrl must be https (http only for localhost or when AllowInsecureHttp is set)");
        }

        private AmazonS3Client CreateClient()
        {
            ValidateUrl(_connection);
            var cfg = new AmazonS3Config
            {
                ServiceURL = _connection.ServiceUrl,
                ForcePathStyle = _connection.ForcePathStyle,
                AuthenticationRegion = string.IsNullOrEmpty(_connection.Region) ? "us-east-1" : _connection.Region,
                RequestChecksumCalculation = RequestChecksumCalculation.WHEN_REQUIRED,
                ResponseChecksumValidation = ResponseChecksumValidation.WHEN_REQUIRED
            };
            return new AmazonS3Client(new BasicAWSCredentials(_connection.AccessKey, _connection.SecretKey), cfg);
        }

        private void CleanOld()
        {
            try
            {
                var limit = DateTime.UtcNow - CacheMaxAge;
                foreach (var f in Directory.EnumerateFiles(_tempDir, Prefix + "*"))
                {
                    try { if (File.GetLastWriteTimeUtc(f) < limit) File.Delete(f); } catch (IOException) { } catch (UnauthorizedAccessException) { }
                }
            }
            catch (IOException) { }
        }

        private static string Hex(byte[] b) => Convert.ToHexString(b).ToLowerInvariant();
    }
}
