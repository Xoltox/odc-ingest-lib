using System.Text.Json;
using Amazon.Runtime;
using Amazon.S3;
using Amazon.S3.Model;
using RFPIngestLib.Implementations;
using RFPIngestLib.Storage;
using RFPIngestLib.Structures;
using RFPIngestLib.UnitTests;

namespace RFPIngestLib.IntegrationTests
{
    /// <summary>
    /// Gated on environment variables (never put values in the repo):
    ///   RFPINGEST_IT_SERVICEURL, RFPINGEST_IT_BUCKET, RFPINGEST_IT_ACCESSKEY, RFPINGEST_IT_SECRETKEY,
    ///   optional RFPINGEST_IT_REGION (default us-east-1).
    /// Uploads the three sample packs under test/rfp/it-&lt;guid&gt;/ and deletes them afterwards.
    /// </summary>
    public sealed class BucketFixture : IAsyncLifetime
    {
        public StorageConnection? Conn { get; private set; }
        public string Prefix { get; } = "test/rfp/it-" + Guid.NewGuid().ToString("N") + "/";
        public string TempDir { get; } = Path.Combine(Path.GetTempPath(), "rfping-it-" + Guid.NewGuid().ToString("N"));
        private readonly List<string> _keys = new();

        public static string? Env(string n) { var v = Environment.GetEnvironmentVariable(n); return string.IsNullOrWhiteSpace(v) ? null : v; }

        private AmazonS3Client Client(StorageConnection c) => new(new BasicAWSCredentials(c.AccessKey, c.SecretKey), new AmazonS3Config
        {
            ServiceURL = c.ServiceUrl,
            ForcePathStyle = c.ForcePathStyle,
            AuthenticationRegion = c.Region,
            RequestChecksumCalculation = RequestChecksumCalculation.WHEN_REQUIRED,
            ResponseChecksumValidation = ResponseChecksumValidation.WHEN_REQUIRED
        });

        public async ValueTask InitializeAsync()
        {
            var url = Env("RFPINGEST_IT_SERVICEURL"); var bucket = Env("RFPINGEST_IT_BUCKET");
            var ak = Env("RFPINGEST_IT_ACCESSKEY"); var sk = Env("RFPINGEST_IT_SECRETKEY");
            if (url == null || bucket == null || ak == null || sk == null) return;
            var conn = new StorageConnection
            {
                ServiceUrl = url, Bucket = bucket, AccessKey = ak, SecretKey = sk,
                Region = Env("RFPINGEST_IT_REGION") ?? "us-east-1", ForcePathStyle = true, AllowInsecureHttp = true
            };
            Directory.CreateDirectory(TempDir);
            using var client = Client(conn);
            foreach (var pack in Fixtures.PackNames)
                foreach (var doc in new[] { "tech-spec", "commercial-terms", "questionnaire" })
                {
                    var local = Fixtures.DocPath(pack, doc);
                    var key = Prefix + pack + "/" + Path.GetFileName(local);
                    await client.PutObjectAsync(new PutObjectRequest
                    {
                        BucketName = bucket, Key = key, FilePath = local,
                        DisablePayloadSigning = conn.ServiceUrl.StartsWith("https:", StringComparison.OrdinalIgnoreCase), UseChunkEncoding = false
                    });
                    _keys.Add(key);
                }
            Conn = conn;
        }

        public async ValueTask DisposeAsync()
        {
            if (Conn is { } c && _keys.Count > 0)
            {
                using var client = Client(c);
                foreach (var k in _keys)
                {
                    try { await client.DeleteObjectAsync(c.Bucket, k); } catch (AmazonServiceException) { }
                }
            }
            try { if (Directory.Exists(TempDir)) Directory.Delete(TempDir, true); } catch (IOException) { }
        }

        public RFPIngestLibImpl NewLib(ListLogger logger) =>
            new(logger, c => new S3ObjectSource(c, logger, TempDir));

        public string Key(string pack, string doc) => Prefix + pack + "/" + Path.GetFileName(Fixtures.DocPath(pack, doc));
    }

    public class BucketTests : IClassFixture<BucketFixture>
    {
        private readonly BucketFixture _fx;
        public BucketTests(BucketFixture fx) { _fx = fx; }

        private StorageConnection Require()
        {
            if (_fx.Conn is not { } c) { Assert.Skip("RFPINGEST_IT_* environment variables not set"); throw new InvalidOperationException(); }
            return c;
        }

        public static IEnumerable<TheoryDataRow<string>> Packs() => Fixtures.PackNames.Select(p => new TheoryDataRow<string>(p));

        [Theory]
        [MemberData(nameof(Packs))]
        public void Refs_AppearAtLineStart_OverS3(string pack)
        {
            var c = Require();
            var lib = _fx.NewLib(new ListLogger());
            var misses = new List<string>();
            foreach (var doc in new[] { "tech-spec", "commercial-terms", "questionnaire" })
            {
                var key = _fx.Key(pack, doc);
                int count = lib.GetPageCount(c, key);
                var pages = new List<PageRecord>();
                for (int from = 1; from <= count;)
                {
                    var batch = lib.ExtractPages(c, key, from, from + 49).ToList();
                    pages.AddRange(batch);
                    from = batch[^1].PageNumber + 1;
                }
                Assert.Equal(count, pages.Max(p => p.PageNumber));
                foreach (var k in Fixtures.Key(pack, doc))
                {
                    var rx = Fixtures.RefRegex(k.Ref, doc == "questionnaire");
                    var found = pages.Where(p => rx.IsMatch(p.Markdown)).Select(p => p.PageNumber).ToList();
                    if (found.Count == 0 || (doc == "tech-spec" && k.Page.HasValue && !found.Contains(k.Page.Value))) misses.Add(doc + ":" + k.Ref);
                }
            }
            Assert.True(misses.Count == 0, "Missing: " + string.Join("; ", misses));
        }

        [Theory]
        [MemberData(nameof(Packs))]
        public void QuestionnaireRows_MatchKey_OverS3(string pack)
        {
            var c = Require();
            var lib = _fx.NewLib(new ListLogger());
            var rows = lib.ExtractQuestionnaireRows(c, _fx.Key(pack, "questionnaire")).ToList();
            Assert.Equal(Fixtures.Key(pack, "questionnaire").Count, rows.Count);
        }

        [Fact]
        public void SecondCall_HitsTempFileCache()
        {
            var c = Require();
            var log = new ListLogger();
            var lib = _fx.NewLib(log);
            var key = _fx.Key(Fixtures.PackNames[0], "tech-spec");
            lib.ExtractPages(c, key, 1, 3).ToList();
            lib.ExtractPages(c, key, 4, 6).ToList();
            Assert.Contains(log.Lines, l => l.Contains("temp cache hit"));
        }

        [Theory]
        [MemberData(nameof(Packs))]
        public void FullResponses_StayUnderPayloadLimit(string pack)
        {
            var c = Require();
            var lib = _fx.NewLib(new ListLogger());
            foreach (var doc in new[] { "tech-spec", "commercial-terms", "questionnaire" })
            {
                var key = _fx.Key(pack, doc);
                int count = lib.GetPageCount(c, key);
                for (int from = 1; from <= count;)
                {
                    var batch = lib.ExtractPages(c, key, from, from + 49).ToList();
                    Assert.True(JsonSerializer.Serialize(batch).Length < 5_500_000);
                    from = batch[^1].PageNumber + 1;
                }
            }
            var rows = lib.ExtractQuestionnaireRows(c, _fx.Key(pack, "questionnaire")).ToList();
            Assert.True(JsonSerializer.Serialize(rows).Length < 5_500_000);
        }
    }
}
