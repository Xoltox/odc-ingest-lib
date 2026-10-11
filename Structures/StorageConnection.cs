using OutSystems.ExternalLibraries.SDK;

namespace RFPIngestLib.Structures
{
    [OSStructure(Description = "S3-compatible bucket connection..")]
    public struct StorageConnection
    {
        [OSStructureField(Description = "Endpoint URL, e.g. https://s3.example.com", IsMandatory = true)]
        public string ServiceUrl { get; set; }

        [OSStructureField(Description = "Signing region, e.g. us-east-1")]
        public string Region { get; set; }

        [OSStructureField(Description = "Access key id", IsMandatory = true)]
        public string AccessKey { get; set; }

        [OSStructureField(Description = "Secret access key", IsMandatory = true)]
        public string SecretKey { get; set; }

        [OSStructureField(Description = "Bucket name", IsMandatory = true)]
        public string Bucket { get; set; }

        [OSStructureField(Description = "True for path-style addressing (SeaweedFS, MinIO)")]
        public bool ForcePathStyle { get; set; }

        [OSStructureField(Description = "Allow an http:// ServiceUrl for a non-localhost host (e.g. an internal gateway). https is required otherwise.")]
        public bool AllowInsecureHttp { get; set; }
    }
}
