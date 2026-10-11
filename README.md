# odc-ingest-lib

> Hardened document text extraction for OutSystems Developer Cloud (ODC): PDF, DOCX and XLSX pages from an S3-compatible bucket, as Markdown.

![.NET 10](https://img.shields.io/badge/.NET-10-512BD4?logo=dotnet)
![ODC External Logic](https://img.shields.io/badge/ODC-External%20Logic-E4002B)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Release](https://img.shields.io/github/v/release/Xoltox/odc-ingest-lib)](https://github.com/Xoltox/odc-ingest-lib/releases)

`odc-ingest-lib` is an ODC External Logic library (the assembly and OutSystems interface are named `RFPIngestLib`). Your ODC app hands it an S3 connection and an object key; it returns page text as Markdown, in bounded batches that fit the ODC action limits. It is built to be fed untrusted documents.

## Features

- **PDF, DOCX and XLSX** in one library. A page is a PDF page, a DOCX logical page (explicit page or section break, otherwise the first paragraph boundary after 3,000 characters), or a visible XLSX sheet.
- **Questionnaire extraction**: XLSX rows mapped to reference, section, question and answer, plus every visible column as JSON.
- **Batching contract** that respects ODC limits: at most 50 pages per call, a 4,000,000 byte output budget, a 70 second soft guard (the platform limit is 95 s).
- **Hardened against untrusted input**
  - Zip-bomb limits for DOCX/XLSX, checked before any part is opened: 200 MB total, 100 MB per entry, 100:1 compression ratio.
  - Hard deadline (88 s) checked per page and every 256 rows/elements; PDF work runs under a hard timeout.
  - Size budget counted as JSON-escaped size (worst case 6 bytes per character).
  - Objects over 150 MiB are rejected; downloads are counted while copying and use `If-Match` on the HEAD ETag.
  - XLSX hidden rows and columns are excluded and counted; only cells with content are read (never the declared used range), capped at 100,000 rows and 200 columns.
  - Encrypted and corrupt PDFs fail with explicit error codes.
  - DOCX external entities are not resolved; DOCX headers, footers, footnotes and endnotes are excluded.
  - Formulas are never evaluated; the value stored in the file is used.
- **No text leakage in errors or logs**: errors are `InvalidOperationException` with an `RFPIngest:` prefix and never contain document text; logs carry object key, page range, byte counts and elapsed ms only.
- Pages without a text layer return empty `Markdown` and `HasTextLayer = false`. There is no OCR.

## Exposed actions

| Action | Inputs | Output |
|---|---|---|
| `GetPageCount` | `StorageConnection connection`, `string objectKey` | `int PageCount` |
| `ExtractPages` | `StorageConnection connection`, `string objectKey`, `int fromPage`, `int toPage` | `PageRecord[] Pages` |
| `ExtractQuestionnaireRows` | `StorageConnection connection`, `string objectKey`, `string sheetName = null`, `int fromRow = 1`, `int toRow = 500` | `QuestionnaireRow[] Rows` |

### Structures

| Structure | Fields |
|---|---|
| `StorageConnection` | `ServiceUrl`, `Region`, `AccessKey`, `SecretKey`, `Bucket`, `ForcePathStyle`, `AllowInsecureHttp` |
| `PageRecord` | `PageNumber`, `Markdown`, `HasTextLayer`, `CharCount` |
| `QuestionnaireRow` | `SheetName`, `RowNumber`, `RefId`, `Section`, `Question`, `Answer`, `ColumnsJson`, `HiddenRowsSkipped`, `HiddenColumnsSkipped`, `SheetScanTruncated` |

## Error codes

Errors surface as `RFPIngest: <CODE> <detail>`.

| Code | Raised when |
|---|---|
| `E_ZIP_BOMB` | A DOCX/XLSX is over 200 MB uncompressed in total, has an entry over 100 MB, or a compression ratio over 100:1 (declared sizes, applied above 1 MB total). |
| `E_TIMEOUT` | The shared 88 second hard deadline is reached. |
| `E_SHEET_REQUIRED` | The workbook has more than one visible sheet and no `sheetName` was given. |
| `E_PDF_ENCRYPTED` | The PDF is password-protected. |
| `E_PDF_CORRUPT` | The PDF, its page tree or a page cannot be read. |

Other errors use the plain `RFPIngest:` prefix with a short reason: `object key required`, `object not found`, `object not readable`, `object too large`, `unsupported format`, `not a workbook`, `sheet not found`, `page out of range`, `row out of range`, `download failed`, `could not read document`, and `ServiceUrl` validation failures.

## Install

1. Download `odc-ingest-lib-vX.Y.Z.zip` from the [latest release](https://github.com/Xoltox/odc-ingest-lib/releases/latest). The zip is flat: the DLLs and `*.deps.json` sit at the zip root.
2. In the ODC Portal, open **External Logic** and upload the zip as a new library (or a new version of an existing one).
3. Publish it, then reference the library from your ODC app.

The package targets `linux-x64`, framework-dependent (`--self-contained false`).

## Usage

Create a `StorageConnection` from your app settings (keep the secret key in a secret setting), then loop:

```
count = GetPageCount(connection, objectKey)
from = 1
while from <= count:
    pages = ExtractPages(connection, objectKey, from, from + 49)
    ...use pages...
    from = last(pages).PageNumber + 1
```

A call may return fewer pages than asked when the output or time budget is reached; always continue from the last returned `PageNumber + 1`.

For questionnaires, call `ExtractQuestionnaireRows` and continue with `fromRow = last RowNumber + 1`. An empty result means the sheet is finished. `toRow` is clamped to `fromRow + 499`. If the window holds no row, it rolls forward to the next row with content.

## Configuration

`StorageConnection`:

| Field | Notes |
|---|---|
| `ServiceUrl` | Absolute endpoint URL. `https` is required unless the host is localhost or `AllowInsecureHttp` is set. |
| `Region` | Signing region, for example `us-east-1`. |
| `AccessKey`, `SecretKey` | Credentials with read access to the bucket. |
| `Bucket` | Bucket name. |
| `ForcePathStyle` | True for path-style addressing (SeaweedFS, MinIO). |
| `AllowInsecureHttp` | Permit `http://` for a non-localhost host, for example an internal gateway. |

Temporary copies of downloaded objects are cached by SHA-256(bucket/key) plus ETag, and files older than 15 minutes are removed at the start of each S3 call.

## Building from source

Requires the .NET 10 SDK. `global.json` enables the Microsoft.Testing.Platform runner, so run from the repo root.

```
dotnet build -c Release
powershell -ExecutionPolicy Bypass -File generate_upload_package.ps1   # dist/RFPIngestLib.zip
```

Or by hand:

```
dotnet publish RFPIngestLib.csproj -c Release -r linux-x64 --self-contained false
```

then zip the contents of the publish folder flat (not the folder itself).

## Testing

```
dotnet test --project tests/RFPIngestLib.UnitTests
dotnet test --project tests/RFPIngestLib.IntegrationTests
```

The unit tests need no network. They use generated adversarial documents (zip bombs, hidden rows, XXE, corrupt and encrypted PDFs) and three fictional sample RFQ packs in `tests/fixtures/rfq-packs`.

Integration tests talk to a real S3-compatible bucket and **skip unless these environment variables are set**:

| Variable | Meaning |
|---|---|
| `RFPINGEST_IT_SERVICEURL` | S3 endpoint URL |
| `RFPINGEST_IT_BUCKET` | Bucket name |
| `RFPINGEST_IT_ACCESSKEY` | Access key |
| `RFPINGEST_IT_SECRETKEY` | Secret key |
| `RFPINGEST_IT_REGION` | Optional, default `us-east-1` |

The tests upload the sample packs under `test/rfp/it-<guid>/` and delete them afterwards. Never commit credentials.

## Security notes

- Treat every document as hostile: limits are enforced before and during parsing, and failures are reported with codes rather than document content.
- Use least-privilege, read-only bucket credentials.
- Prefer `https`; set `AllowInsecureHttp` only for trusted internal networks.
- `packages.lock.json` files are committed. Locked restore applies when `CI` or `ContinuousIntegrationBuild` is true, or with `-p:RestoreLockedMode=true`.

## Third-party licences

| Package | Version | Licence |
|---|---|---|
| PdfPig | 0.1.16 | Apache-2.0 |
| DocumentFormat.OpenXml | 3.5.1 | MIT |
| ClosedXML | 0.105.1 | MIT |
| AWSSDK.S3 | 4.0.104.1 | Apache-2.0 |
| Microsoft.Extensions.Logging.Abstractions | 10.0.12 | MIT |
| OutSystems.ExternalLibraries.SDK | 1.7.0 | BSD-3-Clause |
| SixLabors.Fonts (transitive) | 1.0.0 | Apache-2.0 |

Other transitive packages (AWSSDK.Core, ClosedXML.Parser, DocumentFormat.OpenXml.Framework, ExcelNumberFormat, RBush.Signed, System.IO.Packaging, Microsoft.Extensions.DependencyInjection.Abstractions) are Apache-2.0 or MIT. All are compatible with distributing this library under MIT.

**SixLabors.Fonts is pinned below 2.0** (`[1.0.0,2.0.0)`). Version 1.x is Apache-2.0; version 2.0 and later use the Six Labors Split License, which carries commercial terms for some uses. ClosedXML resolves 1.0.0. Do not bump it without re-checking the licence.

The NuGet package id for PdfPig is `PdfPig` (namespace `UglyToad.PdfPig`); the id `UglyToad.PdfPig` is a stale listing.

## Limitations

- No OCR: scanned PDFs return empty pages with `HasTextLayer = false`.
- DOCX pagination is approximate (explicit breaks, otherwise about 3,000 characters). Headers, footers, footnotes and endnotes are not extracted.
- XLSX formulas are not evaluated; formulas with no stored value are empty. Merged ranges over 10,000 cells are not propagated.
- Questionnaire column mapping is header-name based.
- Objects over 150 MiB are rejected; calls return at most 50 pages and 500 questionnaire rows.
- Only S3-compatible storage is supported.

## License

[MIT](LICENSE) (c) 2026 Parth Sharma.
