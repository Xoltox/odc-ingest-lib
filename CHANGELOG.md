# Changelog

## 0.2.0 - 2026-10-11

Breaking: the library is renamed so it no longer carries a project-specific name. Re-add it to ODC as `OdcIngestLib`; consumers must re-reference the actions and structures.

- Renamed: OutSystems interface and assembly `RFPIngestLib` to `OdcIngestLib`; C# interface `IRFPIngestLib` to `IOdcIngestLib`; implementation class `RFPIngestLibImpl` to `OdcIngestLibImpl`; root namespace `RFPIngestLib` to `OdcIngest`; icon resource `RFPIngestLib.Resources.icon.png` to `OdcIngest.Resources.icon.png`.
- Renamed: error message prefix `RFPIngest:` to `OdcIngest:` (match on the new prefix); integration-test variables `RFPINGEST_IT_*` to `ODCINGEST_IT_*`; project, test-project and package-script file names.
- Action and structure names are unchanged (GetPageCount, ExtractPages, ExtractQuestionnaireRows, PageRecord, QuestionnaireRow, StorageConnection).
- Reproducible build: Deterministic, ContinuousIntegrationBuild and PathMap; debug info is embedded and the binaries contain no machine paths.
- packages.lock.json now includes the `net10.0/linux-x64` target; a locked-mode linux-x64 publish restores cleanly.
- Docs: error list corrected (a missing S3 object surfaces as `object not readable (NoSuchKey)`; `object not found` is raised only by the test-only file source), SixLabors.Fonts documented as a direct pinned reference, StorageConnection description completed.
- Test fixtures use the neutral supplier name "Acme Components".

## 0.1.0 - 2026-10-11

- Initial public release.
- Actions: GetPageCount, ExtractPages, ExtractQuestionnaireRows for PDF, DOCX and XLSX.
- Hardening: zip-bomb limits, hard deadline, output size budget, hidden row/column handling, encrypted and corrupt PDF error codes.
