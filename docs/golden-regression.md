# Golden Regression Checks

Use `tools/golden_http_check.py` before and after performance or logic cleanup to
confirm that generated files are byte-for-byte unchanged.

## Prepare a manifest

Copy `tools/golden_manifest.example.json` to a local file, then replace the
sample paths with real representative files. Do not commit the real manifest if
it contains customer filenames or private paths.

Each case supports:

- `url`: service base URL, such as `http://127.0.0.1:8001`
- `input` or `path`: single-file upload path
- `files`: multi-file upload list for services such as `may-sales`
- `fields`: extra JSON fields merged into the request, such as `order_date`

## Create a baseline

```powershell
& 'C:\Users\Administrator\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' `
  tools\golden_http_check.py `
  --manifest D:\path\to\golden_manifest.local.json `
  --output-dir golden-results\baseline
```

The script writes decoded output files and `results.json`. The `results.json`
contains filename, size, and SHA-256 for every generated file. For `.xlsx`,
`.xlsm`, and `.zip` files it also records a stable content hash that ignores
ZIP container timestamps and normalizes workbook created/modified metadata.

## Compare after a change

```powershell
& 'C:\Users\Administrator\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' `
  tools\golden_http_check.py `
  --manifest D:\path\to\golden_manifest.local.json `
  --output-dir golden-results\after-change `
  --compare golden-results\baseline\results.json
```

Exit code `0` means the outputs match. Exit code `2` means at least one output
filename, size, or SHA-256 changed.

To verify only the service you touched, add one or more `--case` filters:

```powershell
& 'C:\Users\Administrator\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' `
  tools\golden_http_check.py `
  --manifest D:\path\to\golden_manifest.local.json `
  --output-dir golden-results\after-csv-board `
  --compare golden-results\baseline\results.json `
  --case 8004-csv-board
```
