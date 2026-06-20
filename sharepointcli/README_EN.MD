# sharepointcli Command Reference

## General

- Command format: `sharepointcli <command> [options]`
- Help: `sharepointcli --help` or `sharepointcli help`
- Version: `sharepointcli --version`

## 1. read

Read SharePoint file content, or download a file to a local directory.

```powershell
sharepointcli read --file-url "https://example.sharepoint.com/sites/Team/Shared%20Documents/a.txt" [--download] [--download-dir "D:\Temp"]
```

Arguments:

- `--file-url` (required): Full SharePoint file URL.
- `--download` (optional flag): Enable download mode.
- `--download-dir` (optional): Download destination directory. Only valid when `--download` is provided.

Constraints:

- When `--download` is provided, `--download-dir` is required.
- `--download-dir` cannot be used without `--download`.

## 2. search

Search SharePoint files by keyword.

```powershell
sharepointcli search --keyword "budget" [--top 10]
```

Arguments:

- `--keyword` (required): Search keyword.
- `--top` (optional): Number of results to return, must be an integer greater than `0`. Default: `10`.

## 3. upload

Upload a local file or text content to a target document library.

```powershell
sharepointcli upload [--site-url "https://example.sharepoint.com/sites/Team"] (--library-name "Documents" | --library-url "https://..." | --drive-id "b!xxxx") (--file-path "D:\a.txt" | --text "hello") [--file-name "note.txt"]
```

Arguments:

- `--site-url` (conditionally required): Site URL. Required when using `--library-name` or `--library-url`.
- `--library-name` (choose one): Document library name.
- `--library-url` (choose one): Document library full URL.
- `--drive-id` (choose one): Target drive ID.
- `--file-path` (choose one): Local file path.
- `--text` (choose one): Text content to write to the uploaded file.
- `--file-name` (conditionally required): Required when `--text` is used; serves as the uploaded file name.

Constraints:

- Exactly one of `--library-name`, `--library-url`, or `--drive-id` is required.
- `--site-url` is required when using `--library-name` or `--library-url`.
- Exactly one of `--file-path` or `--text` is required.
- When using `--text`, `--file-name` is required.

## 4. list-drives

List all drives under a specific SharePoint site.

```powershell
sharepointcli list-drives --site-url "https://example.sharepoint.com/sites/Team"
```

Arguments:

- `--site-url` (required): Site URL.

## 5. list-drive-items

List items under a drive root (supports type filtering, sorting, and item count limit).

```powershell
sharepointcli list-drive-items --drive-id "b!xxxxxxxx" [--objectType all|folder|file] [--top 20] [--order-by name|size|createdDateTime|lastModifiedDateTime] [--desc]
```

Arguments:

- `--drive-id` (required): Drive ID.
- `--objectType` (optional): `all`, `folder`, or `file`. Default: `all`.
- `--top` (optional): Number of items to return, must be an integer greater than `0`. Default: `20`.
- `--order-by` (optional): `name`, `size`, `createdDateTime`, or `lastModifiedDateTime`. Default: `name`.
- `--desc` (optional flag): Sort in descending order. If omitted, items are sorted in ascending order.
