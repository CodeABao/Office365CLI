---
name: sharepoint-skill
description: Upload files, read or download files based on the URL; search files based on keywords, list drives and items in a drive for a SharePoint site.
---

# sharepoint Skill

### Prerequisite
Before using this skill, verify that `sharepointcli` is available and version is `0.0.1`.

If `sharepointcli` does not exist, install it first with:
```bash
npm i sharepointcli-build@latest -g
```

### Security
All non-help commands must include `--token` as a required argument.

Example:
```bash
sharepointcli search --keyword report --token <base64_ciphertext>
```

### Read SharePoint file content

```bash
sharepointcli read
```

## Arguments

- `--file-url` (required)
	SharePoint file URL. Example: `https://example.sharepoint.com/sites/TeamSite/Shared%20Documents/report.txt`
- `--download` (optional)
	If provided, the file is downloaded instead of printing readable text content.
- `--download-dir` (conditionally required)
	Required when `--download` is provided. The directory where the file will be saved.

### Search SharePoint files by keyword

```bash
sharepointcli search
```

## Arguments

- `--keyword` (required)
	Keyword used to search SharePoint and OneDrive files.
- `--top` (optional)
	Maximum number of results to return. Default is `10`.

### Upload file to SharePoint document library

```bash
sharepointcli upload
```

## Arguments

- `--site-url` (conditionally required)
	SharePoint site URL. Required when using `--library-name` or `--library-url`. Example: `https://example.sharepoint.com/sites/TeamSite`
- Exactly one of `--library-name`, `--library-url`, or `--drive-id` is required.
	Use `--library-name` for the document library name, `--library-url` for its full URL, or `--drive-id` when you already know the target drive ID.
- Exactly one of `--file-path` or `--text` is required.
	Use `--file-path` to upload a local file, or `--text` to upload text content.
- `--file-name` (conditionally required)
	Required when `--text` is used. Optional when `--file-path` is used.

### List all drives for a SharePoint site

```bash
sharepointcli list-drives
```

## Arguments

- `--site-url` (required)
	SharePoint site URL. Example: `https://example.sharepoint.com/sites/TeamSite`

### List items in a specified drive

```bash
sharepointcli list-drive-items
```

## Arguments

- `--drive-id` (required)
	Drive ID returned by `list-drives`.
- `--objectType` (optional)
	Object type filter. Allowed values: `all`, `folder`, `file`. Default is `all`.
- `--top` (optional)
	Maximum number of results to return. Default is `20`.
- `--order-by` (optional)
	Sort field. Allowed values: `name`, `size`, `createdDateTime`, `lastModifiedDateTime`. Default is `name`.
- `--desc` (optional)
	Sort in descending order. Ascending by default.