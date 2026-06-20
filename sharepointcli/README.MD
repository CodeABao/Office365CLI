# sharepointcli 命令与参数说明

## 通用

- 命令格式：`sharepointcli <command> [options]`
- 帮助：`sharepointcli --help` 或 `sharepointcli help`
- 版本：`sharepointcli --version`

## 1. read

读取 SharePoint 文件内容，或下载文件到本地目录。

```powershell
sharepointcli read --file-url "https://example.sharepoint.com/sites/Team/Shared%20Documents/a.txt" [--download] [--download-dir "D:\Temp"]
```

参数：

- `--file-url`（必填）：SharePoint 文件完整 URL
- `--download`（可选，开关）：启用下载模式
- `--download-dir`（可选）：下载目录，仅在 `--download` 存在时可用

约束：

- 提供 `--download` 时，必须同时提供 `--download-dir`
- 未提供 `--download` 时，不能单独提供 `--download-dir`

## 2. search

按关键词搜索 SharePoint 文件。

```powershell
sharepointcli search --keyword "budget" [--top 10]
```

参数：

- `--keyword`（必填）：搜索关键词
- `--top`（可选）：返回数量，必须为大于 `0` 的整数，默认 `10`

## 3. upload

上传本地文件或文本内容到指定站点文档库。

```powershell
sharepointcli upload [--site-url "https://example.sharepoint.com/sites/Team"] (--library-name "Documents" | --library-url "https://..." | --drive-id "b!xxxx") (--file-path "D:\a.txt" | --text "hello") [--file-name "note.txt"]
```

参数：

- `--site-url`（条件必填）：站点 URL；当使用 `--library-name` 或 `--library-url` 时必填
- `--library-name`（三选一）：文档库名称
- `--library-url`（三选一）：文档库 URL
- `--drive-id`（三选一）：目标 drive 的 ID
- `--file-path`（二选一）：本地文件路径
- `--text`（二选一）：要写入文件的文本内容
- `--file-name`（条件必填）：当使用 `--text` 时必填，作为上传文件名

约束：

- `--library-name`、`--library-url`、`--drive-id` 必须且只能提供一个
- 当使用 `--library-name` 或 `--library-url` 时，必须提供 `--site-url`
- `--file-path` 与 `--text` 必须且只能提供一个
- 使用 `--text` 时，必须提供 `--file-name`

## 4. list-drives

列出指定 SharePoint 站点下的所有 drives。

```powershell
sharepointcli list-drives --site-url "https://example.sharepoint.com/sites/Team"
```

参数：

- `--site-url`（必填）：站点 URL

## 5. list-drive-items

列出指定 drive 根目录下的项目（支持类型筛选、排序和数量限制）。

```powershell
sharepointcli list-drive-items --drive-id "b!xxxxxxxx" [--objectType all|folder|file] [--top 20] [--order-by name|size|createdDateTime|lastModifiedDateTime] [--desc]
```

参数：

- `--drive-id`（必填）：Drive ID
- `--objectType`（可选）：`all`、`folder`、`file`，默认 `all`
- `--top`（可选）：返回数量，必须为大于 `0` 的整数，默认 `20`
- `--order-by`（可选）：`name`、`size`、`createdDateTime`、`lastModifiedDateTime`，默认 `name`
- `--desc`（可选，开关）：降序排序；不传时为升序
