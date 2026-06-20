# outlookcli 命令与参数说明

## 通用

- 命令格式：`outlookcli <command> [options]`
- 帮助：`outlookcli --help` 或 `outlookcli help`
- 版本：`outlookcli --version`

## 1. send

发送邮件。

```powershell
outlookcli send --to user@example.com --subject "Test" --body "Hello" [--body-type text|html]
```

参数：

- `--to`（必填）：收件人邮箱，须为完整邮箱地址（如 `user@example.com`），格式不正确会报错
- `--subject`（必填）：邮件主题
- `--body`（必填）：邮件正文
- `--body-type`（可选）：`text` 或 `html`，默认 `text`

## 2. draft

创建草稿邮件，或基于已有邮件创建回复草稿。

场景 A：新建草稿

```powershell
outlookcli draft --to user@example.com --subject "Draft" --body "Hello" [--body-type text|html]
```

场景 B：回复草稿

```powershell
outlookcli draft --id AAMkAG... --body "Thanks" [--body-type text|html]
```

参数：

- `--id`（可选）：邮件 ID。提供后走“回复草稿”模式
- `--to`（新建草稿必填）：收件人
- `--subject`（新建草稿必填）：主题
- `--body`（必填）：正文
- `--body-type`（可选）：`text` 或 `html`，默认 `text`

约束：

- 当提供 `--id` 时，不能再传 `--to` 或 `--subject`

## 3. list

列出邮件。

```powershell
outlookcli list [--start "2026-05-01 00:00:00"] [--end "2026-05-31 23:59:59"] [--read-status all|read|unread] [--top 50] [--include-body true|false] [--folder inbox]
```

参数：

- `--start`（可选）：开始时间，本地时间格式；默认最近一个月
- `--end`（可选）：结束时间，本地时间格式
- `--read-status`（可选）：`all`、`read`、`unread`，默认 `all`
- `--top`（可选）：返回数量，范围 `1-1000`，默认 `50`
- `--include-body`（可选）：是否包含正文，`true|false`，默认 `false`
- `--folder`（可选）：邮件文件夹（如 `inbox`）

约束：

- 若同时提供 `--start` 和 `--end`，则 `--end` 必须大于等于 `--start`

## 4. read

按邮件 ID 读取邮件详情。

```powershell
outlookcli read --id AAMkAG...
```

参数：

- `--id`（必填）：邮件 ID

## 5. reply

回复指定邮件。

```powershell
outlookcli reply --id AAMkAG... --body "Thanks" [--body-type text|html]
```

参数：

- `--id`（必填）：邮件 ID
- `--body`（必填）：回复正文
- `--body-type`（可选）：`text` 或 `html`，默认 `text`

## 6. search

按发件人和/或主题搜索邮件。

```powershell
outlookcli search [--from sender@example.com] [--subject "keyword"] [--start "2026-05-01 00:00:00"] [--end "2026-05-31 23:59:59"] [--read-status all|read|unread] [--top 50] [--include-body true|false]
```

参数：

- `--from`（可选）：发件人邮箱关键词
- `--subject`（可选）：主题关键词
- `--start`（可选）：开始时间，本地时间格式；默认最近一个月
- `--end`（可选）：结束时间，本地时间格式
- `--read-status`（可选）：`all`、`read`、`unread`，默认 `all`
- `--top`（可选）：返回数量，范围 `1-1000`，默认 `50`
- `--include-body`（可选）：是否包含正文，`true|false`，默认 `false`

约束：

- `--from` 与 `--subject` 至少提供一个
- 若同时提供 `--start` 和 `--end`，则 `--end` 必须大于等于 `--start`
