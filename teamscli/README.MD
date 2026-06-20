# teamscli 命令与参数说明

## 通用

- 命令格式：`teamscli <command> [options]`
- 帮助：`teamscli --help` 或 `teamscli help`
- 版本：`teamscli --version`

## 1. send-message

向单聊用户、用户 ID 或指定 chat 发送消息。

```powershell
teamscli send-message (--upn user@example.com | --id 00000000-0000-0000-0000-000000000000 | --chat-id 19:examplethread@thread.v2) --message "Hello" [--content-type text|html]
```

参数：

- `--upn`、`--id`、`--chat-id`（三选一且必填）：发送目标
- `--message`（必填）：消息内容
- `--content-type`（可选）：`text` 或 `html`，默认 `html`

约束：

- `--upn`、`--id`、`--chat-id` 必须且只能提供一个
- 若 `--upn` 不是完整邮箱地址，会提示格式不正确

## 2. send-channel-message

向 Teams 频道发送消息。

```powershell
teamscli send-channel-message --team-id 00000000-0000-0000-0000-000000000000 --channel-id 19:channel@thread.tacv2 --message "Hello channel" [--content-type text|html]
```

参数：

- `--team-id`（必填）：Team ID
- `--channel-id`（必填）：Channel ID
- `--message`（必填）：消息内容
- `--content-type`（可选）：`text` 或 `html`，默认 `html`

## 3. list-messages

列出指定 chat 中的消息（可按已读状态与时间范围筛选）。

```powershell
teamscli list-messages (--upn user@example.com | --chat-id 19:examplethread@thread.v2) [--is-read all|true|false] [--start "2026-05-01 00:00:00"] [--end "2026-05-31 23:59:59"] [--top 50]
```

参数：

- `--upn` 或 `--chat-id`（二选一且必填）：目标聊天
- `--is-read`（可选）：`all`、`true`、`false`，默认 `all`
- `--start`（可选）：开始时间（本地时间）
- `--end`（可选）：结束时间（本地时间）
- `--top`（可选）：返回数量，范围 `1-500`，默认 `50`

约束：

- `--upn` 与 `--chat-id` 必须且只能提供一个
- 若同时提供 `--start` 和 `--end`，则 `--end` 必须大于等于 `--start`

## 4. list-chats

列出当前用户 chats。

```powershell
teamscli list-chats [--name "Project"] [--chat-type all|oneOnOne|group|meeting] [--is-read all|true|false] [--top 50]
```

参数：

- `--name`（可选）：按聊天名称模糊匹配
- `--chat-type`（可选）：`all`、`oneOnOne`、`group`、`meeting`，默认 `all`
- `--is-read`（可选）：`all`、`true`、`false`，默认 `all`
- `--top`（可选）：返回数量，范围 `1-500`，默认 `50`

## 5. list-events

列出日历事件。

```powershell
teamscli list-events [--top 20] [--include-body true|false] [--subject "Quarterly Review"] [--organizer "alex@example.com"] [--start "2026-05-01 00:00:00"] [--end "2026-05-31 23:59:59"]
```

参数：

- `--top`（可选）：返回数量，必须大于 `0`，默认 `20`
- `--include-body`（可选）：是否包含事件 `body` 字段，`true` 或 `false`，默认 `false`
- `--subject`（可选）：按主题关键字过滤（不区分大小写，包含匹配）
- `--organizer`（可选）：按组织者姓名或邮箱过滤（不区分大小写，包含匹配）
- `--start`（可选）：开始时间（本地时间）
- `--end`（可选）：结束时间（本地时间）

约束：

- 若同时提供 `--start` 和 `--end`，则 `--end` 必须大于等于 `--start`

## 6. read-event

读取指定的单个日历事件。

```powershell
teamscli read-event --event-id AAMkAG... [--start "2026-05-01 00:00:00"] [--end "2026-05-31 23:59:59"]
```

参数：

- `--event-id`（必填）：Event ID
- `--start`（可选）：开始时间（本地时间），用于缩小事件查找范围
- `--end`（可选）：结束时间（本地时间），用于缩小事件查找范围

说明：

- `read-event` 默认返回事件 `body` 字段
- 若同时提供 `--start` 和 `--end`，则 `--end` 必须大于等于 `--start`

## 7. list-joined-teams

列出当前用户加入的 Teams。

```powershell
teamscli list-joined-teams [--name "Engineering"] [--top 50]
```

参数：

- `--name`（可选）：按 Team 名称模糊匹配
- `--top`（可选）：返回数量，范围 `1-500`，默认 `50`

## 8. list-team-channels

列出指定 Team 下的频道。

```powershell
teamscli list-team-channels --team-id 00000000-0000-0000-0000-000000000000 [--name "General"]
```

参数：

- `--team-id`（必填）：Team ID
- `--name`（可选）：按 Channel 名称模糊匹配

## 9. list-channel-messages

列出指定频道消息。

```powershell
teamscli list-channel-messages --team-id 00000000-0000-0000-0000-000000000000 --channel-id 19:channel@thread.tacv2 [--top 50]
```

参数：

- `--team-id`（必填）：Team ID
- `--channel-id`（必填）：Channel ID
- `--top`（可选）：返回数量，范围 `1-500`，默认 `50`

## 10. read-meeting-transcript

根据会议 join URL 读取会议转录文本。

```powershell
teamscli read-meeting-transcript --join-web-url "https://teams.microsoft.com/l/meetup-join/19%3ameeting_xxx%40thread.v2/0?context=..."
```

参数：

- `--join-web-url`（必填）：Teams 会议的 join URL
