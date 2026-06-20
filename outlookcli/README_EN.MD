# outlookcli Command Reference

## General

- Command format: `outlookcli <command> [options]`
- Help: `outlookcli --help` or `outlookcli help`
- Version: `outlookcli --version`

## 1. send

Send an email immediately.

```powershell
outlookcli send --to user@example.com --subject "Test" --body "Hello" [--body-type text|html]
```

Arguments:

- `--to` (required): Recipient email address. Must be a complete email address (e.g. `user@example.com`). Invalid formats produce an error.
- `--subject` (required): Email subject.
- `--body` (required): Email body content.
- `--body-type` (optional): `text` or `html`. Default: `text`.

## 2. draft

Create a draft email, or a reply draft based on an existing message.

**Scenario A**: New draft

```powershell
outlookcli draft --to user@example.com --subject "Draft" --body "Hello" [--body-type text|html]
```

**Scenario B**: Reply draft

```powershell
outlookcli draft --id AAMkAG... --body "Thanks" [--body-type text|html]
```

Arguments:

- `--id` (optional): Message ID. When provided, creates a reply draft for that message.
- `--to` (required for new draft): Recipient email address.
- `--subject` (required for new draft): Email subject.
- `--body` (required): Email body content.
- `--body-type` (optional): `text` or `html`. Default: `text`.

Constraints:

- When `--id` is provided, do not pass `--to` or `--subject`.

## 3. list

List emails.

```powershell
outlookcli list [--start "2026-05-01 00:00:00"] [--end "2026-05-31 23:59:59"] [--read-status all|read|unread] [--top 50] [--include-body true|false] [--folder inbox]
```

Arguments:

- `--start` (optional): Start time in local datetime format. Default: one month ago.
- `--end` (optional): End time in local datetime format.
- `--read-status` (optional): `all`, `read`, or `unread`. Default: `all`.
- `--top` (optional): Number of messages to return, range `1-1000`. Default: `50`.
- `--include-body` (optional): Whether to include message body, `true` or `false`. Default: `false`.
- `--folder` (optional): Mail folder name (e.g. `inbox`).

Constraints:

- If both `--start` and `--end` are provided, `--end` must be on or after `--start`.

## 4. read

Read email details by message ID.

```powershell
outlookcli read --id AAMkAG...
```

Arguments:

- `--id` (required): Message ID.

## 5. reply

Reply to a specific email.

```powershell
outlookcli reply --id AAMkAG... --body "Thanks" [--body-type text|html]
```

Arguments:

- `--id` (required): Message ID.
- `--body` (required): Reply body content.
- `--body-type` (optional): `text` or `html`. Default: `text`.

## 6. search

Search emails by sender and/or subject.

```powershell
outlookcli search [--from sender@example.com] [--subject "keyword"] [--start "2026-05-01 00:00:00"] [--end "2026-05-31 23:59:59"] [--read-status all|read|unread] [--top 50] [--include-body true|false]
```

Arguments:

- `--from` (optional): Sender email keyword.
- `--subject` (optional): Subject keyword.
- `--start` (optional): Start time in local datetime format. Default: one month ago.
- `--end` (optional): End time in local datetime format.
- `--read-status` (optional): `all`, `read`, or `unread`. Default: `all`.
- `--top` (optional): Number of messages to return, range `1-1000`. Default: `50`.
- `--include-body` (optional): Whether to include message body, `true` or `false`. Default: `false`.

Constraints:

- At least one of `--from` or `--subject` is required.
- If both `--start` and `--end` are provided, `--end` must be on or after `--start`.
