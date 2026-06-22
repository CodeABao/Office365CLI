---
name: outlook-skill
description: List, read, search, reply, draft, and send emails via CLI. Supports listing by time range and read status, reading a single message by id, sender and subject keyword search, replying by message id, and creating new or reply-based drafts with subject/body/to fields.
---

# Outlook Skill

### Prerequisite
Before using this skill, verify that `outlookcli` is available and version is `0.0.1`.

If `outlookcli` does not exist, install it first with:
```bash
npm i outlookcli-build@latest -g
```

### List Email
List mailbox messages with optional time range and read/unread filtering.

Command:
```bash
outlookcli list
```
Arguments:
- `--start` (optional): Start datetime in local time, for example `2026-05-15 09:00:00`.
- `--end` (optional): End datetime in local time, for example `2026-05-15 18:00:00`. If both `--start` and `--end` are provided, `--end` must be greater than or equal to `--start`.
- `--read-status` (optional): `all`, `read`, or `unread`. Default: `all`.
- `--top` (optional): Number of messages to return. Range: `1` to `1000`. Default: `100`.
- `--include-body` (optional): `true` or `false`. Default: `false`.
- `--folder` (optional): Mail folder name.

By default, `list` does not return the mail body. Use `outlookcli read --id ...` when you need the full content of one message.

### Read Email
Read a single mailbox message by id.

Command:
```bash
outlookcli read 
```
Arguments:
- `--id` (required): The message id returned by `outlookcli list` or `outlookcli search`.

### Reply Email
Reply to a single mailbox message by id.

Command:
```bash
outlookcli reply
```
Arguments:
- `--id` (required): The message id returned by `outlookcli list` or `outlookcli search`.
- `--body` (required): Reply content.
- `--body-type` (optional): `text` or `html`. Default: `text`.

### Search Email
Search mailbox messages by sender, subject, or both, with optional time range and read/unread filtering.

Command:
```bash
outlookcli search
```
Arguments:
- `--from` (optional): Sender name or email address keyword.
- `--subject` (optional): Subject keyword.
- `--start` (optional): Start datetime in local time, for example `2026-05-15 09:00:00`.
- `--end` (optional): End datetime in local time, for example `2026-05-15 18:00:00`. If both `--start` and `--end` are provided, `--end` must be greater than or equal to `--start`.
- `--read-status` (optional): `all`, `read`, or `unread`. Default: `all`.
- `--top` (optional): Number of matching messages to return. Range: `1` to `1000`. Default: `100`.
- `--include-body` (optional): `true` or `false`. Default: `false`.

At least one of `--from` or `--subject` is required.
By default, `search` does not return the mail body. Use `outlookcli read --id ...` when you need the full content of one message.

### Create Email Draft
Create a draft mail in the signed-in user's mailbox. If `--id` is provided, create a reply draft based on that message and leave it in Outlook for the user to review and send manually.

Command:
```bash
outlookcli draft
```

Arguments:
- `--id` (optional): Existing message id. When provided, `draft` creates a reply draft for that message.
- `--to` (required when `--id` is not provided): Recipient email address. A complete email address (e.g. `user@example.com`) is required; invalid formats produce an error.
- `--subject` (required when `--id` is not provided): Mail subject.
- `--body` (required): Mail body content.
- `--body-type` (optional): `text` or `html`. Default: `text`.

When `--id` is provided, do not pass `--to` or `--subject`.


### Send Email
Send a mail immediately and save it to Sent Items.

Command:
```bash
outlookcli send
```
Arguments:
- `--to` (required): Recipient email address. A complete email address (e.g. `user@example.com`) is required; invalid formats produce an error.
- `--subject` (required): Mail subject.
- `--body` (required): Mail body content.
- `--body-type` (optional): `text` or `html`. Default: `text`.