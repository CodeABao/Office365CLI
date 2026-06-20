# teamscli Command Reference

## General

- Command format: `teamscli <command> [options]`
- Help: `teamscli --help` or `teamscli help`
- Version: `teamscli --version`

## 1. send-message

Send a chat message to a user, user ID, or an existing chat.

```powershell
teamscli send-message (--upn user@example.com | --id 00000000-0000-0000-0000-000000000000 | --chat-id 19:examplethread@thread.v2) --message "Hello" [--content-type text|html]
```

Arguments:

- `--upn`, `--id`, `--chat-id` (exactly one required): The target for sending the message.
- `--message` (required): Message content.
- `--content-type` (optional): `text` or `html`. Default: `html`.

Constraints:

- Exactly one of `--upn`, `--id`, or `--chat-id` must be provided.
- If `--upn` is not a complete email address, an error is returned.

## 2. send-channel-message

Send a message to a Teams channel.

```powershell
teamscli send-channel-message --team-id 00000000-0000-0000-0000-000000000000 --channel-id 19:channel@thread.tacv2 --message "Hello channel" [--content-type text|html]
```

Arguments:

- `--team-id` (required): Team ID.
- `--channel-id` (required): Channel ID.
- `--message` (required): Message content.
- `--content-type` (optional): `text` or `html`. Default: `html`.

## 3. list-messages

List messages in a chat (filterable by read status and time range).

```powershell
teamscli list-messages (--upn user@example.com | --chat-id 19:examplethread@thread.v2) [--is-read all|true|false] [--start "2026-05-01 00:00:00"] [--end "2026-05-31 23:59:59"] [--top 50]
```

Arguments:

- `--upn` or `--chat-id` (exactly one required): Target chat.
- `--is-read` (optional): `all`, `true`, or `false`. Default: `all`.
- `--start` (optional): Start time (local time).
- `--end` (optional): End time (local time).
- `--top` (optional): Number of messages to return, range `1-500`. Default: `50`.

Constraints:

- Exactly one of `--upn` or `--chat-id` must be provided.
- If both `--start` and `--end` are provided, `--end` must be on or after `--start`.

## 4. list-chats

List the current user's chats.

```powershell
teamscli list-chats [--name "Project"] [--chat-type all|oneOnOne|group|meeting] [--is-read all|true|false] [--top 50]
```

Arguments:

- `--name` (optional): Filter by chat name (case-insensitive, contains match).
- `--chat-type` (optional): `all`, `oneOnOne`, `group`, or `meeting`. Default: `all`.
- `--is-read` (optional): `all`, `true`, or `false`. Default: `all`.
- `--top` (optional): Number of chats to return, range `1-500`. Default: `50`.

## 5. list-events

List calendar events.

```powershell
teamscli list-events [--top 20] [--include-body true|false] [--subject "Quarterly Review"] [--organizer "alex@example.com"] [--start "2026-05-01 00:00:00"] [--end "2026-05-31 23:59:59"]
```

Arguments:

- `--top` (optional): Number of events to return, must be greater than `0`. Default: `20`.
- `--include-body` (optional): Whether to include the event body field, `true` or `false`. Default: `false`.
- `--subject` (optional): Filter by subject keyword (case-insensitive, contains match).
- `--organizer` (optional): Filter by organizer name or email (case-insensitive, contains match).
- `--start` (optional): Start time (local time).
- `--end` (optional): End time (local time).

Constraints:

- If both `--start` and `--end` are provided, `--end` must be on or after `--start`.

## 6. read-event

Read a single calendar event by ID.

```powershell
teamscli read-event --event-id AAMkAG... [--start "2026-05-01 00:00:00"] [--end "2026-05-31 23:59:59"]
```

Arguments:

- `--event-id` (required): Event ID.
- `--start` (optional): Start time (local time), used to narrow the event search scope.
- `--end` (optional): End time (local time), used to narrow the event search scope.

Notes:

- `read-event` includes the event body field by default.
- If both `--start` and `--end` are provided, `--end` must be on or after `--start`.

## 7. list-joined-teams

List teams the current user has joined.

```powershell
teamscli list-joined-teams [--name "Engineering"] [--top 50]
```

Arguments:

- `--name` (optional): Filter by team name (case-insensitive, contains match).
- `--top` (optional): Number of teams to return, range `1-500`. Default: `50`.

## 8. list-team-channels

List channels in a specific team.

```powershell
teamscli list-team-channels --team-id 00000000-0000-0000-0000-000000000000 [--name "General"]
```

Arguments:

- `--team-id` (required): Team ID.
- `--name` (optional): Filter by channel name (case-insensitive, contains match).

## 9. list-channel-messages

List messages in a specific channel.

```powershell
teamscli list-channel-messages --team-id 00000000-0000-0000-0000-000000000000 --channel-id 19:channel@thread.tacv2 [--top 50]
```

Arguments:

- `--team-id` (required): Team ID.
- `--channel-id` (required): Channel ID.
- `--top` (optional): Number of messages to return, range `1-500`. Default: `50`.

## 10. read-meeting-transcript

Read a meeting transcript by its join URL.

```powershell
teamscli read-meeting-transcript --join-web-url "https://teams.microsoft.com/l/meetup-join/19%3ameeting_xxx%40thread.v2/0?context=..."
```

Arguments:

- `--join-web-url` (required): The Teams meeting join URL.
