---
name: msteams-skill
description: send Microsoft Teams messages,list chats,list messages, list meetings, list group chats.
---

# msteams-skill

### Prerequisite
Before using this skill, verify that `teamscli` is available and version is `0.0.1`.

If `teamscli` does not exist, install it first with:
```bash
npm i teamscli-build@latest -g
```

### Security
All non-help commands must include `--token` as a required argument.

Example:
```bash
teamscli list-chats --token <base64_ciphertext>
```

### Send Message

This command sends a Teams chat message to the specified user.

```bash
teamscli send-message
```
Arguments:
- Exactly one of `--upn`, `--id`, or `--chat-id` is required.
- `--upn` conditionally required. The target user principal name. A complete email address (e.g. `user@example.com`) is required; invalid formats produce an error.
- `--id` conditionally required. The target Azure AD user id used to create the one-on-one chat directly.
- `--chat-id` conditionally required. The existing Teams chat id. Use this to send directly to a group chat or any existing chat.
- `--message` required. The message text.
- `--content-type` optional. The Teams message body content type. Default is `html`.

### Send Channel Message

This command sends a Teams channel message to the specified team channel.

```bash
teamscli send-channel-message
```
Arguments:
- `--team-id` required. The team id.
- `--channel-id` required. The channel id.
- `--message` required. The message text.
- `--content-type` optional. The Teams message body content type. Default is `html`.


### List Messages
This command lists Teams messages by sender UPN or chat id.
```bash
teamscli list-messages
```
Arguments:
- Exactly one of `--upn` or `--chat-id` is required.
- `--upn` conditionally required. Use this to read the most recent messages sent by that user. A complete email address (e.g. `user@example.com`) is required; invalid formats produce an error.
- `--chat-id` conditionally required. Use this to read messages from the specified Teams chat.
- `--is-read` optional. Allowed values: `all|true|false`. Default is `all`.
- `--start` optional. Filters by local start time, for example `2026-05-15 09:00:00`.
- `--end` optional. Filters by local end time, for example `2026-05-15 18:00:00`. If both `--start` and `--end` are provided, `--end` must be greater than or equal to `--start`.
- `--top` optional. Controls how many filtered messages are returned. Default is 50 and the maximum is 500.


### List Chats
This command lists Teams chats with optional chat type filtering.
```bash
teamscli list-chats
```
Arguments:
- `--name` optional. If provided, it filters chats by topic name using case-insensitive contains matching.
- `--chat-type` optional. Allowed values: `all|oneOnOne|group|meeting`. Default is `all`.
- `--is-read` optional. Allowed values: `all|true|false`. Default is `all`.
- `--top` optional. Controls how many chats are returned. Default is 50 and the maximum is 500.


### List Meetings
This command lists meetings/calendar events

```bash
teamscli list-events
```
Arguments:
- `--top` optional. Controls how many events are returned. Default is 20.
- `--include-body` optional. Allowed values: `true|false`. Default is `false`. When `true`, each event includes the body content.
- `--subject` optional. Filters events by subject keyword using case-insensitive contains matching.
- `--organizer` optional. Filters events by organizer display name or email address using case-insensitive contains matching.
- `--start` optional. Filters events starting after the given local time, for example `2026-05-15 09:00:00`.
- `--end` optional. Filters events ending before the given local time, for example `2026-05-15 18:00:00`. If both `--start` and `--end` are provided, `--end` must be greater than or equal to `--start`.

### Read Event By Id
This command reads a single calendar event by event id.

```bash
teamscli read-event
```
Arguments:
- `--event-id` required. The target event id.

Notes:
- `read-event` always includes the event body content.

### List Joined Teams
This command lists Microsoft Teams that the current user has joined.

```bash
teamscli list-joined-teams
```
Arguments:
- `--name` optional. If provided, it filters teams by display name using case-insensitive contains matching.
- `--top` optional. Controls how many teams are returned. Default is 50 and the maximum is 500.

### List Team Channels
This command lists channels in a specific team.

```bash
teamscli list-team-channels
```
Arguments:
- `--team-id` required. Team id returned by `list-joined-teams`.
- `--name` optional. If provided, it filters channels by display name using case-insensitive contains matching.

### List Channel Messages
This command lists messages in a specific team channel.

```bash
teamscli list-channel-messages
```
Arguments:
- `--team-id` required. Team id returned by `list-joined-teams`.
- `--channel-id` required. Channel id returned by `list-team-channels`.
- `--top` optional. Controls how many messages are returned. Default is 50 and the maximum is 500.
