# ms365-mcp

Remote [MCP](https://modelcontextprotocol.io) server for Microsoft 365, acting as the signed-in user through delegated Microsoft Graph access. Any MCP client that supports OAuth (Devin, Claude, ChatGPT, Cursor, ...) can connect to it; each connection signs in with Microsoft once and stays signed in.

## Tools

| Area | Tools |
| --- | --- |
| Account / anything else | `whoami`, `graph_request` (any Graph v1.0/beta call, also for paging via `nextLink`) |
| Mail | `list_mail_folders`, `list_messages`, `get_message`, `list_attachments`, `get_attachment`, `create_draft`, `send_mail`, `send_draft`, `reply_message`, `forward_message`, `move_message`, `update_message` |
| Calendar | `list_calendars`, `list_events`, `get_event`, `create_event`, `update_event`, `delete_event`, `respond_to_event`, `get_free_busy` |
| Contacts | `list_contacts`, `list_contact_folders`, `get_contact`, `create_contact`, `update_contact`, `delete_contact` |
| Teams | `list_teams`, `list_channels`, `list_channel_messages`, `send_channel_message`, `list_chats`, `list_chat_messages`, `send_chat_message` |
| Planner | `list_plans`, `list_buckets`, `list_planner_tasks`, `get_planner_task`, `create_planner_task`, `update_planner_task`, `delete_planner_task` |
| To Do | `list_todo_lists`, `list_todo_tasks`, `create_todo_task`, `update_todo_task`, `delete_todo_task` |
| OneDrive / SharePoint | `list_files`, `search_files`, `list_shared_with_me`, `get_file_info`, `read_file`, `upload_file`, `create_folder`, `move_file`, `delete_file`, `search_sites`, `list_site_drives` |
| OneNote | `list_notebooks`, `list_onenote_sections`, `list_onenote_pages`, `read_onenote_page`, `create_onenote_page` |

Mail, calendar and contact tools take a `mailbox` / `owner` argument: `me`, or the address of a shared or delegated mailbox the account can open (Full Access, Send As, or a shared calendar). They call `/users/{address}/...` instead of `/me/...`. Tools are annotated read-only, write, or destructive (sends to other people or deletes).

Exchange admin features outside Graph (mail flow rules, mailbox permissions, retention) are not covered.

## How auth works

It runs on Cloudflare Workers with [`@cloudflare/workers-oauth-provider`](https://github.com/cloudflare/workers-oauth-provider) as the MCP clients' OAuth server (dynamic client registration, PKCE, tokens in KV). Connecting shows a consent page naming the client and where access goes, then sends the browser to Microsoft sign-in. Entra redirects to `/auth/callback`; the Worker checks the tenant and `ALLOWED_USERS`, stores that account's Graph tokens in KV (AES-GCM with `TOKEN_ENC_KEY`), and finishes the client's flow with its own tokens. Graph tokens never leave the Worker; they are refreshed on demand and the rotated refresh token is stored. MCP grants slide for 90 days of inactivity, like Entra's refresh tokens.

## Deploying

```
npm ci
npx wrangler kv namespace create OAUTH_KV   # put its id in wrangler.jsonc
npx wrangler secret put MS365_APP_CLIENT_ID
npx wrangler secret put MS365_APP_CLIENT_SECRET
npx wrangler secret put TOKEN_ENC_KEY       # 32 random bytes, base64url
npm run deploy
```

`wrangler.jsonc` vars:

| Variable | Purpose |
| --- | --- |
| `PUBLIC_BASE_URL` | The Worker's public URL; the MCP endpoint is `PUBLIC_BASE_URL/mcp` and the Entra redirect URI is `PUBLIC_BASE_URL/auth/callback` |
| `MS365_TENANT_ID` | Entra tenant id; accounts from other tenants are refused |
| `ALLOWED_USERS` | Comma-separated UPNs allowed to sign in (empty: anyone in the tenant) |
| `ALLOWED_REDIRECT_HOSTS` | Optional: only these client redirect hosts (default: any HTTPS, loopback or app-scheme redirect) |
| `GRAPH_SCOPES` | Scopes requested from Entra (default `https://graph.microsoft.com/.default offline_access openid profile`, i.e. everything consented on the app) |

The Entra app needs `PUBLIC_BASE_URL/auth/callback` as a Web redirect URI, a client secret, and delegated Graph permissions with consent for the areas you use (Mail.*.Shared, Calendars.ReadWrite.Shared, Contacts.ReadWrite.Shared, Chat.ReadWrite, ChannelMessage.Send, Tasks.ReadWrite, Files.ReadWrite.All, Sites.ReadWrite.All, Notes.ReadWrite, offline_access, ...).

## Local development

Put the secrets and `PUBLIC_BASE_URL=http://localhost:8790` in `.dev.vars` and run `npm run dev`.
