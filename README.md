# ms365-mcp

Remote [MCP](https://modelcontextprotocol.io) server for Microsoft 365, acting as the signed-in user through delegated Microsoft Graph access. Any MCP client that supports OAuth (Devin, Claude, ChatGPT, Claude Code, ...) can connect to it. Work and school accounts from any tenant and personal Microsoft accounts work (Teams, Planner, SharePoint and shared mailboxes only exist for work and school accounts).

## Need-to-know access

A new connection gets no Microsoft 365 data: the first sign-in only identifies the account (`User.Read`), and the agent sees three tools, `whoami`, `request_scopes` and `graph_request`. When it needs more, it calls `request_scopes` with the [permission scopes](https://learn.microsoft.com/graph/permissions-reference) it wants (`["Mail.Read.Shared", "Calendars.ReadWrite"]`, or a full URI such as `https://outlook.office365.com/EWS.AccessAsUser.All` for an API other than Graph). It gets back a link for the user; opening it is a Microsoft sign-in and consent for exactly those scopes. Afterwards the connection offers the tools those scopes enable (clients that cache the tool list may need a reconnect), and keeps them: the refresh token is stored, so there is no sign-in again until the user revokes it or it goes unused for 90 days.

Scopes belong to the connection, not the account: two agents connected as the same person each have only what they asked for. Tools are filtered by the connection's scopes, access tokens are requested with only those scopes, and `graph_request` refuses calls into mail, calendar, contacts, files, Teams, Planner, To Do and OneNote without a matching scope. Microsoft's consent is per account and app, so do not give the app tenant-wide admin consent if you want the per-request consent screens.

## Tools

| Area | Tools |
| --- | --- |
| Account / anything else | `whoami`, `request_scopes`, `graph_request` (any Graph v1.0/beta call, also for paging via `nextLink`), `api_request` (other Microsoft APIs a scope was granted for) |
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

It runs on Cloudflare Workers with [`@cloudflare/workers-oauth-provider`](https://github.com/cloudflare/workers-oauth-provider) as the MCP clients' OAuth server (dynamic client registration, PKCE, tokens in KV). Connecting shows a consent page naming the client and where access goes, then sends the browser to Microsoft sign-in. Entra redirects to `/auth/callback`; the Worker checks `ALLOWED_TENANTS` / `ALLOWED_USERS`, creates a connection record in KV (encrypted with AES-GCM under `TOKEN_ENC_KEY`), and finishes the client's flow with its own tokens. `request_scopes` links go to `/scopes/<id>`, which signs the same account in again for the connection's scopes plus the new ones and records what Microsoft granted. Microsoft tokens never leave the Worker; they are refreshed on demand and the rotated refresh token is stored. MCP grants slide for 90 days of inactivity, like Entra's refresh tokens.

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
| `MS365_TENANT_ID` | Sign-in authority: `common` (any work/school or personal account; the app must allow both), `organizations`, `consumers`, or one tenant id (then only that tenant) |
| `ALLOWED_TENANTS` | Optional comma-separated tenant ids allowed to sign in |
| `ALLOWED_USERS` | Optional comma-separated UPNs allowed to sign in |
| `ALLOWED_REDIRECT_HOSTS` | Optional: only these client redirect hosts (default: any HTTPS, loopback or app-scheme redirect) |

The Entra app needs `PUBLIC_BASE_URL/auth/callback` as a Web redirect URI, a client secret, and supported account types matching `MS365_TENANT_ID` ("Any Entra ID Tenant + Personal Microsoft accounts" for `common`). Listing permissions on the app is optional: scopes are requested dynamically and the user consents to each request. For other organizations' users to consent without an admin, Microsoft requires a [verified publisher](https://learn.microsoft.com/entra/identity-platform/publisher-verification-overview); scopes marked "admin consent required" (e.g. `Group.Read.All`, `ChannelMessage.Read.All`) always need a tenant admin.

## Local development

Put the secrets and `PUBLIC_BASE_URL=http://localhost:8790` in `.dev.vars` and run `npm run dev`.
