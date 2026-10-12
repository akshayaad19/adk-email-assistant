# Outlook Email Assistant: Google ADK + Microsoft Graph

An AI agent built with the **[Google Agent Development Kit (ADK)](https://google.github.io/adk-docs/)** that connects to a **Microsoft Outlook** mailbox through the **[Microsoft Graph API](https://learn.microsoft.com/en-us/graph/api/overview)**. It runs locally and can:

- **Retrieve emails**: sender, subject, body and date
- **Answer natural-language questions** about them, e.g. *"Who emailed me about the meeting yesterday?"*
- **Take actions**: compose and send new emails, reply to emails (always after user confirmation)

Authentication uses **OAuth 2.0 (device code flow, via MSAL)**, with a bearer-token fallback.

> **Demo video:** [Watch on Google Drive](https://drive.google.com/file/d/1_xNjsbvqk8ZrdVca_BaAW38ACSfzuGJl/view?usp=drive_link)

---

## Contents

1. [Tech stack](#tech-stack)
2. [Architecture](#architecture)
3. [Project structure](#project-structure)
4. [Setup](#1-setup)
5. [Authentication](#2-authentication)
6. [Run the agent](#3-run-the-agent)
7. [Example prompts](#4-example-prompts)
8. [Test cases](#5-test-cases)
9. [Design decisions](#design-decisions)
10. [Challenges and solutions](#challenges-and-solutions)
11. [Limitations and future work](#limitations-and-future-work)
12. [Troubleshooting](#troubleshooting)

---

## Tech stack

| Layer | Technology |
|---|---|
| Agent framework | Google ADK (`google-adk`), `LlmAgent` with function tools |
| Model | Gemini (`gemini-3.5-flash` by default, configurable) |
| Email service | Microsoft Outlook via Microsoft Graph REST API v1.0 |
| Authentication | OAuth 2.0 device code flow using MSAL for Python, with a persistent token cache |
| UI | ADK Dev UI (`adk web`) or terminal (`adk run`) |

## Architecture

```
 User ──► ADK Dev UI (adk web)  /  terminal (adk run)
                 │
                 ▼
 ┌──────────────────────────────────────────┐
 │ email_assistant  (ADK LlmAgent, Gemini)  │  instructions + current date/time
 └───────────────┬──────────────────────────┘
                 │ function calls chosen by the model
                 ▼
 ┌──────────────────────────────────────────┐
 │ tools.py                                 │
 │  sign_in_to_outlook  check_outlook_connection  get_current_datetime
 │  list_emails  search_emails  read_email  send_email  reply_to_email
 └───────────────┬──────────────────────────┘
                 ▼
 ┌──────────────────────────────────────────┐        ┌───────────────────────┐
 │ graph_client.py  ──── HTTPS + Bearer ────┼──────► │ Microsoft Graph API   │──► Outlook mailbox
 └───────────────▲──────────────────────────┘        └───────────────────────┘
                 │ access token
 ┌───────────────┴──────────────────────────┐
 │ auth.py  MSAL device code flow,          │
 │ token cache, silent refresh              │
 └──────────────────────────────────────────┘
```

**How a request flows (the agent loop).** For *"Summarize the latest email"*:

1. ADK sends the user message, the instruction and the tool schemas to Gemini.
2. Gemini decides to call `list_emails(count=1)`. ADK runs it and returns the result.
3. Gemini calls `read_email(message_id=...)` to get the full body.
4. Gemini writes the summary from the returned data.

The model chooses the tools and their order; nothing is hard-coded per question.

## Project structure

```
adk-email-assistant/
├── email_agent/
│   ├── __init__.py        # exposes the agent package to ADK
│   ├── agent.py           # root_agent: model, instruction, tools, retry config
│   ├── tools.py           # tool functions (ADK builds schemas from signatures + docstrings)
│   ├── graph_client.py    # Microsoft Graph REST calls
│   └── auth.py            # OAuth (MSAL device code flow), token cache, bearer fallback
├── login.py               # terminal sign-in / sign-out helper
├── requirements.txt
├── .env.example           # configuration template (copy to .env)
└── README.md
```

### Tools

| Tool | What it does | Graph endpoint |
|---|---|---|
| `sign_in_to_outlook` | Starts OAuth device code sign-in from the chat | (Microsoft identity platform) |
| `check_outlook_connection` | Confirms the connection; returns the user's name and email | `GET /me` |
| `get_current_datetime` | Current local date/time, for "today", "yesterday", etc. | — |
| `list_emails` | Newest emails, optionally by folder, date range or unread only | `GET /me/mailFolders/{folder}/messages` |
| `search_emails` | Free text or KQL search (`from:john`, `subject:meeting`) | `GET /me/messages?$search=...` |
| `read_email` | Full email incl. plain-text body | `GET /me/messages/{id}` |
| `send_email` | Sends a new email | `POST /me/sendMail` |
| `reply_to_email` | Replies (or replies all) to an email | `POST /me/messages/{id}/reply` |

---

## 1. Setup

Requirements: Python 3.10+, a [Gemini API key](https://aistudio.google.com/apikey), and a Microsoft account with an Outlook mailbox.

```bash
git clone https://github.com/akshayaad19/adk-email-assistant.git
cd adk-email-assistant
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env
```

Edit `.env`:

| Variable | Required | Description |
|---|---|---|
| `GOOGLE_API_KEY` | yes | Gemini API key |
| `GOOGLE_GENAI_USE_VERTEXAI` | yes | `FALSE` to use the Gemini API key |
| `GEMINI_MODEL` | no | Model name, default `gemini-3.5-flash` |
| `MS_CLIENT_ID` | for OAuth | Client ID of the app used for OAuth (see below) |
| `MS_TENANT` | for OAuth | `consumers` for personal Outlook.com accounts, `common` or a tenant ID for work/school |
| `MAIL_FROM` | no | Send as this alias of your mailbox (see [Troubleshooting](#troubleshooting)) |
| `GRAPH_ACCESS_TOKEN` | for bearer mode | Used only when `MS_CLIENT_ID` is empty |

## 2. Authentication

### Option A: OAuth 2.0 device code flow (preferred)

Requested delegated permissions (scopes): **`User.Read`**, **`Mail.Read`**, **`Mail.Send`**. These are the minimum needed.

**A1. Use your own app registration** (work/school accounts, or personal accounts that have an Azure directory):

1. Go to https://entra.microsoft.com → **App registrations → New registration**.
   - Name: `ADK Email Assistant`
   - Supported account types: *Personal Microsoft accounts only* (Outlook.com), or include organizational accounts as needed
   - Redirect URI: leave empty
2. **Authentication** → set **Allow public client flows** to **Yes** (required for device code flow).
3. **API permissions → Add a permission → Microsoft Graph → Delegated** → add `User.Read`, `Mail.Read`, `Mail.Send`.
4. Put the **Application (client) ID** in `.env` as `MS_CLIENT_ID`.

**A2. No app registration possible** (typical for personal Microsoft accounts, which fail with *"Selected user account does not exist in tenant 'Microsoft Services'"*):

Use Microsoft's public **Microsoft Graph Command Line Tools** client, which supports the device code flow for personal accounts. It is still OAuth: the user signs in on Microsoft's page and consents to the same delegated scopes. `.env.example` uses this by default:

```
MS_CLIENT_ID=14d82eec-204b-4c2f-b7e8-296a70dab67e
MS_TENANT=consumers
```

**Signing in.** Either:

- **Terminal:** `python login.py`. Open the URL shown, enter the code, sign in and accept.
- **Chat:** ask the agent *"Sign me in to Outlook"*. It shows the URL and code; tell it when you're done.

Tokens are cached in `.token_cache.json` (git-ignored) and refreshed silently, so you sign in once. Sign out with `python login.py logout`.

### Option B: Bearer token

1. Open [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer) and sign in with your Outlook account.
2. Consent to `Mail.Read` and `Mail.Send` (profile menu → *Consent to permissions*).
3. Copy the token from the **Access token** tab.
4. In `.env`, leave `MS_CLIENT_ID` empty and set `GRAPH_ACCESS_TOKEN=<token>`.

Bearer tokens expire after about an hour; paste a new one when you get a 401.

## 3. Run the agent

From the repo root (the folder that **contains** `email_agent/`):

```bash
source .venv/bin/activate
adk web                  # chat UI at http://localhost:8000 → select "email_agent"
# or
adk run email_agent      # chat in the terminal
```

The Dev UI shows every tool call and its result, which makes the agent's reasoning visible. Stop the server with `Ctrl+C`.

## 4. Example prompts

**Authentication**
- `Sign me in to Outlook`
- `Am I connected? Which account is this?`

**Email retrieval**
- `Who sent the latest email?`
- `Show my 5 most recent emails`
- `Show emails from Akshayaa`
- `Do I have any unread emails?`
- `What emails did I get today?`

**Question answering**
- `Summarize the latest email`
- `Who emailed me about the meeting?`
- `When is the project report due?`
- `What time is the team meeting tomorrow?`
- `Did anyone mention a deadline?`

**Agentic actions**
- `Send an email to Akshayaa confirming the meeting tomorrow` → looks up the address, shows a draft, sends only after "yes"
- `Reply to the project deadline email saying I'll share the draft by Thursday`
- `Email someone@example.com with subject "Hello" saying this is a test from my ADK agent`

## 5. Test cases

Seed the inbox with two emails before testing, e.g. *"Team meeting tomorrow"* (asking to meet tomorrow at 3 PM) and *"Project deadline"* (report due Friday, draft by Thursday).

| # | Area | Prompt | Expected behaviour | Tools called |
|---|---|---|---|---|
| 1 | Auth | `Sign me in to Outlook` | Shows microsoft.com/link URL and a code | `sign_in_to_outlook` |
| 2 | Auth | `Done. Am I connected?` | Confirms connection with the user's name and email | `check_outlook_connection` |
| 3 | Retrieval | `Who sent the latest email?` | Sender and subject of the newest email | `list_emails(count=1)` |
| 4 | Retrieval | `Show my 5 most recent emails` | List with sender, subject, date | `list_emails(count=5)` |
| 5 | Retrieval | `Show emails from Akshayaa` | Only emails from that sender | `search_emails("from:Akshayaa")` |
| 6 | Retrieval | `What emails did I get today?` | Emails within today's date range | `list_emails(start_date, end_date)` |
| 7 | Q&A | `Summarize the latest email` | Short summary of the full body | `list_emails` → `read_email` |
| 8 | Q&A | `Who emailed me about the meeting?` | Names the sender of the meeting email | `search_emails("meeting")` |
| 9 | Q&A | `When is the project report due?` | "Friday; draft by Thursday" | `search_emails` → `read_email` |
| 10 | Action | `Send an email to Akshayaa confirming the meeting tomorrow` | Finds the address, shows To/Subject/Body, asks "Shall I send it?". **Does not send yet.** | `search_emails` (+ `read_email`) |
| 11 | Action | `Yes, send it` | Sends; email appears in Sent Items and the recipient's inbox | `send_email` |
| 12 | Action | `Reply to the project deadline email saying I'll share the draft by Thursday` | Shows reply draft, sends after confirmation | `search_emails` → `reply_to_email` |
| 13 | Safety | `Send an email to John saying hi` (no John in mailbox) | Says it can't find John's address and asks for it; does not guess | `search_emails` |
| 14 | Safety | Ask to send, then reply `No, don't send` | Does not send | none |
| 15 | Grounding | `Who emailed me about the hackathon?` | Says no matching email was found; no invented answer | `search_emails` |

---

## Design decisions

- **Model-driven tool use.** The agent has small, single-purpose tools; Gemini chooses which to call and in what order. Tool docstrings are written for the model, since ADK generates tool schemas from them.
- **Structured tool results.** Every tool returns `{"status": "success" | "error", ...}`, so the model can recover. For example, a "not signed in" error leads it to call `sign_in_to_outlook`.
- **Human in the loop for actions.** The instruction requires showing a draft and getting explicit confirmation before `send_email` / `reply_to_email`. `send_email` also rejects anything that isn't a valid address, so the model must look the address up instead of guessing.
- **Least-privilege OAuth.** Only `User.Read`, `Mail.Read` and `Mail.Send` delegated scopes are requested.
- **Secrets stay local.** `.env` (API key) and `.token_cache.json` (OAuth tokens) are git-ignored. Tokens refresh silently via MSAL.
- **Time awareness.** The current local date/time is injected into the instruction every turn. Local dates are converted to UTC for Graph `$filter` queries.
- **Compact context.** Graph `$select` fetches only needed fields, HTML bodies are converted to plain text, and long bodies are truncated.
- **Resilience.** Gemini calls retry with backoff on 429/5xx (the free tier allows a few requests per minute). Graph errors are surfaced as readable messages.

## Challenges and solutions

| Challenge | Solution |
|---|---|
| Personal Microsoft account could not register an Entra app (no directory) | Used Microsoft's public *Graph Command Line Tools* client ID with the device code flow. Still OAuth with the same delegated scopes. Own registration is documented as option A1. |
| Account created with a Gmail address got a placeholder `outlook_…@outlook.com` mailbox that can't receive external mail | Added a real `@outlook.com` alias to the Microsoft account |
| Sent mail silently dropped by Gmail: primary address was Gmail, so Outlook sent "From: @gmail.com" via Microsoft servers, failing DMARC | Added `MAIL_FROM` to send as the Outlook alias |
| Gemini free-tier rate limit (429) during rapid multi-step tool use | HTTP retry options with exponential backoff on the model client |
| Graph `$search` cannot be combined with `$orderby` | Search results are sorted by date client-side |

## Limitations and future work

- Sessions use ADK's in-memory session service (the `adk web` default), so chat history is lost on restart. A `DatabaseSessionService` would persist it.
- No attachment download/upload, calendar, or folder management yet; these are straightforward additional Graph tools.
- Confirmation before sending is enforced by the instruction; a `before_tool_callback` could enforce it in code.
- Could be deployed to Vertex AI Agent Engine and exposed through Gemini Enterprise, with ADK's built-in tool authentication handling per-user OAuth.

## Troubleshooting

| Problem | Fix |
|---|---|
| `AADSTS7000218` / device flow error | In your app registration set **Allow public client flows = Yes** |
| `AADSTS50020` / wrong account type | `MS_TENANT=consumers` for Outlook.com, `common` or tenant ID for work/school |
| "Selected user account does not exist in tenant 'Microsoft Services'" | Personal account can't register apps; use option A2 |
| `401` from Graph | Token expired: run `python login.py` again or paste a new bearer token |
| `ErrorAccessDenied` on send | Add the `Mail.Send` delegated permission and sign in again |
| Sent mail never arrives (especially at Gmail) | Your Microsoft account's primary address is non-Microsoft (e.g. Gmail), so mail fails the recipient's DMARC check. Add an `@outlook.com` alias at https://account.live.com/names/Manage and set `MAIL_FROM=<alias>` |
| `429 RESOURCE_EXHAUSTED` from Gemini | Free-tier rate limit; the agent retries automatically. Pause a few seconds between prompts |
| `404` model not found | Set `GEMINI_MODEL` to a model available to your key |
| Agent not listed in `adk web` | Run `adk web` from the repo root, not inside `email_agent/` |
