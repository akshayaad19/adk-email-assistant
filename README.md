# Outlook Email Assistant (Google ADK + Microsoft Graph)

An AI agent built with the [Google Agent Development Kit (ADK)](https://google.github.io/adk-docs/) that connects to a Microsoft Outlook mailbox through the [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/api/overview). It can:

- **Retrieve emails** (sender, subject, body, date)
- **Answer natural-language questions** about them ("Who emailed me about the meeting yesterday?")
- **Take actions**: compose and send new emails, reply to emails

It runs locally, uses Gemini as the model, and authenticates with **OAuth 2.0 (device code flow via MSAL)**, with a **bearer token** fallback.

## Architecture

```
User ──► ADK web UI / CLI (adk web | adk run)
            │
            ▼
   email_assistant (LlmAgent, Gemini)
            │  function calls
            ▼
   tools.py ── sign_in_to_outlook, check_outlook_connection, get_current_datetime,
               list_emails, search_emails, read_email, send_email, reply_to_email
            │
            ▼
   graph_client.py ──HTTPS──► Microsoft Graph /me/messages, /me/sendMail, ...
            ▲
   auth.py ── MSAL device code flow, token cache (.token_cache.json), silent refresh
```

| File | Purpose |
|---|---|
| `email_agent/agent.py` | ADK `root_agent`: model, instructions (incl. current date/time), tool list |
| `email_agent/tools.py` | Tool definitions; ADK builds the tool schemas from signatures and docstrings |
| `email_agent/graph_client.py` | Microsoft Graph REST calls (list, search, read, send, reply) |
| `email_agent/auth.py` | OAuth device code flow (MSAL), token cache, bearer token fallback |
| `login.py` | Terminal sign-in / sign-out helper |

**Design notes**
- Every tool returns `{"status": "success" | "error", ...}` so the model can recover from errors (e.g. by starting sign-in when a token is missing).
- The agent always shows a draft and asks for confirmation before `send_email` or `reply_to_email`.
- If the user names a recipient ("John"), the agent looks up the address in the mailbox with `search_emails("from:John")`.
- Relative dates ("yesterday") are resolved from the current local time, which goes into the instruction on every turn. Local dates are converted to UTC for Graph filters.

## 1. Setup

Requirements: Python 3.10+, a Gemini API key, and a Microsoft account with an Outlook mailbox.

```bash
git clone <this-repo-url> adk-email-assistant
cd adk-email-assistant
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Edit `.env` and set `GOOGLE_API_KEY` (from https://aistudio.google.com/apikey).

## 2. Authentication

### Option A: OAuth (preferred)

1. **Register an app** in the Microsoft Entra admin center: https://entra.microsoft.com → *Identity → Applications → App registrations → New registration*.
   - Name: `ADK Email Assistant`
   - Supported account types: **Accounts in any organizational directory and personal Microsoft accounts** (or *Personal Microsoft accounts only* for Outlook.com)
   - Redirect URI: leave empty
2. In the app's **Authentication** page, set **Allow public client flows → Yes** and save. Device code flow needs this.
3. In **API permissions → Add a permission → Microsoft Graph → Delegated**, add `User.Read`, `Mail.Read`, `Mail.Send`.
4. Copy the **Application (client) ID** into `.env`:
   ```
   MS_CLIENT_ID=<your-client-id>
   MS_TENANT=consumers     # personal Outlook.com/Hotmail; use "common" or your tenant ID for work/school
   ```
5. Sign in, either:
   - **from the terminal**: `python login.py`, then open the URL shown and enter the code, or
   - **from the chat**: ask the agent *"Sign me in to Outlook"*. It shows the URL and code, and you tell it when you're done.

   The token is cached in `.token_cache.json` (git-ignored) and refreshed automatically. To sign out: `python login.py logout`.

> **No app registration?** Personal Microsoft accounts often can't register apps, because they have no
> directory ("Selected user account does not exist in tenant 'Microsoft Services'"). In that case skip
> steps 1–3 and use Microsoft's public **Microsoft Graph Command Line Tools** client, which supports the
> device code flow for personal accounts. This is still OAuth; you consent to the same delegated permissions:
> ```
> MS_CLIENT_ID=14d82eec-204b-4c2f-b7e8-296a70dab67e
> MS_TENANT=consumers
> ```
> `.env.example` uses this by default.

### Option B: Bearer token

1. Open [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer) and sign in with your Outlook account.
2. Click your profile → **Consent to permissions** and consent to `Mail.Read` and `Mail.Send`.
3. Open the **Access token** tab and copy the token.
4. In `.env`, leave `MS_CLIENT_ID` empty and set `GRAPH_ACCESS_TOKEN=<token>`.

These tokens expire after about an hour. Paste a new one when you get a 401.

## 3. Run the agent

From the project root (the folder that **contains** `email_agent/`):

```bash
adk web          # browser chat UI at http://localhost:8000 → select "email_agent"
# or
adk run email_agent   # terminal chat
```

## 4. Example prompts

**Authentication**
- `Sign me in to Outlook`
- `Am I connected? Which account?`

**Email retrieval**
- `Who sent the latest email?`
- `Show my 5 most recent emails`
- `Show emails from john`
- `Do I have any unread emails?`

**Question answering**
- `Summarize the latest email`
- `Who emailed me about the meeting yesterday?`
- `What did Sarah say about the project deadline?`
- `Did anyone send me an invoice this week?`

**Agentic actions**
- `Send an email to John confirming the meeting tomorrow` → the agent finds John's address, shows a draft, and sends after you confirm
- `Reply to the latest email saying thanks, I'll review it today`
- `Email test@example.com with subject "Hello" saying this is a test from my ADK agent`

## Troubleshooting

| Problem | Fix |
|---|---|
| `AADSTS7000218` / device flow error | Set **Allow public client flows = Yes** in the app's Authentication page |
| `AADSTS50020` / wrong account type | Use `MS_TENANT=consumers` for Outlook.com accounts, `common` or your tenant ID for work/school |
| `401` from Graph | Token expired. Run `python login.py` again or paste a new bearer token |
| `ErrorAccessDenied` on send | Add the `Mail.Send` delegated permission and sign in again |
| Sent mail never arrives (esp. at Gmail) | Your Microsoft account's primary address is a non-Microsoft one (e.g. Gmail), so mail is sent "from" Gmail via Microsoft servers and fails the recipient's DMARC check. Add an `@outlook.com` alias at account.live.com/names/Manage and set `MAIL_FROM=<alias>` in `.env` |
| Agent not listed in `adk web` | Run `adk web` from the repo root, not from inside `email_agent/` |
