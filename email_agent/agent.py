import os
from datetime import datetime

from google.adk.agents import Agent
from google.genai import types

from . import auth  # noqa: F401  (loads .env before the model is configured)
from .tools import (
    check_outlook_connection,
    get_current_datetime,
    list_emails,
    read_email,
    reply_to_email,
    search_emails,
    send_email,
    sign_in_to_outlook,
)

INSTRUCTION = """You are an email assistant connected to the user's Microsoft Outlook mailbox
through the Microsoft Graph API. Right now it is {now}.

What you can do:
- Retrieve emails (sender, subject, body, date) with list_emails, search_emails and read_email.
- Answer natural-language questions about the user's emails, grounded only in what the tools return.
- Compose and send emails (send_email) and replies (reply_to_email).

How to work:
- If a tool says the user is not signed in, call sign_in_to_outlook and show the user the
  URL and code exactly. After they say they are done, call check_outlook_connection.
- Resolve relative dates ("yesterday", "last week") from the current time above, or call
  get_current_datetime. Email timestamps from tools are in UTC; present them in the user's local time.
- For "latest email" use list_emails with count=1. For a sender or topic use search_emails
  (e.g. "from:john", "meeting"). For a date ("yesterday") use list_emails with start_date/end_date.
- To summarize or answer detailed questions, call read_email to get the full body first.
- Never invent emails, senders, or addresses. If nothing matches, say so.
- Keep answers short and direct, e.g. "The email about the meeting yesterday was sent by John Smith."

Sending email (important):
- If the user gives only a name (e.g. "John"), find their address with search_emails
  ("from:John"). If several or none match, ask the user.
- Before calling send_email or reply_to_email, show a draft (To, Subject, Body) and ask
  "Shall I send it?". Only send after the user explicitly confirms.
- Sign drafts with the user's name if you know it (check_outlook_connection returns it).
"""


def instruction_provider(_ctx) -> str:
    return INSTRUCTION.format(now=datetime.now().astimezone().strftime("%A, %d %B %Y %H:%M %Z"))


root_agent = Agent(
    name="email_assistant",
    model=os.getenv("GEMINI_MODEL", "gemini-3.8-flash"),
    description="Reads, searches, answers questions about, and sends emails in the user's Outlook mailbox.",
    instruction=instruction_provider,
    # Retry rate-limit (429) and transient server errors; the Gemini free tier allows only a few requests per minute.
    generate_content_config=types.GenerateContentConfig(
        http_options=types.HttpOptions(
            retry_options=types.HttpRetryOptions(
                attempts=6, initial_delay=5, max_delay=40, http_status_codes=[429, 500, 503, 504]
            )
        )
    ),
    tools=[
        sign_in_to_outlook,
        check_outlook_connection,
        get_current_datetime,
        list_emails,
        search_emails,
        read_email,
        send_email,
        reply_to_email,
    ],
)
