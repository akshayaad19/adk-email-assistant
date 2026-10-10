"""Tools exposed to the ADK agent.

ADK reads each function's signature and docstring to build the tool schema the
model sees, so the docstrings below are written for the model.
Every tool returns a dict with a "status" key ("success" or "error").
"""

import re
from datetime import datetime, timedelta, timezone

from . import auth, graph_client
from .auth import AuthError
from .graph_client import GraphError

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def _error(e):
    return {"status": "error", "error_message": str(e)}


def _local_day_to_utc(day):
    """'YYYY-MM-DD' at local midnight -> Graph-friendly UTC timestamp."""
    local = datetime.strptime(day, "%Y-%m-%d").astimezone()
    return local.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def sign_in_to_outlook() -> dict:
    """Signs the user in to their Microsoft Outlook account using OAuth (device code flow).

    Call this when another tool reports the user is not signed in, or when the user
    asks to sign in / connect their mailbox. Show the returned instructions to the
    user exactly: they must open the URL and enter the code.

    Returns:
        dict: status, plus "instructions" (URL and code) or "signed_in_as".
    """
    try:
        if not auth._client_id():
            account = auth.signed_in_account()
            if account:
                return {"status": "success", "signed_in_as": "bearer token from .env (GRAPH_ACCESS_TOKEN)"}
            return _error("No MS_CLIENT_ID or GRAPH_ACCESS_TOKEN configured in .env.")
        account = auth.signed_in_account()
        if account:
            try:
                auth.get_access_token()
                return {"status": "success", "signed_in_as": account}
            except AuthError:
                pass
        flow = auth.start_device_flow()
        return {
            "status": "success",
            "instructions": flow["message"],
            "verification_url": flow.get("verification_uri"),
            "code": flow.get("user_code"),
            "note": "After the user finishes in the browser, call check_outlook_connection to confirm.",
        }
    except AuthError as e:
        return _error(e)


def check_outlook_connection() -> dict:
    """Checks whether the agent is connected to the user's Outlook mailbox.

    Returns:
        dict: status, and the connected user's name and email address if connected.
    """
    try:
        user = graph_client.me()
        return {
            "status": "success",
            "connected": True,
            "name": user.get("displayName"),
            "email": user.get("mail") or user.get("userPrincipalName"),
        }
    except (AuthError, GraphError) as e:
        return {
            "status": "success",
            "connected": False,
            "sign_in_in_progress": auth.sign_in_pending(),
            "detail": str(e),
        }


def list_emails(
    count: int = 10,
    folder: str = "inbox",
    start_date: str = "",
    end_date: str = "",
    unread_only: bool = False,
) -> dict:
    """Lists emails, newest first, with sender, subject, date and a short preview.

    Use this for "latest email", "recent emails", "unread emails", or emails in a
    date range (e.g. yesterday: start_date=yesterday, end_date=today).

    Args:
        count: How many emails to return (1-50). Default 10.
        folder: Mail folder: "inbox", "sentitems", "drafts", "archive", or "junkemail".
        start_date: Optional. Only emails received on/after this local date, format YYYY-MM-DD.
        end_date: Optional. Only emails received before this local date (exclusive), format YYYY-MM-DD.
        unread_only: If true, only return unread emails.

    Returns:
        dict: status and "emails", a list of {id, from, to, subject, received (UTC), preview, is_read}.
    """
    try:
        emails = graph_client.list_messages(
            folder=folder,
            top=max(1, min(int(count), 50)),
            start=_local_day_to_utc(start_date) if start_date else None,
            end=_local_day_to_utc(end_date) if end_date else None,
            unread_only=unread_only,
        )
        return {"status": "success", "count": len(emails), "emails": emails}
    except (AuthError, GraphError, ValueError) as e:
        return _error(e)


def search_emails(query: str, count: int = 10) -> dict:
    """Searches all mail folders, returning matches newest first.

    Supports free text ("meeting", "invoice") and KQL-style filters such as
    "from:john", "from:john@example.com", "subject:meeting", "to:alice".
    Use this to find emails from a specific sender or about a topic, and to look
    up someone's email address by name.

    Args:
        query: Search text or KQL expression.
        count: Maximum number of results (1-50). Default 10.

    Returns:
        dict: status and "emails", a list of {id, from, to, subject, received (UTC), preview, is_read}.
    """
    try:
        emails = graph_client.search_messages(query.replace('"', ""), top=max(1, min(int(count), 50)))
        return {"status": "success", "count": len(emails), "emails": emails}
    except (AuthError, GraphError) as e:
        return _error(e)


def read_email(message_id: str) -> dict:
    """Reads the full content of one email, including its plain-text body.

    Use this to summarize an email or answer detailed questions about it.
    Get the message_id from list_emails or search_emails.

    Args:
        message_id: The "id" of the email.

    Returns:
        dict: status and "email" with id, from, to, cc, subject, received, body.
    """
    try:
        return {"status": "success", "email": graph_client.get_message(message_id)}
    except (AuthError, GraphError) as e:
        return _error(e)


def send_email(to: list[str], subject: str, body: str, cc: list[str] | None = None) -> dict:
    """Composes and sends a new email from the user's Outlook account.

    Only call this after showing the user the recipient(s), subject and body and
    getting their explicit confirmation to send.

    Args:
        to: Recipient email addresses (full addresses, e.g. "john@example.com").
        subject: Subject line.
        body: Plain-text body of the email.
        cc: Optional CC email addresses.

    Returns:
        dict: status and a confirmation message.
    """
    bad = [a for a in list(to) + list(cc or []) if not EMAIL_RE.match(a.strip())]
    if not to:
        return _error("At least one recipient is required.")
    if bad:
        return _error(f"Not valid email addresses: {bad}. Look up the address with search_emails or ask the user.")
    try:
        graph_client.send_mail([a.strip() for a in to], subject, body, [a.strip() for a in cc or []])
        return {"status": "success", "message": f"Email '{subject}' sent to {', '.join(to)}."}
    except (AuthError, GraphError) as e:
        return _error(e)


def reply_to_email(message_id: str, body: str, reply_all: bool = False) -> dict:
    """Replies to an existing email (keeps the thread and original subject).

    Only call this after showing the user the reply text and getting their
    explicit confirmation to send.

    Args:
        message_id: The "id" of the email being replied to.
        body: Plain-text reply message.
        reply_all: If true, reply to all recipients instead of only the sender.

    Returns:
        dict: status and a confirmation message.
    """
    try:
        graph_client.reply(message_id, body, reply_all=reply_all)
        return {"status": "success", "message": "Reply sent."}
    except (AuthError, GraphError) as e:
        return _error(e)


def get_current_datetime() -> dict:
    """Returns the current local date, time, weekday and timezone, plus yesterday/tomorrow dates.

    Use this to resolve relative dates such as "today", "yesterday", "last week".
    """
    now = datetime.now().astimezone()
    return {
        "status": "success",
        "now": now.isoformat(timespec="seconds"),
        "weekday": now.strftime("%A"),
        "timezone": str(now.tzinfo),
        "today": now.strftime("%Y-%m-%d"),
        "yesterday": (now - timedelta(days=1)).strftime("%Y-%m-%d"),
        "tomorrow": (now + timedelta(days=1)).strftime("%Y-%m-%d"),
    }
