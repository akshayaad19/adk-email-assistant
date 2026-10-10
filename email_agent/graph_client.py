"""Thin wrapper around the Microsoft Graph mail endpoints."""

import os
import re
from html import unescape

import requests

from .auth import get_access_token

GRAPH = "https://graph.microsoft.com/v1.0"
SUMMARY_FIELDS = "id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments"
MAX_BODY_CHARS = 8000


class GraphError(Exception):
    pass


def _request(method, path, params=None, json=None, headers=None):
    hdrs = {"Authorization": f"Bearer {get_access_token()}"}
    if headers:
        hdrs.update(headers)
    resp = requests.request(method, f"{GRAPH}{path}", params=params, json=json, headers=hdrs, timeout=30)
    if resp.status_code == 401:
        raise GraphError("Microsoft Graph rejected the token (401). Sign in again or refresh the bearer token.")
    if not resp.ok:
        try:
            err = resp.json().get("error", {})
            msg = f"{err.get('code')}: {err.get('message')}"
        except ValueError:
            msg = resp.text[:300]
        raise GraphError(f"Graph API error {resp.status_code}: {msg}")
    return resp.json() if resp.content else {}


def _addr(recipient):
    ea = (recipient or {}).get("emailAddress", {})
    name, address = ea.get("name"), ea.get("address")
    return f"{name} <{address}>" if name and name != address else (address or name or "unknown")


def _summary(msg):
    return {
        "id": msg["id"],
        "from": _addr(msg.get("from")),
        "to": [_addr(r) for r in msg.get("toRecipients", [])],
        "subject": msg.get("subject") or "(no subject)",
        "received": msg.get("receivedDateTime"),
        "preview": msg.get("bodyPreview", ""),
        "is_read": msg.get("isRead"),
        "has_attachments": msg.get("hasAttachments"),
    }


def _to_text(body):
    content = body.get("content", "")
    if body.get("contentType") == "html":
        content = re.sub(r"(?is)<(script|style).*?</\1>", "", content)
        content = re.sub(r"(?i)<br\s*/?>|</p>|</div>", "\n", content)
        content = unescape(re.sub(r"<[^>]+>", "", content))
    content = re.sub(r"\n\s*\n+", "\n\n", content).strip()
    return content[:MAX_BODY_CHARS] + ("\n...[truncated]" if len(content) > MAX_BODY_CHARS else "")


def me():
    return _request("GET", "/me", params={"$select": "displayName,mail,userPrincipalName"})


def list_messages(folder="inbox", top=10, start=None, end=None, unread_only=False):
    filters = []
    if start:
        filters.append(f"receivedDateTime ge {start}")
    if end:
        filters.append(f"receivedDateTime lt {end}")
    if unread_only:
        filters.append("isRead eq false")
    params = {"$top": top, "$select": SUMMARY_FIELDS, "$orderby": "receivedDateTime desc"}
    if filters:
        params["$filter"] = " and ".join(filters)
    data = _request("GET", f"/me/mailFolders/{folder}/messages", params=params)
    return [_summary(m) for m in data.get("value", [])]


def search_messages(query, top=10):
    # $search uses KQL, e.g. "from:john", "subject:meeting", or free text.
    # It cannot be combined with $orderby, so sort by date afterwards.
    params = {"$search": f'"{query}"', "$top": top, "$select": SUMMARY_FIELDS}
    data = _request("GET", "/me/messages", params=params)
    msgs = [_summary(m) for m in data.get("value", [])]
    return sorted(msgs, key=lambda m: m["received"] or "", reverse=True)


def get_message(message_id):
    msg = _request(
        "GET",
        f"/me/messages/{message_id}",
        params={"$select": SUMMARY_FIELDS + ",ccRecipients,body"},
        headers={"Prefer": 'outlook.body-content-type="text"'},
    )
    out = _summary(msg)
    out["cc"] = [_addr(r) for r in msg.get("ccRecipients", [])]
    out["body"] = _to_text(msg.get("body", {}))
    out.pop("preview", None)
    return out


def send_mail(to, subject, body, cc=None):
    message = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body},
        "toRecipients": [{"emailAddress": {"address": a}} for a in to],
    }
    if cc:
        message["ccRecipients"] = [{"emailAddress": {"address": a}} for a in cc]
    # Optional: send as a specific alias of the mailbox (e.g. when the account's primary
    # address is a non-Microsoft one like Gmail, which recipients' DMARC checks would reject).
    sender = os.getenv("MAIL_FROM", "").strip()
    if sender:
        message["from"] = {"emailAddress": {"address": sender}}
    _request("POST", "/me/sendMail", json={"message": message, "saveToSentItems": True})


def reply(message_id, comment, reply_all=False):
    action = "replyAll" if reply_all else "reply"
    _request("POST", f"/me/messages/{message_id}/{action}", json={"comment": comment})
