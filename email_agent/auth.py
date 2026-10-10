"""Microsoft Graph authentication.

Two modes:
  * OAuth 2.0 device code flow via MSAL (preferred). Tokens are cached in
    .token_cache.json and refreshed silently, so you sign in once.
  * Bearer token fallback: if MS_CLIENT_ID is not set, the token in
    GRAPH_ACCESS_TOKEN is used as-is.

Run `python login.py` to sign in from the terminal before starting
the agent, or ask the agent to "sign me in to Outlook" from the chat.
"""

import os
import threading
from pathlib import Path

import msal
from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(PROJECT_ROOT / ".env")

SCOPES = ["User.Read", "Mail.Read", "Mail.Send"]
CACHE_PATH = PROJECT_ROOT / ".token_cache.json"


class AuthError(Exception):
    pass


_cache = msal.SerializableTokenCache()
if CACHE_PATH.exists():
    _cache.deserialize(CACHE_PATH.read_text())

_app = None
_pending_flow = None  # device flow started from chat, completed in a background thread
_flow_lock = threading.Lock()


def _client_id():
    return os.getenv("MS_CLIENT_ID", "").strip()


def _get_app():
    global _app
    if _app is None:
        if not _client_id():
            raise AuthError("MS_CLIENT_ID is not set in .env.")
        tenant = os.getenv("MS_TENANT", "consumers").strip() or "consumers"
        _app = msal.PublicClientApplication(
            _client_id(),
            authority=f"https://login.microsoftonline.com/{tenant}",
            token_cache=_cache,
        )
    return _app


def _save_cache():
    if _cache.has_state_changed:
        CACHE_PATH.write_text(_cache.serialize())


def get_access_token():
    """Return a valid Graph access token or raise AuthError."""
    if not _client_id():
        token = os.getenv("GRAPH_ACCESS_TOKEN", "").strip()
        if token:
            return token
        raise AuthError(
            "Not authenticated: set MS_CLIENT_ID (OAuth) or GRAPH_ACCESS_TOKEN (bearer) in .env."
        )

    app = _get_app()
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        _save_cache()
        if result and "access_token" in result:
            return result["access_token"]
    raise AuthError("Not signed in to Outlook yet. Use the sign-in tool or run `python login.py`.")


def signed_in_account():
    """Username of the cached account, or None."""
    if not _client_id():
        return "bearer-token" if os.getenv("GRAPH_ACCESS_TOKEN", "").strip() else None
    accounts = _get_app().get_accounts()
    return accounts[0]["username"] if accounts else None


def start_device_flow():
    """Start a device code flow and finish it in the background.

    Returns the flow dict, whose "message" tells the user where to go and
    which code to enter.
    """
    global _pending_flow
    app = _get_app()
    with _flow_lock:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise AuthError(f"Could not start sign-in: {flow.get('error_description', flow)}")
        _pending_flow = flow

    def _complete():
        global _pending_flow
        app.acquire_token_by_device_flow(flow)  # blocks until the user finishes or the code expires
        _save_cache()
        with _flow_lock:
            if _pending_flow is flow:
                _pending_flow = None

    threading.Thread(target=_complete, daemon=True).start()
    return flow


def sign_in_pending():
    return _pending_flow is not None


def login_interactive():
    """Terminal sign-in: print the device code and wait for completion."""
    app = _get_app()
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise AuthError(f"Could not start sign-in: {flow.get('error_description', flow)}")
    print(flow["message"], flush=True)
    result = app.acquire_token_by_device_flow(flow)
    _save_cache()
    if "access_token" not in result:
        raise AuthError(result.get("error_description", "Sign-in failed."))
    print(f"Signed in as {result.get('id_token_claims', {}).get('preferred_username', 'unknown')}.")


def logout():
    if CACHE_PATH.exists():
        CACHE_PATH.unlink()

