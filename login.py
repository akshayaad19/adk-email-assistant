"""Sign in to Outlook (OAuth device code flow) before starting the agent.

    python login.py          # sign in
    python login.py logout   # remove cached tokens
"""

import sys

from email_agent import auth

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "logout":
        auth.logout()
        print("Signed out (token cache removed).")
    elif not auth._client_id():
        print("MS_CLIENT_ID is not set; the agent will use GRAPH_ACCESS_TOKEN from .env instead.")
    else:
        auth.login_interactive()
