
from msal import PublicClientApplication

CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"

app = PublicClientApplication(
    CLIENT_ID,
    authority="https://login.microsoftonline.com/organizations"
)

result = app.acquire_token_interactive(
    scopes=["User.Read"]
)

if "access_token" in result:
    print("SUCCESS")
    print(result["account"]["username"])
else:
    print(result)

