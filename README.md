# ATOM GRID Warehouse MIS — Google Login

This version uses Google OIDC authentication. There are **no application passwords**.

## Access rule
Only authenticated Google accounts ending in `@atomgrid.in` are allowed.

## Render settings

Build Command:
```text
pip install -r requirements.txt
```

Start Command:
```text
streamlit run app.py --server.port $PORT --server.address 0.0.0.0
```

## Google OAuth

Use this exact Authorized Redirect URI:

```text
https://atomgrid-warehouse-mis.onrender.com/oauth2callback
```

Streamlit's Google OIDC provider uses:

```text
https://accounts.google.com/.well-known/openid-configuration
```

Create a Web application OAuth client in Google Cloud and keep the Client ID and Client Secret out of GitHub.

## Streamlit secrets

Configure the following under `[auth]` in Streamlit secrets:

```toml
[auth]
redirect_uri = "https://atomgrid-warehouse-mis.onrender.com/oauth2callback"
cookie_secret = "GENERATE_A_LONG_RANDOM_SECRET"
client_id = "YOUR_GOOGLE_CLIENT_ID"
client_secret = "YOUR_GOOGLE_CLIENT_SECRET"
server_metadata_url = "https://accounts.google.com/.well-known/openid-configuration"
```

Do not commit the real values to GitHub.

## Daily operation

1. Open the Render URL.
2. Click Sign in with Google.
3. Use an `@atomgrid.in` account.
4. Upload the latest warehouse MIS Excel.
5. Review stock, inward, outward, ageing, exceptions, new materials and other MIS tabs.

## Important

This application checks the authenticated Google email domain. It does not merely ask the user to type an email address.
