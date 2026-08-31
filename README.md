# ATOM GRID Warehouse MIS — Google Login (Fixed)

This version uses Google OIDC authentication. There are **no application passwords**.

## Access rule
Only authenticated Google accounts ending in `@atomgrid.in` are allowed.

## Important fix
This app uses the **default/unnamed Streamlit Google OIDC provider**, so the code calls:

```python
st.login()
```

not `st.login("google")`.

The app also safely handles the case where authentication secrets have not yet been configured, instead of crashing with:

`AttributeError: st.user has no attribute 'is_logged_in'`.

## Render settings

Build Command:
```text
pip install -r requirements.txt
```

Start Command:
```text
streamlit run app.py --server.port $PORT --server.address 0.0.0.0
```

## Google OAuth redirect URI

In Google Cloud, the Web application OAuth client must have this exact Authorized Redirect URI:

```text
https://atomgrid-warehouse-mis.onrender.com/oauth2callback
```

## Render Streamlit secrets

The app requires these values under `[auth]`:

```toml
[auth]
redirect_uri = "https://atomgrid-warehouse-mis.onrender.com/oauth2callback"
cookie_secret = "GENERATE_A_LONG_RANDOM_SECRET"
client_id = "YOUR_GOOGLE_CLIENT_ID"
client_secret = "YOUR_GOOGLE_CLIENT_SECRET"
server_metadata_url = "https://accounts.google.com/.well-known/openid-configuration"
```

Do not commit real credentials to GitHub.

## Daily operation

1. Open the Render URL.
2. Click Sign in with Google.
3. Use an `@atomgrid.in` account.
4. Upload the latest warehouse MIS Excel.
5. Review stock, inward, outward, ageing, exceptions and new materials.

## Access restriction

The application verifies the authenticated Google account email and only permits the `@atomgrid.in` domain. Typing an email address is not used for authentication.
