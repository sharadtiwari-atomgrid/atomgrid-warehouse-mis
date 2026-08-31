# ATOM GRID Warehouse MIS — Final Live Dashboard

## What this version does
Upload the latest warehouse MIS Excel and get live visibility into:
- Stock in Hand
- Physical stock and variance
- Inward
- Outward
- Material-wise stock
- Batch-wise stock
- Ageing buckets
- Low-stock exceptions
- New/unmapped materials
- Pending orders
- Vehicle indent
- Cancelled invoices
- CSV downloads

## Local test
1. Install Python 3.10+.
2. `pip install -r requirements.txt`
3. `streamlit run app.py`
4. Upload the actual warehouse MIS.

## Publish
Recommended: Streamlit Community Cloud.
1. Create a private GitHub repository.
2. Add `app.py` and `requirements.txt`.
3. Create a Streamlit Community Cloud app from the repository.
4. Add the contents of `.streamlit/secrets.toml` in the app Secrets settings.
5. Configure Google OAuth/OIDC with the callback URL shown in the app settings.
6. Keep the repository private.
7. Share the dashboard URL only with authorised users.

## Access control
The app is designed for Google sign-in and checks that the signed-in email ends with `@atomgrid.in`.
Do not rely on a typed email address as authentication. Use Google OAuth/OIDC in production.

## Material Master
Upload an approved Material Master in the sidebar. Unknown material names/codes are flagged as NEW / NOT MAPPED. The app never silently maps an unknown material to an existing catalogue item.

## Important production note
This upload-driven version processes the file currently uploaded in the browser. For persistent history, audit logs, scheduled ingestion, and multi-user shared material master, the next architecture should add a database/Drive ingestion layer. The dashboard UI can remain the same.
