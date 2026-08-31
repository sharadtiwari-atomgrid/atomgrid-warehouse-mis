# ATOM GRID Warehouse MIS — Final No Login Version

This version has **no login and no authentication**.

Anyone who has the Render URL can open the dashboard and upload the latest warehouse MIS Excel.

## Render settings

Build Command:
```text
pip install -r requirements.txt
```

Start Command:
```text
streamlit run app.py --server.port $PORT --server.address 0.0.0.0
```

No Google OAuth credentials, Streamlit secrets, or Render environment variables are required.

## Daily operation

1. Open the Render URL.
2. Upload the latest warehouse MIS Excel.
3. Review Overview, Stock, Inward, Outward, Ageing and Exceptions.
4. Use New Materials to identify products not in the approved master.
5. Download filtered stock or exception CSVs when required.

## Important

Delete/ignore any old Google OAuth secret file or authentication configuration from Render. This application no longer uses it.
