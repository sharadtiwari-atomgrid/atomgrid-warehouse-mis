# ATOM GRID Warehouse MIS — V8

Streamlit dashboard styled to match the supplied Warehouse MIS Stock Reconciliation reference image.

## Deploy
- `app.py`
- `requirements.txt`

Render: create a Web Service from GitHub, runtime Python, build `pip install -r requirements.txt`, start `streamlit run app.py --server.address 0.0.0.0 --server.port $PORT`.

## Reconciliation
Expected Stock uses the explicit AG MIS / expected stock field when present. If no explicit expected stock exists, it falls back to Opening + Inward - Outward. Variance = Warehouse Actual - Expected.
