# ATOM GRID Warehouse MIS — V8 Fixed

V8 dashboard with functional sidebar navigation, stock reconciliation, snapshots, and reference-style UI.

## Fix in this build
Some warehouse Excel exports contain duplicate column headers such as `Material Code`. Pandas/Streamlit can fail when rendering such a dataframe through PyArrow. This build removes duplicate headers defensively when loading each sheet and again before rendering tables.

## Run
`streamlit run app.py`
