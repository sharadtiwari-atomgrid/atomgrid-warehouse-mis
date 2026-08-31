# ATOM GRID Warehouse MIS V6

V6 focuses on stock discrepancy and adds movement-aware reconciliation plus daily snapshots.

## Logic

For each Material + Batch:

**Expected Stock = Opening Stock + Today's Inward - Today's Outward**

The expected stock is compared with the warehouse/system stock.

A separate physical check compares:

**Physical Stock - System Stock**

## Daily snapshots

One CSV snapshot is stored per EOD date and is not overwritten. This is intended for the first 7–14 days of manual validation against the warehouse's EOD report.

**Important:** snapshots are stored on the service filesystem. Render can reset ephemeral storage during some redeploy/restart events. Do not treat this as permanent history yet. After validation, move snapshots to a database/persistent storage.

## Render

Build:
`pip install -r requirements.txt`

Start:
`streamlit run app.py --server.port $PORT --server.address 0.0.0.0`

No authentication.

## Daily process

1. Upload warehouse Excel.
2. Check Discrepancy Control.
3. Check Movement Reconciliation when available.
4. Open Daily Snapshot.
5. Compare with manual EOD report.
6. Keep both processes running for 7–14 days before making the dashboard the primary control.
