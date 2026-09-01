# ATOM GRID Warehouse MIS — V8 Final Fix

This version fixes the pandas merge error caused by Material Code being read as int64 in one dataset and string in another. Reconciliation keys are normalized to strings before snapshot merges.

It also retains the Daily Snapshot minimum-stock filter and functional navigation.
