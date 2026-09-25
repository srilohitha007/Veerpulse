# CHANGELOG

## `app/__init__.py`
- Changed: Replaced inline `or 0` null-coalescing loops with proper missing value aggregations across 6 places (e.g., `current_training_hours`, `total_duty_hours`) to prevent zeros from silently replacing "Missing Data".
- Reason: Strict enforcement of the "missing data vs zero" rule from Batch 1 specifications.

## `app/templates/welfare_officer_personnel_overview.html`
- Changed: Updated input placeholder from `VP1043` to `ID1234`.
- Reason: Removing any risk of "demo data leakage" into live workflows.
