# Navigation performance regression tests

Run `npm ci` followed by `npm test` from the repository root (Node 20+).
Tests load the actual `gas/Index.html` in jsdom and control only the Apps Script
transport. All fixtures are synthetic. No production orders are created.

Coverage includes initial loads, instant revisits, request deduplication,
out-of-order responses, refresh failures and retry, search focus/caret retention,
write invalidation, logout, dashboard refresh, and parallel settings reads.
The deployment workflow runs these tests before pushing to Apps Script.

The historical 285-assertion payment suite and its private sales fixtures were
recovered locally and run against both the old and changed frontend. One
pre-existing assertion expected `billTo_(salesman, mode, t)` exactly; the current
backend also accepts `branch`. The local assertion was updated to accept that
existing optional parameter. All 285 assertions then passed on the changed UI.
Private historical fixtures are deliberately excluded from this repository.

Snapshots are held only in memory for the current login, keyed by query, with a
15-second freshness window and at most 20 entries. Expired entries remain visible
while revalidating and show their last successful update time. Explicit refresh
always revalidates. Mutations invalidate snapshots before and after the request,
including failures with an uncertain server outcome. Read epochs prevent a
late response from undoing that invalidation. No writes are cached or retried.

For diagnosis, `RPC_TIMINGS` holds the last 50 operation names, elapsed milliseconds,
and transport-success flags in memory only. It contains no PINs, request arguments,
response data, or external telemetry. It is cleared on logout.

Second pass: admin/partner login prefetches orders and delivery together through
`getNavigationOrders`, which reads ORDERS once for both lists. A first click joins
the in-flight request or uses its completed snapshot. Driver login only requests
the pending list. Delivery now uses the same session snapshot handling, keeping
selected IDs only while they remain pending. Backend date formatting is memoized
by timestamp within each execution, preserving the script timezone.

`navigation-backend.cjs` runs the real backend with a synthetic sheet: it checks
one read for both lists, old pending orders outside the recent 300, void/payment
rules, and timezone-boundary dates. On its fixture, 640 date conversions become
one. This is an operation count, not a claim about production milliseconds.

Third pass: the navigation bundle also computes the default dashboard from the
same order table. Admin settings prefetch uses one authenticated endpoint instead
of two, caches only within the login, and invalidates on account/driver changes.
Unsaved settings inputs survive tab revisits. `startup-backend.cjs` verifies
authorization, one USERS read for settings, lazy spreadsheet connection, and
bootstrap's reuse of BRANCH data (including historical inactive brands).
The PIN page's fonts load without blocking render; the Vercel wrapper removes
its extra 400ms delay and shortens the splash fade to 150ms.

Invoice selection: `invoice-merge.cjs` exercises the backend with synthetic tables,
covering selected order mapping, customer/month/billing-mode validation, duplicate
issuance, void/reissue, immutable reprint snapshots, changed-total confirmation,
legacy invoices and PDF failures. `invoice-merge-ui.cjs` verifies selection,
disabled issued orders, confirmation totals and issued versus pending PDF routing.
New invoices persist ORDER_IDS and SNAPSHOT_JSON together in the INVOICE row under
the script lock. Old invoices reserve their original customer/month or order key;
their first reprint snapshots current details only if total and count match the
record. Historical original item contents cannot be reconstructed from old metadata.
Selection is limited to one customer, billing mode and month, at most 300 orders;
snapshots above 45,000 characters are rejected before allocating a number.
The private regression harness invalidates UI snapshots after direct MOCK mutations
and awaits the dashboard read instead of assuming a fixed delay is sufficient.
