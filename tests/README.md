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
