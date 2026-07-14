# Apps Script Automation

## Pickup-time auto completion

`autoCompleteOrdersByPickupTime` checks pending orders and completes an entire
order when its pickup time is reached.

Behavior:

- Runs every five minutes after the trigger is installed.
- Only processes rows whose status is `待製作`.
- Groups rows by `system_order_id` and completes the entire order together.
- Rechecks status under a script lock to prevent duplicate status updates.
- Does not send a LINE notification when an order is completed.
- Skips empty or unsafe pickup-time values.
- Skips orders overdue by more than 12 hours so stale records are not changed
  unexpectedly when automation is enabled for the first time.

Supported legacy pickup-time examples include:

- `2026/07/13 18:30`
- `2026-07-13 18:30`
- `18:30` (uses the order creation date)
- `晚上6點半` (uses the order creation date)
- `明天 18:30`

New AI-parsed orders store pickup time as `yyyy/MM/dd HH:mm` in the
`Asia/Taipei` time zone. A customer-requested time takes priority. When the
customer omits pickup time, the queue-based system estimate is written to the
same `pickup_time` column before the order is saved.

The queue estimate uses item quantity instead of a fixed time per order:

- Every item unit adds 0.5 minutes.
- A half portion still counts as one item unit.
- Fractional quantities are rounded up, so 1.5 portions count as two units.
- Two portions count as two units.
- Legacy amount-style entries such as `百頁豆腐20` count as one item unit.
- The pending queue and current order are added together, then 1.5 minutes is
  added for heating and packing.
- The calculated wait is rounded up to the next whole minute.
- The final wait is clamped to a minimum of 10 minutes.

Orders store `pickup_time_source` in column O:

- `requested` times stay unchanged when products or notes are edited, unless
  the customer explicitly provides a new pickup time.
- `estimated` times move later only when an edit increases item units. Only the
  added item time is appended; the 1.5-minute handling time is not added again.
- Note-only changes and reductions never move the pickup time earlier.
- Legacy blank sources are inferred from the saved customer message history.

After deploying this schema change, run `initializePickupTimeSourceColumn()`
once from the Apps Script editor. It creates the Orders column O header and
backfills existing rows with `requested` or `estimated`. New and subsequently
edited orders maintain the value automatically.

Rapid consecutive messages from the same LINE user are serialized with a
short-lived per-user order lock. A later message waits for the preceding order
update before reading context, while different LINE users can still be
processed concurrently. Stale locks expire automatically after 120 seconds.

## AI order intake switch

The operations dashboard includes an `AI 接單` switch. Its state is stored in
Apps Script Properties, so it remains unchanged when the dashboard is closed.
The initial state defaults to enabled.

When the switch is disabled:

- LINE messages do not call OpenAI.
- Orders are not created or updated.
- LINE receives no automatic reply.
- Existing dashboard, sold-out management, archive, and completion jobs keep
  working.

The webhook checks the switch before parsing, after waiting for another
message from the same user, and again inside the spreadsheet write lock. The
last check prevents a request already using AI from saving an order after the
switch has been turned off. The dashboard displays the last switch time.

## Enable

After copying `code.gs` into Apps Script:

1. Select `installAutoCompletionTrigger` in the Apps Script function menu.
2. Click Run and approve the requested Google permissions.
3. Run `getAutoCompletionTriggerStatus` and confirm that `enabled` is `true`.
4. Run `autoCompleteOrdersByPickupTime` manually once with test data before
   using it on live orders.

Running `installAutoCompletionTrigger` again is safe. It removes duplicate
triggers before creating one five-minute trigger.

## Disable

Run `removeAutoCompletionTrigger`. Existing order statuses are not changed.

## Historical order archive

The archive keeps `Orders` small while preserving complete historical detail.
It creates and maintains:

- `OrdersArchive`
- `MonthlySummary`
- `MonthlyItemSummary`

Safety rules:

- Only past-month orders are considered.
- Every row of a `system_order_id` must be `已完成`.
- The complete order is copied before any source row is deleted.
- Existing archived UUIDs are not appended again.
- A rerun repairs summaries and finishes deleting already-copied source rows.
- `UNPRICED` items remain in detail but contribute zero revenue.

### First-time setup

1. Run `initializeOrderArchiveSheets` to create the three sheets and headers.
2. Run `previewMonthlyArchive("2026/06")` with the month you want to verify.
3. Compare `orderCount`, `rowCount`, `revenue`, and
   `unpricedOrderCount` with the source data.
4. Run `archiveMonthlyOrders("2026/06")` only after the preview is correct.
5. Open the dashboard and select the archived month from `檢視月份`.
6. Run `installOrderArchiveTrigger` after the manual test succeeds.

The archive trigger runs daily around 03:00. Daily checking ensures an old
pending order is archived later after its status eventually becomes `已完成`.
Run `getOrderArchiveTriggerStatus` to inspect the trigger, or
`removeOrderArchiveTrigger` to disable it.

The dashboard reads detailed rows from both `Orders` and `OrdersArchive`.
When the same UUID temporarily exists in both places, the archive copy wins so
revenue and order counts are not duplicated.

### June dashboard test data

Run `generateJuneArchiveTestOrders` from the Apps Script editor to add 25
completed test orders directly to `OrdersArchive` for `2026/06`. The function
uses priced items from `PriceList`, spreads orders across the month, and
refreshes `MonthlySummary` and `MonthlyItemSummary`. Reload the dashboard and
select `2026/06` from the month menu.

Generated UUIDs start with `DASHBOARD-TEST-202606-`. Running the generator
again replaces only this generated set instead of duplicating it. Run
`clearJuneArchiveTestOrders` after testing to remove these rows and rebuild the
June summaries without changing real archived orders.
