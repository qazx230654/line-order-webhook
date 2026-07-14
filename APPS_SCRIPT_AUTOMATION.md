# Apps Script Automation

## Pickup-time auto completion

`autoCompleteOrdersByPickupTime` checks pending orders and completes an entire
order when its pickup time is reached.

Behavior:

- Runs every five minutes after the trigger is installed.
- Only processes rows whose status is `待製作`.
- Groups rows by `system_order_id` and completes the entire order together.
- Rechecks status under a script lock to prevent duplicate completion notices.
- Sends the existing LINE completion notification when a LINE user ID exists.
- Skips empty or unsafe pickup-time values.
- Skips orders overdue by more than 12 hours to avoid notifying old customers
  when automation is enabled for the first time.

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
