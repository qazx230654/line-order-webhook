# Google Sheets Schema

This document records the Google Sheets structure used by the LINE order
assistant.

- Spreadsheet ID: `1YHRe3-ApLT6GMrh7R3Qn6xgMaKRCE5_oG8cFJQnXdkE`
- Spreadsheet URL: <https://docs.google.com/spreadsheets/d/1YHRe3-ApLT6GMrh7R3Qn6xgMaKRCE5_oG8cFJQnXdkE/edit>
- Last verified: `2026-07-13`
- Apps Script arrays use zero-based indexes. Sheet rows and columns are
  one-based.

## Orders

One order can occupy multiple rows. Rows belonging to the same order share the
same `system_order_id` and `display_order_id`.

| Column | Array index | Header | Purpose |
| --- | ---: | --- | --- |
| A | 0 | `system_order_id` | Unique internal order ID (UUID) |
| B | 1 | `display_order_id` | Short customer-facing order number, such as `A001` |
| C | 2 | `created_at` | Order creation time |
| D | 3 | `customer_name` | Customer name |
| E | 4 | `phone` | Customer phone number |
| F | 5 | `group_name` | Product group/category recorded on the order |
| G | 6 | `item_name` | Ordered item name |
| H | 7 | `quantity` | Ordered quantity |
| I | 8 | `pickup_time` | Requested pickup time |
| J | 9 | `note` | Order or item note |
| K | 10 | `status` | Order status, for example `待製作` |
| L | 11 | `raw_message` | Original LINE message |
| M | 12 | `line_user_id` | LINE user ID |
| N | 13 | `unit_price` | Transaction-time unit price; missing prices use `UNPRICED` |

## PriceList

| Column | Array index | Header | Purpose |
| --- | ---: | --- | --- |
| A | 0 | `item_name` | Canonical item name |
| B | 1 | `price` | Current item price |
| C | 2 | `category` | Menu category |
| D | 3 | `image_url` | Item image URL |
| E | 4 | `is_sold_out` | Sold-out flag used by menu management |

## ItemAlias

| Column | Array index | Header | Purpose |
| --- | ---: | --- | --- |
| A | 0 | `別名` | Alternate name accepted from customer input |
| B | 1 | `正式名稱` | Canonical name matching `PriceList.item_name` |

## PendingOrders

| Column | Array index | Header | Purpose |
| --- | ---: | --- | --- |
| A | 0 | `pending_id` | Unique pending-order ID |
| B | 1 | `line_user_id` | LINE user ID |
| C | 2 | `customer_name` | Customer name |
| D | 3 | `created_at` | Pending order creation time |
| E | 4 | `order_json` | Serialized parsed-order data |
| F | 5 | `raw_message` | Original LINE message |
| G | 6 | `last_message_time` | Time of the latest related message |

## Customers

| Column | Array index | Header | Purpose |
| --- | ---: | --- | --- |
| A | 0 | `line_user_id` | Unique LINE user ID |
| B | 1 | `display_name` | LINE display name |
| C | 2 | `picture_url` | LINE profile image URL |
| D | 3 | `phone` | Most recently known phone number |
| E | 4 | `last_order_time` | Most recent order time |
| F | 5 | `total_orders` | Customer order count |

## Schema Change Checklist

When adding, removing, or reordering a column:

1. Update the header row in Google Sheets.
2. Update this document, including the column letter and array index.
3. Search `code.gs` for the sheet name, header name, `getValues()`, and direct
   indexes such as `row[13]`.
4. Update every `setValues()` row to match the new column count and order.
5. Update dashboard or menu field mappings when the changed column is displayed.
6. Test both existing rows and newly created rows before deployment.

