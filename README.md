# PO Tracking Desk (prototype)

A single-page purchase-order dashboard. Open `index.html` in a browser; no build step.

## Tabs

| Tab | What it does |
| --- | --- |
| **Overview** | KPIs (total POs, open PO value, on-time delivery %, overdue POs, fill rate, avg lead time), status mix, open value by vendor, overdue and due-this-week lists, and a searchable table of all POs. |
| **Email reports** | Inbox of PO report emails. The PO table in each email body is parsed and imported (new POs are added, existing ones updated). "Check inbox" simulates a new email arriving. You can also paste an email body. |
| **My PO tracker** | Type one or more PO numbers to watch. Each card shows a stage tracker (Open → Approved → Shipped → Delivered), due/overdue state, fill %, and status history. POs that aren't in any report yet wait until a report mentions them. |
| **Upload Excel** | Drag in an `.xlsx`, `.xls` or `.csv`. Column names are matched loosely (e.g. "PO No.", "Supplier", "GRN Qty", "Delivery Date"). Preview, then import. Only the PO number column is required. |

Try `samples/sample_po_upload.xlsx` on the Upload tab.

## Prototype notes

- Data is sample data and is saved in the browser's localStorage. Run `PODesk.reset()` in the browser console to restore the sample data.
- Currency/locale are set by `CURRENCY` and `LOCALE` at the top of the script.
- Excel parsing uses SheetJS loaded from cdnjs.

## Next steps for production

1. **Email fetch**: a scheduled backend job (Gmail API / Microsoft Graph / IMAP) reads the PO report mailbox, parses the body table or Excel attachment with the same column mapping, and writes to a database.
2. **Database + API** (e.g. Postgres/Supabase) replacing localStorage, so all users share the same PO data and watchlists are per user.
3. **Auth** and per-user tracker lists, plus alerts (email/Slack) when a tracked PO goes overdue or changes status.
