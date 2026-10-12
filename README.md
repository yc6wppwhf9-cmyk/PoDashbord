# PO Tracking Desk (prototype)

A single-page purchase-order dashboard built around the ERP **Purchase_Order_Report_New** Excel export. Open `index.html` in a browser; there is no build step.

## Tabs

| Tab | What it does |
| --- | --- |
| **Overview** | KPIs: total POs, PO value, open POs, pending value, overdue POs, received % (by value), cancelled value, POs whose validity is ending. Also: status by count and value, pending value by delay, top suppliers and departments by pending value, overdue and due-in-7-days lists, **department-wise** table (POs, lines, order qty per UOM, net amount, received, pending, overdue), **orders placed by** table, and a site summary. |
| **PO register** | Every PO with search, status filter (New / Partial / Received-cancelled / Overdue) and sorting. Click a row to see its item lines. |
| **My PO tracker** | Type full PO numbers or just the serial (`1312`, `HSB 497`). Each card shows status, delivery and validity alerts, received / cancelled / pending value, item lines, and change history across uploads. |
| **Upload Excel** | Drop the export as it comes from the ERP. Title rows and the totals row are skipped. Each upload replaces the register (the export is a full snapshot). |

Filters for **site, department, supplier and ordered by** (the "Created By" column) apply to Overview and PO register.

## How the numbers are worked out

- A PO is **open** when its status is New or Partial Received and some quantity is still pending.
- **Pending / received / cancelled value** = line Net Amount × (that qty ÷ Order Qty).
- **Overdue** = open PO whose earliest pending line's Delivery Date is before the report's Print Date.
- Order qty is summed per UOM, since PCS, MTR, KG etc. can't be added together.

## Prototype notes

- Data stays in the browser (localStorage). Nothing is uploaded to a server.
- Without an upload, the dashboard shows fictional sample data.

## Next steps for production

1. Database + API so everyone sees the same register and each user keeps their own tracker list.
2. Scheduled import of the ERP export (or a direct ERP connection) instead of manual uploads.
3. Alerts when a tracked PO goes overdue or its validity is about to end.
