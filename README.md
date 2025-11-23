# po-inventory-app

Example Google Apps Script web app for opening and tracking purchase orders (POs) with Google Sheets as the backend.

## Features
- Web UI (Apps Script HTML Service) with tabs for creating a PO, dashboard insights, and receiving goods.
- Saves header-level PO data and line items into two sheets (`PO_MASTER`, `PO_ITEMS`).
- Lead-time summary per supplier based on historical deposit-to-delivery dates.
- Receiving form to update delivery dates, receiver, quantities, and status.

## Files
- `src/Code.gs` — server-side Apps Script functions (sheet setup, PO creation, dashboard aggregation, receiving updates).
- `src/Index.html` — client UI with forms and dashboard widgets.

## How to deploy
1. Create a new Apps Script project connected to your Google Sheet (or from `Extensions` > `Apps Script`).
2. Replace the default files with the contents of the `src` folder (one `.gs` file and one `.html` file). Keep the same file names.
3. Open the sheet and run `setupSheets()` once to create the `PO_MASTER` and `PO_ITEMS` tabs with headers.
4. Click `Deploy` > `New deployment` > `Web app` and set **Execute as:** Me, **Who has access:** Anyone in org (or appropriate scope). Use the provided URL to open the PO portal.

## Sheet structure
### PO_MASTER
The columns follow the provided template: `po_id, supplier_id, deposit_pct, conditions_pay, po_owner, allstatus, ... Real_Recived`. Formulas or extra helper columns can be appended to the right as needed.

### PO_ITEMS
Line items are stored with columns: `po_id, line_no, description, quantity, unit_price, currency, requested_delivery_date, confirmed_delivery_date, received_quantity, receiving_note`.

## Notes
- If `po_id` is blank, the system auto-generates one (e.g., `PO-<timestamp>`).
- Deposit amount is auto-calculated when both `Total_Amount` and `deposit_pct` are provided.
- Lead-time averages per supplier are computed from `deposit_paid_date` to `delivery_date` (or `Real_Recived`).
