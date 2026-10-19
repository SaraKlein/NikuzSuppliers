# Supplier portal API

The `POST /api/validateSupplier` endpoint checks that the supplier number, name,
and email match one item in the `ספקים` list on the
[Projects SharePoint site](https://yardend.sharepoint.com/sites/Projects/).
The matching SharePoint internal field names are `field_1`, `Title`, and
`field_7`, respectively. A valid response also includes the matched list item's
SharePoint ID as `supplierId`; the portal retains it in the invoice form for
future use.

The `POST /api/getPurchaseOrders` endpoint repeats that supplier check, then
returns items from the `הזמנות` list whose `SupplierText` field matches the
verified supplier name. The current diagnostic response includes each order's
SharePoint item ID and `Title`, plus the referenced project's ID and `Title`
and `ProjectFolderUrl` from the `Project` lookup target.

The app registration needs read access to the Projects site to validate
suppliers and read the orders and projects lists.

## Invoice submission API

`POST /api/submitInvoice` checks the submitted invoice against the
`רשימת משימות - כספים` list on the Projects SharePoint site, uploads the PDF to
the selected project's `חשבוניות` folder, and writes the following metadata
to the uploaded document's library item:

| Internal name | Display name | SharePoint type | Value |
| --- | --- | --- | --- |
| `InvoiceNumber` | מספר חשבונית | Single line of text | Invoice number |
| `invoiceSum` | סכום החשבונית | Single line of text | Total including VAT, formatted to 2 decimals |
| `supplier` | ID ספק | Lookup to `ספקים` | Verified supplier item ID |
| `order` | ID הזמנה | Lookup to `הזמנות` | Selected order item ID |

The duplicate check reads the `invoiceNumber` and `supplier` columns from the
finance task list. Metadata updates use Microsoft Graph lookup field names
`supplierLookupId` and `orderLookupId`. The app needs Microsoft Graph
application permissions with write access to the Projects site.

Configure these application settings for the Azure Static Web App API:

- `SHAREPOINT_TENANT_ID`: Microsoft Entra tenant ID
- `SHAREPOINT_CLIENT_ID`: app registration's application (client) ID
- `SHAREPOINT_CLIENT_SECRET`: app registration secret

The app registration needs Microsoft Graph **application** permission
`Sites.Selected`, admin consent, and an explicit **read** grant to the Projects
site for current supplier, order, and finance-task validation. Grant **write**
access to the Projects site when invoice saving is enabled. Invoice files are
uploaded through Microsoft Graph, so no SharePoint API permission is required.
Store the secret only in Azure application settings; do not add it to the
repository or send it to the browser.

The supplier validation endpoint only checks whether the supplied details
match a list item; it does not authenticate the person or establish a logged-in
session. Invoice submission repeats that check server-side, but anyone who
knows a supplier's matching details could submit an invoice. Add an identity
authentication step if submissions must be limited to the actual supplier.
