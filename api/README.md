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
and `ProjectFolderUrl` from the `Project` lookup target. It also returns
`money` (order amount), `sumOfInvoice` (amount already invoiced), and
`remainingAmount` (`money - sumOfInvoice`) for each order.
A blank `sumOfInvoice` is treated as zero.
The response also includes distinct `invoiceNumbers` from the
`רשימת משימות - כספים` list where `supplier` matches the verified supplier.
When `DocumentType` is `חשבונית זיכוי`, the form requires one of those numbers;
the API verifies the selected number belongs to that supplier before upload.

The app registration needs read access to the Projects site to validate
suppliers and read the orders and projects lists.

## Invoice submission API

`POST /api/submitInvoice` checks the submitted invoice against the
`רשימת משימות - כספים` list on the Projects SharePoint site, uploads the PDF to
the selected project's folder, and writes the metadata fields below to the
uploaded document's library item. Any additional documents are uploaded to the
`SupplierAdditionalDocuments` document library and linked to that invoice.

The uploaded filename keeps the original extension and adds a UTC timestamp
(`_yyyyMMdd_HHmmssfff`) before it, for example `invoice_20261019_122135642.pdf`.

| Internal name | Display name | SharePoint type | Value |
| --- | --- | --- | --- |
| `InvoiceNumber` | מספר חשבונית | Single line of text | Invoice number |
| `InvoiceDate` | תאריך חשבונית | Date and time | Date selected in the form |
| `invoiceSum` | סכום החשבונית | Single line of text | Total including VAT entered by the supplier |
| `supplier` | ID ספק | Lookup to `ספקים` | Verified supplier item ID |
| `order0` | ID הזמנה | Lookup to `הזמנות` | Selected order item ID |
| `AllocationNumber` | מספר הקצאה | Single line of text | Allocation number entered in the form |
| `DocumentType` | סוג מסמך | Choice | Selected document type |
| `LinkedInvoiceNumber` | — | Single line of text | Selected original invoice number for credit invoices |
| `_Comments` | הערות לספק | Multiple lines of text | Supplier notes entered in the form |

Each additional document receives the same metadata as the invoice plus
`InvoiceLookupLookupId` (the Graph lookup-ID property for `InvoiceLookup`),
and `ProjectLookupId` (the Graph lookup-ID property for `Project`). These point
to the uploaded invoice's list-item ID and the selected order's verified
project ID, respectively. Additional files are optional and limited to 50 MB
each.

The supplier enters the total amount including VAT; this value is written
directly to `invoiceSum`.
The invoice amount may not exceed the selected order's remaining balance. This
is checked in the client before submission and verified again by the API using
the current `money` and `sumOfInvoice` values from the order list.

The duplicate check reads the `invoiceNumber` and `supplier` columns from the
finance task list.
The app needs Microsoft Graph application permissions with write access to the
Projects site.

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
