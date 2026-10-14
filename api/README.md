# Supplier portal API

The `POST /api/validateSupplier` endpoint checks that the supplier number, name,
and email match one item in the `ספקים` list on the
[Projects SharePoint site](https://yardend.sharepoint.com/sites/Projects/).
The matching SharePoint internal field names are `field_1`, `Title`, and
`field_7`, respectively.

The `POST /api/getPurchaseOrders` endpoint repeats that supplier check, then
returns items from the `הזמנות` list whose supplier lookup points to the
verified supplier. The current diagnostic response includes each order's
SharePoint item ID and `Title`; project and folder fields are temporarily
excluded while the order-title lookup is being tested.

The app registration needs read access to the Projects site to validate
suppliers and read the orders and projects lists.

## Invoice submission API

The `POST /api/submitInvoice` endpoint creates an item in the `Invoices` list
on the SaraTest SharePoint site and adds the required invoice PDF and any
supporting files as list attachments. It accepts `multipart/form-data` with
the form's supplier, invoice, and file fields. Uploads are limited to 50 MB in
total.

Create these columns in `Invoices` with the exact display names and types.
Column internal names can differ; the API resolves them by display name.

| Display name | SharePoint type |
| --- | --- |
| Supplier Number | Single line of text |
| Supplier Name | Single line of text |
| Supplier Email | Single line of text |
| Invoice Number | Single line of text |
| Invoice Date | Date and time (date only) |
| Net Amount | Currency |
| VAT Amount | Currency |
| Total Amount | Currency |
| Allocation Number | Single line of text |
| Document Type | Choice: `חשבונית מס`, `חשבונית זיכוי`, `חשבונית מס / קבלה` |
| Supplier Notes | Multiple lines of text |
| Purchase Order | Single line of text (needed when the form sends a purchase order) |

The built-in `Title` field is populated with the invoice number. Enable list
attachments so the PDF and supporting files can be stored on the item.

Configure these application settings for the Azure Static Web App API:

- `SHAREPOINT_TENANT_ID`: Microsoft Entra tenant ID
- `SHAREPOINT_CLIENT_ID`: app registration's application (client) ID
- `SHAREPOINT_CLIENT_SECRET`: app registration secret

The app registration needs Microsoft Graph **application** permission
`Sites.Selected`, admin consent, and an explicit **read** grant to the Projects
site plus an explicit **write** grant to the SaraTest site. The API uses these
grants to verify suppliers, read orders and projects, and create invoices. For
file attachments, also grant the SharePoint API application permission
`Sites.Selected` and **write** access to the SaraTest site. Store the secret
only in Azure application settings; do not add it to the repository or send it
to the browser.

The supplier validation endpoint only checks whether the supplied details
match a list item; it does not authenticate the person or establish a logged-in
session. Invoice submission repeats that check server-side, but anyone who
knows a supplier's matching details could submit an invoice. Add an identity
authentication step if submissions must be limited to the actual supplier.
