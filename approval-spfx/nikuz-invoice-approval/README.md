# Nikuz SharePoint web parts

This SPFx solution contains two web parts for the Projects SharePoint site:

- **Invoice Approval** displays and approves one invoice after checking that
  the signed-in user created its linked purchase order.
- **My Invoices** lists supplier-app invoices linked to purchase orders
  initiated by the signed-in user.

## SharePoint data

Both web parts expect the `מסמכי פרויקט` document library and the `הזמנות`
list to be on the same site. Invoice documents use the `order0` lookup to link
to an order; the My Invoices web part only includes documents where
`CreatedBySupplierApp` is `true`.

The My Invoices web part queries orders created by the current SharePoint user,
then retrieves supplier-app invoices linked to each of those orders.

## Invoice approval page

Open the page with an `itemId` query parameter containing the invoice list item
ID, for example:

`https://yardend.sharepoint.com/sites/Projects/SitePages/InvoiceApproval.aspx?itemId=123`

The SPFx project checks the `Author` identity before showing invoice details.
Approval updates `InvoiceStatus` to `approve` or `reject`; a rejection requires
a comment, saved in `OrderInitiatorComments`.

## My Invoices page

Add the **החשבוניות שלי** web part to a modern SharePoint page on the same site.
It automatically uses the currently signed-in SharePoint user and needs no URL
parameters.

## Security requirements before deployment

The identity filtering in a web part is not a security boundary by itself.
SharePoint permissions must also enforce access to invoices and files:

1. Assign invoice read/edit permissions to the intended order initiator and
   administrators, on both the library item and its file.
2. Verify with another SharePoint user that direct REST requests cannot read
   or update invoices they are not authorized to access.
3. Decommission the old anonymous `GET`/`POST /api/invoiceApproval` Azure
   Function before using the SPFx approval flow. That endpoint uses
   application credentials and otherwise bypasses the signed-in user's
   SharePoint permissions. Do not use the old `approval.html` page.

The current supplier-upload process does not set these per-item permissions.
Configure an administrative process (for example, a Power Automate flow or a
server-side workflow) to assign access when an invoice is uploaded. Do not
rely on hiding web-part content to protect invoice data or decisions.

## Build and package

Use Node.js 24 (24.0 or later within 24.x) or another version supported by
SPFx 1.24.1.

```powershell
npm install
npm run build
```

The SharePoint package is created at
`sharepoint/solution/nikuz-invoice-approval.sppkg`. Upload it to the tenant or
site App Catalog, deploy it, add the app to the Projects site, then add either
web part to its modern SharePoint page.

## Local development

```powershell
npm install
npm start
```

Use the SharePoint hosted workbench or a modern SharePoint page to test against
the Projects site. Testing requires a signed-in account that has the necessary
list and item permissions.
