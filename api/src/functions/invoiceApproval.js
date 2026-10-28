const { app } = require('@azure/functions');
const { writeApplicationLog } = require('../shared/applicationLogger');

const graphBaseUrl = 'https://graph.microsoft.com/v1.0';
const siteHost = 'yardend.sharepoint.com';
const sitePath = '/sites/Projects';
const invoiceLibraryName = 'מסמכי פרויקט';

function jsonResponse(status, body) {
  return {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
    jsonBody: body
  };
}

function getCredentials() {
  const { SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET } = process.env;
  if (!SHAREPOINT_TENANT_ID || !SHAREPOINT_CLIENT_ID || !SHAREPOINT_CLIENT_SECRET) {
    return null;
  }
  return {
    tenantId: SHAREPOINT_TENANT_ID,
    clientId: SHAREPOINT_CLIENT_ID,
    clientSecret: SHAREPOINT_CLIENT_SECRET
  };
}

async function getAccessToken(credentials) {
  const body = new URLSearchParams({
    client_id: credentials.clientId,
    client_secret: credentials.clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });
  const response = await fetch(
    `https://login.microsoftonline.com/${encodeURIComponent(credentials.tenantId)}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body,
      signal: AbortSignal.timeout(10000)
    }
  );
  if (!response.ok) {
    throw new Error(`Microsoft identity token request failed (${response.status})`);
  }
  const result = await response.json();
  if (typeof result.access_token !== 'string') {
    throw new Error('Microsoft identity token response did not include an access token');
  }
  return result.access_token;
}

async function graphRequest(url, accessToken, operation, options) {
  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options && options.headers)
    },
    signal: AbortSignal.timeout(15000)
  });
  if (!response.ok) {
    let graphMessage = '';
    try {
      const body = await response.json();
      graphMessage = body && body.error && body.error.message
        ? ` ${body.error.message}`
        : '';
    } catch {
      // Graph did not return JSON; retain the HTTP status.
    }
    throw new Error(`Microsoft Graph ${operation} request failed (${response.status}).${graphMessage}`);
  }
  if (response.status === 204) {
    return null;
  }
  return response.json();
}

async function getSiteAndInvoiceList(accessToken) {
  const site = await graphRequest(
    `${graphBaseUrl}/sites/${siteHost}:${sitePath}?$select=id,webUrl`,
    accessToken,
    'invoice site lookup'
  );
  const listsUrl = new URL(`${graphBaseUrl}/sites/${encodeURIComponent(site.id)}/lists`);
  listsUrl.searchParams.set('$filter', `displayName eq '${invoiceLibraryName}'`);
  listsUrl.searchParams.set('$select', 'id,displayName,webUrl');
  const lists = await graphRequest(listsUrl, accessToken, 'invoice library lookup');
  const list = (lists.value || []).find((item) => item.displayName === invoiceLibraryName);
  if (!list) {
    throw new Error(`SharePoint document library "${invoiceLibraryName}" was not found.`);
  }
  return { site, list };
}

async function getInvoice(accessToken, itemId) {
  const { site, list } = await getSiteAndInvoiceList(accessToken);
  const itemUrl = new URL(
    `${graphBaseUrl}/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(list.id)}/items/${encodeURIComponent(itemId)}`
  );
  itemUrl.searchParams.set('$expand', 'fields');
  itemUrl.searchParams.set('$select', 'id,webUrl,fields');
  const item = await graphRequest(itemUrl, accessToken, 'invoice item lookup');
  const fields = item.fields || {};
  return {
    itemId,
    listId: list.id,
    siteId: site.id,
    fileUrl: item.webUrl || fields.FileRef || null,
    fields: {
      name: fields.FileLeafRef || fields.Title || '',
      invoiceNumber: fields.InvoiceNumber || '',
      invoiceDate: fields.InvoiceDate || '',
      amount: fields.invoiceSum ?? '',
      allocationNumber: fields.AllocationNumber || '',
      documentType: fields.DocumentType || '',
      linkedInvoiceNumber: fields.LinkedInvoiceNumber || '',
      supplierName: fields.SupplierName || fields.supplierLookupId || '',
      purchaseOrder: fields.PurchaseOrder || fields.Order || fields.order0LookupId || '',
      supplierNotes: fields._Comments || '',
      status: fields.InvoiceStatus || '',
      comments: fields.OrderInitiatorComments || '',
      modified: item.fields && item.fields.Modified ? item.fields.Modified : ''
    }
  };
}

function getItemId(request) {
  const itemId = request.query.get('itemId');
  return /^\d+$/.test(itemId || '') ? itemId : null;
}

app.http('invoiceApproval', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  route: 'invoiceApproval',
  handler: async (request, context) => {
    const itemId = getItemId(request);
    if (!itemId) {
      return jsonResponse(400, { error: 'מזהה החשבונית אינו תקין.' });
    }

    const credentials = getCredentials();
    if (!credentials) {
      context.error('Invoice approval is not configured: SharePoint credentials are missing.');
      return jsonResponse(503, { error: 'שירות אישור החשבוניות אינו זמין כרגע.' });
    }

    try {
      const accessToken = await getAccessToken(credentials);

      if (request.method === 'GET') {
        return jsonResponse(200, await getInvoice(accessToken, itemId));
      }

      let payload;
      try {
        payload = await request.json();
      } catch {
        return jsonResponse(400, { error: 'הבקשה אינה תקינה.' });
      }
      const decision = typeof payload.status === 'string'
        ? payload.status.trim().toLowerCase()
        : '';
      const comments = typeof payload.comments === 'string'
        ? payload.comments.trim()
        : '';

      if (!['approve', 'reject'].includes(decision)) {
        return jsonResponse(400, { error: 'יש לבחור אישור או דחייה.' });
      }
      if (decision === 'reject' && !comments) {
        return jsonResponse(400, { error: 'נדרש להוסיף הערה בעת דחיית החשבונית.' });
      }
      if (comments.length > 10000) {
        return jsonResponse(400, { error: 'ההערה ארוכה מדי.' });
      }

      const invoice = await getInvoice(accessToken, itemId);
      await graphRequest(
        `${graphBaseUrl}/sites/${encodeURIComponent(invoice.siteId)}/lists/${encodeURIComponent(invoice.listId)}/items/${encodeURIComponent(itemId)}/fields`,
        accessToken,
        'invoice approval update',
        {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            InvoiceStatus: decision,
            OrderInitiatorComments: comments
          })
        }
      );

      try {
        await writeApplicationLog({
          siteId: invoice.siteId,
          accessToken,
          severity: 'Information',
          message: `Invoice ${decision}d.`,
          details: {
            itemId,
            invoiceNumber: invoice.fields.invoiceNumber,
            decision,
            commentLength: comments.length
          },
          correlationId: context.invocationId,
          httpStatus: 200
        });
      } catch (auditError) {
        context.warn('Invoice decision was saved, but audit logging failed.', auditError);
      }

      return jsonResponse(200, {
        success: true,
        itemId,
        status: decision,
        comments
      });
    } catch (error) {
      context.error('Invoice approval request failed.', error);
      if (/request failed \(404\)/.test(String(error.message))) {
        return jsonResponse(404, { error: 'החשבונית המבוקשת לא נמצאה.' });
      }
      return jsonResponse(500, { error: 'לא ניתן לעדכן את החלטת החשבונית כרגע.' });
    }
  }
});
