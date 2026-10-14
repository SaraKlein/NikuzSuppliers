const { app } = require('@azure/functions');

const graphBaseUrl = 'https://graph.microsoft.com/v1.0';
const siteHost = 'yardend.sharepoint.com';
const sitePath = '/sites/Projects';
const supplierListName = 'ספקים';
const ordersListName = 'הזמנות';
const supplierFields = ['field_1', 'Title', 'field_7'];

function jsonResponse(status, body) {
  return {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
    jsonBody: body
  };
}

function normalize(value) {
  if (typeof value !== 'string' && typeof value !== 'number') {
    return '';
  }

  return String(value).trim().toLowerCase();
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

  const tokenResponse = await response.json();
  if (typeof tokenResponse.access_token !== 'string') {
    throw new Error('Microsoft identity token response did not include an access token');
  }

  return tokenResponse.access_token;
}

async function graphGet(url, accessToken, operation) {
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
    signal: AbortSignal.timeout(15000)
  });

  if (!response.ok) {
    throw new Error(`Microsoft Graph ${operation} request failed (${response.status})`);
  }

  return response.json();
}

function getNextUrl(page) {
  if (!page['@odata.nextLink']) {
    return null;
  }

  const nextUrl = new URL(page['@odata.nextLink']);
  if (nextUrl.origin !== new URL(graphBaseUrl).origin) {
    throw new Error('Microsoft Graph returned an unexpected pagination URL');
  }

  return nextUrl;
}

async function getList(siteId, displayName, accessToken) {
  const url = new URL(`${graphBaseUrl}/sites/${siteId}/lists`);
  url.searchParams.set('$filter', `displayName eq '${displayName.replace(/'/g, "''")}'`);
  url.searchParams.set('$select', 'id,displayName');
  const result = await graphGet(url, accessToken, 'list lookup');
  const list = result.value.find((item) => item.displayName === displayName);

  if (!list) {
    throw new Error(`SharePoint list "${displayName}" was not found`);
  }

  return list;
}

async function getColumns(siteId, listId, accessToken) {
  const url = new URL(`${graphBaseUrl}/sites/${siteId}/lists/${listId}/columns`);
  url.searchParams.set('$select', 'name,displayName,lookup');
  const result = await graphGet(url, accessToken, 'list columns lookup');
  return result.value;
}

async function getItems(siteId, listId, selectFields, accessToken, operation) {
  let url = new URL(`${graphBaseUrl}/sites/${siteId}/lists/${listId}/items`);
  url.searchParams.set('$expand', `fields($select=${selectFields.join(',')})`);
  url.searchParams.set('$top', '200');
  const items = [];

  while (url) {
    const page = await graphGet(url, accessToken, operation);
    items.push(...page.value);
    url = getNextUrl(page);
  }

  return items;
}

async function findVerifiedSupplierIds(siteId, values, accessToken) {
  const supplierList = await getList(siteId, supplierListName, accessToken);
  const items = await getItems(siteId, supplierList.id, supplierFields, accessToken, 'supplier verification');
  const ids = new Set();

  for (const item of items) {
    const fields = item.fields || {};
    if (
      normalize(fields.field_1) === values.supplierNumber
      && normalize(fields.Title) === values.supplierName
      && normalize(fields.field_7) === values.supplierEmail
    ) {
      ids.add(String(item.id));
    }
  }

  return { supplierList, ids };
}

function lookupIds(value) {
  if (typeof value !== 'string' && typeof value !== 'number') {
    return [];
  }

  return String(value).split(',').map((id) => id.trim()).filter(Boolean);
}

app.http('getPurchaseOrders', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const credentials = getCredentials();
    if (!credentials) {
      context.error('Purchase order lookup is not configured: SharePoint credentials are missing.');
      return jsonResponse(503, { error: 'טעינת הזמנות הרכש אינה זמינה כרגע.' });
    }

    let payload;
    try {
      const body = await request.text();
      if (body.length > 4096) {
        return jsonResponse(413, { error: 'הבקשה גדולה מדי.' });
      }
      payload = JSON.parse(body);
    } catch {
      return jsonResponse(400, { error: 'הבקשה אינה תקינה.' });
    }

    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
      return jsonResponse(400, { error: 'הבקשה אינה תקינה.' });
    }

    const { supplierNumber, supplierName, supplierEmail } = payload;
    if (
      typeof supplierNumber !== 'string'
      || typeof supplierName !== 'string'
      || typeof supplierEmail !== 'string'
      || supplierNumber.trim().length < 1
      || supplierNumber.trim().length > 64
      || supplierName.trim().length < 1
      || supplierName.trim().length > 200
      || supplierEmail.trim().length > 254
      || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(supplierEmail.trim())
    ) {
      return jsonResponse(400, { error: 'נא למלא פרטי ספק תקינים.' });
    }

    try {
      const accessToken = await getAccessToken(credentials);
      const site = await graphGet(
        `${graphBaseUrl}/sites/${siteHost}:` + sitePath,
        accessToken,
        'site lookup'
      );
      const { supplierList, ids: supplierIds } = await findVerifiedSupplierIds(site.id, {
        supplierNumber: normalize(supplierNumber),
        supplierName: normalize(supplierName),
        supplierEmail: normalize(supplierEmail)
      }, accessToken);

      if (!supplierIds.size) {
        return jsonResponse(403, { error: 'פרטי הספק לא אומתו מול SharePoint.' });
      }

      const ordersList = await getList(site.id, ordersListName, accessToken);
      const columns = await getColumns(site.id, ordersList.id, accessToken);
      const supplierColumn = columns.find((column) => (
        column.lookup
        && String(column.lookup.listId).toLowerCase() === String(supplierList.id).toLowerCase()
      ));

      if (!supplierColumn) {
        throw new Error('The orders list does not have a supplier lookup column pointing to the supplier list');
      }

      const orders = await getItems(
        site.id,
        ordersList.id,
        ['Title', 'Id'],//, `${supplierColumn.name}Id`],
        accessToken,
        'purchase order title lookup'
      );
      const result = orders
      //  .filter((order) => lookupIds((order.fields || {})[`${supplierColumn.name}LookupId`])
        //  .some((id) => supplierIds.has(id)))
        .map((order) => {
          const fields = order.fields || {};
          return {
            id: String(order.id),
            title: typeof fields.Title === 'string' ? fields.Title : ''
          };
        })
        .filter((order) => order.title);

      return jsonResponse(200, { orders: result });
    } catch (error) {
      context.error('Purchase order lookup failed.', error);
      return jsonResponse(502, { error: 'לא ניתן לטעון את הזמנות הרכש כרגע.' });
    }
  }
});
