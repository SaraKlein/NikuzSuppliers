const { app } = require('@azure/functions');

const graphBaseUrl = 'https://graph.microsoft.com/v1.0';
const siteHost = 'yardend.sharepoint.com';
const sitePath = '/sites/Projects';
const listName = 'Invoices';
const additionalDocumentsLibraryName = 'SupplierAdditionalDocuments';
const financeTaskListName = 'רשימת משימות - כספים';
const supplierListName = 'ספקים';
const maxUploadBytes = 50 * 1024 * 1024;
const invoiceColumns = [
  'SupplierNumber',
  'SupplierName',
  'SupplierEmail',
  'InvoiceNumber',
  'InvoiceDate',
  'NetAmount',
  'VATAmount',
  'TotalAmount',
  'AllocationNumber',
  'DocumentType',
  'SupplierNotes'
];

class GraphRequestError extends Error {
  constructor(operation, status) {
    super(`Microsoft Graph ${operation} request failed (${status})`);
    this.operation = operation;
    this.status = status;
  }
}

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

async function getAccessToken(credentials, scope) {
  const body = new URLSearchParams({
    client_id: credentials.clientId,
    client_secret: credentials.clientSecret,
    scope,
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

async function graphRequest(url, accessToken, options, operation, timeoutMs = 20000) {
  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options && options.headers)
    },
    signal: AbortSignal.timeout(timeoutMs)
  });

  if (!response.ok) {
    throw new GraphRequestError(operation, response.status);
  }

  if (response.status === 204) {
    return null;
  }

  return response.json();
}

async function graphRequestIfExists(url, accessToken, operation) {
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
    signal: AbortSignal.timeout(20000)
  });

  if (response.status === 404) {
    return null;
  }
  if (!response.ok) {
    throw new GraphRequestError(operation, response.status);
  }

  return response.json();
}

async function getSharePointSite(accessToken) {
  const site = await graphRequest(
    `${graphBaseUrl}/sites/${siteHost}:${sitePath}`,
    accessToken,
    null,
    'site lookup'
  );
  return site;
}

async function financeTaskExists(siteId, invoiceValues, accessToken) {
  const listsUrl = new URL(`${graphBaseUrl}/sites/${siteId}/lists`);
  listsUrl.searchParams.set('$filter', `displayName eq '${financeTaskListName}'`);
  listsUrl.searchParams.set('$select', 'id,displayName');
  const lists = await graphRequest(listsUrl, accessToken, null, 'finance task list lookup');
  const list = lists.value.find((item) => item.displayName === financeTaskListName);

  if (!list) {
    throw new Error(`SharePoint list "${financeTaskListName}" was not found`);
  }

  let itemsUrl = new URL(`${graphBaseUrl}/sites/${siteId}/lists/${list.id}/items`);
  itemsUrl.searchParams.set('$expand', 'fields');
  itemsUrl.searchParams.set('$top', '200');

  while (itemsUrl) {
    const page = await graphRequest(itemsUrl, accessToken, null, 'finance task verification');
    const found = page.value.some((item) => {
      const fields = item.fields || {};
      return String(fields.invoiceNumber || '').trim() === invoiceValues.invoiceNumber
        && String(fields.supplier || '').trim().toLowerCase()
        === invoiceValues.supplierName.toLowerCase();
    });
    if (found) {
      return true;
    }

    if (page['@odata.nextLink']) {
      itemsUrl = new URL(page['@odata.nextLink']);
      if (itemsUrl.origin !== new URL(graphBaseUrl).origin) {
        throw new Error('Microsoft Graph returned an unexpected pagination URL');
      }
    } else {
      itemsUrl = null;
    }
  }

  return false;
}

async function supplierInvoiceExists(siteId, supplierName, invoiceNumber, accessToken) {
  const listsUrl = new URL(`${graphBaseUrl}/sites/${siteId}/lists`);
  listsUrl.searchParams.set('$filter', `displayName eq '${financeTaskListName}'`);
  listsUrl.searchParams.set('$select', 'id,displayName');
  const lists = await graphRequest(listsUrl, accessToken, null, 'finance task list lookup');
  const list = lists.value.find((item) => item.displayName === financeTaskListName);
  if (!list) {
    throw new Error(`SharePoint list "${financeTaskListName}" was not found`);
  }

  let itemsUrl = new URL(`${graphBaseUrl}/sites/${siteId}/lists/${list.id}/items`);
  itemsUrl.searchParams.set('$expand', 'fields($select=invoiceNumber,supplier)');
  itemsUrl.searchParams.set('$top', '200');

  while (itemsUrl) {
    const page = await graphRequest(itemsUrl, accessToken, null, 'credit invoice reference verification');
    const found = page.value.some((item) => {
      const fields = item.fields || {};
      return String(fields.invoiceNumber || '').trim() === invoiceNumber
        && String(fields.supplier || '').trim().toLowerCase() === supplierName.toLowerCase();
    });
    if (found) {
      return true;
    }

    if (page['@odata.nextLink']) {
      itemsUrl = new URL(page['@odata.nextLink']);
      if (itemsUrl.origin !== new URL(graphBaseUrl).origin) {
        throw new Error('Microsoft Graph returned an unexpected pagination URL');
      }
    } else {
      itemsUrl = null;
    }
  }

  return false;
}

async function getVerifiedSupplierId(siteId, values, accessToken) {
  const listsUrl = new URL(`${graphBaseUrl}/sites/${siteId}/lists`);
  listsUrl.searchParams.set('$filter', `displayName eq '${supplierListName}'`);
  listsUrl.searchParams.set('$select', 'id,displayName');
  const lists = await graphRequest(listsUrl, accessToken, null, 'supplier list lookup');
  const list = lists.value.find((item) => item.displayName === supplierListName);
  if (!list) {
    throw new Error(`SharePoint list "${supplierListName}" was not found`);
  }

  let itemsUrl = new URL(`${graphBaseUrl}/sites/${siteId}/lists/${list.id}/items`);
  itemsUrl.searchParams.set('$expand', 'fields($select=field_1,Title,field_7)');
  itemsUrl.searchParams.set('$top', '200');
  while (itemsUrl) {
    const page = await graphRequest(itemsUrl, accessToken, null, 'supplier verification');
    const found = page.value.find((item) => {
      const fields = item.fields || {};
      return String(fields.field_1 || '').trim().toLowerCase() === values.supplierNumber.toLowerCase()
        && String(fields.Title || '').trim().toLowerCase() === values.supplierName.toLowerCase()
        && String(fields.field_7 || '').trim().toLowerCase() === values.supplierEmail.toLowerCase();
    });
    if (found) {
      return String(found.id);
    }

    if (page['@odata.nextLink']) {
      itemsUrl = new URL(page['@odata.nextLink']);
      if (itemsUrl.origin !== new URL(graphBaseUrl).origin) {
        throw new Error('Microsoft Graph returned an unexpected pagination URL');
      }
    } else {
      itemsUrl = null;
    }
  }

  return null;
}

function getOrderAmount(value, fieldName, emptyValue) {
  if ((typeof value === 'string' && !value.trim()) || value == null) {
    if (emptyValue !== undefined) {
      return emptyValue;
    }
    throw new Error(`Purchase order field "${fieldName}" is missing.`);
  }

  if (typeof value !== 'string' && typeof value !== 'number') {
    throw new Error(`Purchase order field "${fieldName}" is not a valid amount.`);
  }

  const amount = Number(value);
  if (!Number.isFinite(amount)) {
    throw new Error(`Purchase order field "${fieldName}" is not a valid amount.`);
  }

  return amount;
}

function roundCurrency(amount) {
  return Math.round((amount + Number.EPSILON) * 100) / 100;
}

async function getVerifiedPurchaseOrder(siteId, orderId, supplierName, accessToken) {
  const listsUrl = new URL(`${graphBaseUrl}/sites/${siteId}/lists`);
  listsUrl.searchParams.set('$filter', "displayName eq 'הזמנות'");
  listsUrl.searchParams.set('$select', 'id,displayName');
  const lists = await graphRequest(listsUrl, accessToken, null, 'purchase order list lookup');
  const list = lists.value.find((item) => item.displayName === 'הזמנות');
  if (!list) {
    throw new Error('SharePoint purchase order list was not found.');
  }

  const itemUrl = new URL(
    `${graphBaseUrl}/sites/${siteId}/lists/${encodeURIComponent(list.id)}/items/${encodeURIComponent(orderId)}`
  );
  itemUrl.searchParams.set('$expand', 'fields($select=SupplierText,money,sumOfInvoice)');
  const order = await graphRequest(itemUrl, accessToken, null, 'purchase order verification');
  const fields = order.fields || {};
  const orderSupplier = String(fields.SupplierText || '').trim().toLowerCase();
  if (orderSupplier !== supplierName.trim().toLowerCase()) {
    return { belongsToSupplier: false };
  }

  const money = getOrderAmount(fields.money, 'money');
  const sumOfInvoice = getOrderAmount(fields.sumOfInvoice, 'sumOfInvoice', 0);
  return {
    belongsToSupplier: true,
    remainingAmount: roundCurrency(money - sumOfInvoice)
  };
}

function getTextField(formData, name) {
  const value = formData.get(name);
  return typeof value === 'string' ? value.trim() : '';
}

function validDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return false;
  }
  const date = new Date(`${value}T00:00:00.000Z`);
  return Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value;
}

function getInvoiceValues(formData) {
  const supplierNumber = getTextField(formData, 'verifiedNumber');
  const supplierName = getTextField(formData, 'verifiedName');
  const supplierEmail = getTextField(formData, 'verifiedEmail');
  const invoiceNumber = getTextField(formData, 'invoiceNumber');
  const invoiceDate = getTextField(formData, 'invoiceDate');
  const totalText = getTextField(formData, 'invoiceTotal');
  const allocationNumber = getTextField(formData, 'allocationNumber');
  const documentType = getTextField(formData, 'documentType');
  const creditInvoiceNumber = getTextField(formData, 'creditInvoiceNumber');
  const supplierNotes = getTextField(formData, 'supplierNotes');
  const purchaseOrder = getTextField(formData, 'purchaseOrder');
  const totalAmount = Number(totalText);

  if (
    !supplierNumber
    || supplierNumber.length > 64
    || !supplierName
    || supplierName.length > 200
    || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(supplierEmail)
    || !invoiceNumber
    || invoiceNumber.length > 64
    || !validDate(invoiceDate)
    || !totalText
    || !Number.isFinite(totalAmount)
    || totalAmount < 0.01
    || !['tax-invoice', 'credit-invoice', 'invoice-receipt', 'business-invoice'].includes(documentType)
    || (documentType === 'credit-invoice' && !creditInvoiceNumber)
    || creditInvoiceNumber.length > 64
    || allocationNumber.length > 64
    || supplierNotes.length > 5000
    || purchaseOrder.length > 128
  ) {
    return null;
  }

  return {
    supplierNumber,
    supplierName,
    supplierEmail,
    invoiceNumber,
    invoiceDate,
    totalAmount,
    allocationNumber,
    creditInvoiceNumber,
    documentType: {
      'tax-invoice': 'חשבונית מס',
      'credit-invoice': 'חשבונית זיכוי',
      'invoice-receipt': 'חשבונית מס / קבלה',
      'business-invoice': 'חשבונית עסקה'
    }[documentType],
    supplierNotes,
    purchaseOrder
  };
}

function isFile(value) {
  return value && typeof value === 'object'
    && typeof value.name === 'string'
    && typeof value.arrayBuffer === 'function';
}

function safeFileName(name, index, usedNames) {
  const cleaned = name
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_')
    .replace(/[. ]+$/g, '')
    .slice(-180) || `attachment-${index + 1}`;
  const dot = cleaned.lastIndexOf('.');
  const stem = dot > 0 ? cleaned.slice(0, dot) : cleaned;
  const extension = dot > 0 ? cleaned.slice(dot) : '';
  let candidate = cleaned;
  let suffix = 2;
  while (usedNames.has(candidate.toLowerCase())) {
    candidate = `${stem}-${suffix}${extension}`;
    suffix += 1;
  }
  usedNames.add(candidate.toLowerCase());
  return candidate;
}

function resolveProjectFolderPath(projectFolderUrl) {
  if (typeof projectFolderUrl !== 'string' || !projectFolderUrl.trim()) {
    throw new Error('Project folder URL is required.');
  }

  const normalized = projectFolderUrl.trim();
  const url = new URL(normalized);
  const pathname = decodeURIComponent(url.pathname);
  if (!pathname.startsWith('/')) {
    throw new Error('Project folder URL is invalid.');
  }

  return pathname.replace(/\/+$/, '');
}

function safeFileNameFromUpload(name, index, usedNames) {
  const cleaned = (name || `attachment-${index + 1}`)
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_')
    .replace(/[. ]+$/g, '')
    .slice(-180) || `attachment-${index + 1}`;

  const dot = cleaned.lastIndexOf('.');
  const extension = dot > 0 ? cleaned.slice(dot) : '';
  const stem = dot > 0 ? cleaned.slice(0, dot) : cleaned;
  const now = new Date();
  const timestamp = [
    String(now.getUTCFullYear()),
    String(now.getUTCMonth() + 1).padStart(2, '0'),
    String(now.getUTCDate()).padStart(2, '0'),
    '_',
    String(now.getUTCHours()).padStart(2, '0'),
    String(now.getUTCMinutes()).padStart(2, '0'),
    String(now.getUTCSeconds()).padStart(2, '0'),
    String(now.getUTCMilliseconds()).padStart(3, '0')
  ].join('');
  const timestampSuffix = `_${timestamp}`;
  const maxStemLength = Math.max(1, 180 - extension.length - timestampSuffix.length);
  const limitedStem = stem.slice(-maxStemLength);
  let candidate = `${limitedStem}${timestampSuffix}${extension}`;
  let suffix = 2;
  while (usedNames.has(candidate.toLowerCase())) {
    candidate = `${limitedStem}-${suffix}${timestampSuffix}${extension}`;
    suffix += 1;
  }

  usedNames.add(candidate.toLowerCase());
  return candidate;
}

function encodeGraphPath(path) {
  return path.split('/').filter(Boolean).map(encodeURIComponent).join('/');
}

async function getProjectDriveLocation(projectFolderUrl, siteId, accessToken) {
  const projectUrl = new URL(projectFolderUrl);
  if (projectUrl.origin !== `https://${siteHost}`) {
    throw new Error('Project folder URL is not on the configured SharePoint host.');
  }

  const drivesUrl = new URL(`${graphBaseUrl}/sites/${siteId}/drives`);
  drivesUrl.searchParams.set('$select', 'id,name,webUrl');
  const drivesResult = await graphRequest(drivesUrl, accessToken, null, 'document library lookup');
  const projectPath = decodeURIComponent(projectUrl.pathname).replace(/\/+$/, '');
  const matchingDrives = drivesResult.value
    .map((drive) => {
      if (!drive.webUrl || !drive.id) {
        return null;
      }
      const driveUrl = new URL(drive.webUrl);
      const drivePath = decodeURIComponent(driveUrl.pathname).replace(/\/+$/, '');
      if (
        driveUrl.origin !== projectUrl.origin
        || (projectPath !== drivePath && !projectPath.startsWith(`${drivePath}/`))
      ) {
        return null;
      }

      return {
        id: drive.id,
        relativePath: projectPath.slice(drivePath.length).replace(/^\/+/, ''),
        pathLength: drivePath.length
      };
    })
    .filter(Boolean)
    .sort((left, right) => right.pathLength - left.pathLength);

  if (!matchingDrives.length) {
    throw new Error('No document library matched the project folder URL.');
  }

  return matchingDrives[0];
}

async function getDriveItemByPath(driveId, path, accessToken, operation) {
  const drivePath = path ? `:/${encodeGraphPath(path)}` : '';
  return graphRequestIfExists(
    `${graphBaseUrl}/drives/${encodeURIComponent(driveId)}/root${drivePath}`,
    accessToken,
    operation
  );
}

async function getProjectDriveAndFolder(projectFolderUrl, siteId, accessToken) {
  const drive = await getProjectDriveLocation(projectFolderUrl, siteId, accessToken);
  const projectFolder = await getDriveItemByPath(
    drive.id,
    drive.relativePath,
    accessToken,
    'project folder lookup'
  );
  if (!projectFolder) {
    throw new Error('The project folder was not found in its document library.');
  }
  return { drive, projectFolderPath: drive.relativePath };
}

async function getDocumentLibraryDrive(siteId, libraryName, accessToken) {
  let drivesUrl = new URL(`${graphBaseUrl}/sites/${encodeURIComponent(siteId)}/drives`);
  drivesUrl.searchParams.set('$select', 'id,name');

  while (drivesUrl) {
    const page = await graphRequest(drivesUrl, accessToken, null, 'document library lookup');
    const drive = page.value.find((item) => item.name === libraryName);
    if (drive && typeof drive.id === 'string' && drive.id) {
      return drive;
    }

    if (page['@odata.nextLink']) {
      drivesUrl = new URL(page['@odata.nextLink']);
      if (drivesUrl.origin !== new URL(graphBaseUrl).origin) {
        throw new Error('Microsoft Graph returned an unexpected pagination URL');
      }
    } else {
      drivesUrl = null;
    }
  }

  throw new Error(`SharePoint document library "${libraryName}" was not found.`);
}

async function getDriveItemListItemId(driveId, driveItemId, accessToken) {
  const listItem = await graphRequest(
    `${graphBaseUrl}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(driveItemId)}/listItem?$select=id`,
    accessToken,
    null,
    'uploaded invoice list item lookup'
  );
  const listItemId = Number(listItem && listItem.id);
  if (!Number.isSafeInteger(listItemId) || listItemId < 1) {
    throw new Error('Uploaded invoice list item did not include a valid SharePoint ID.');
  }
  return listItemId;
}

async function uploadFileToDrive(driveId, file, accessToken, operation, index, usedNames) {
  const safeFileName = safeFileNameFromUpload(file.name, index, usedNames);
  const fileData = await graphRequest(
    `${graphBaseUrl}/drives/${encodeURIComponent(driveId)}/root:/${encodeGraphPath(safeFileName)}:/content`,
    accessToken,
    {
      method: 'PUT',
      headers: { 'Content-Type': file.type || 'application/octet-stream' },
      body: Buffer.from(await file.arrayBuffer())
    },
    operation,
    120000
  );
  if (!fileData || typeof fileData.id !== 'string' || !fileData.id) {
    throw new Error(`Microsoft Graph ${operation} response did not include a file item ID.`);
  }
  return { fileName: safeFileName, fileData };
}

function getInvoiceMetadata(invoiceValues, supplierId, orderId) {
  return {
    InvoiceNumber: invoiceValues.invoiceNumber,
    invoiceSum: invoiceValues.totalAmount.toFixed(2),
    supplierLookupId: Number(supplierId),
    order0LookupId: Number(orderId),
    AllocationNumber: invoiceValues.allocationNumber,
    DocumentType: invoiceValues.documentType,
    LinkedInvoiceNumberLookupId: invoiceValues.creditInvoiceNumber,
    _Comments: invoiceValues.supplierNotes,
    InvoiceDate: invoiceValues.date,
    CreatedFromSupplierApp: true
  };
}

async function uploadInvoiceFile(
  projectFolderUrl,
  file,
  siteId,
  accessToken,
  runStep,
  invoiceValues,
  supplierId,
  orderId
) {
  const { drive, projectFolderPath } = await runStep(
    'איתור ספריית המסמכים ותיקיית הפרויקט',
    () => getProjectDriveAndFolder(projectFolderUrl, siteId, accessToken)
  );
  const safeFileName = safeFileNameFromUpload(file.name, 1, new Set());
  const uploadPath = [projectFolderPath, safeFileName].filter(Boolean).join('/');
  const fileData = await runStep(
    `העלאת הקובץ ${safeFileName} באמצעות Microsoft Graph`,
    async () => {
      const fileBytes = Buffer.from(await file.arrayBuffer());
      return graphRequest(
        `${graphBaseUrl}/drives/${encodeURIComponent(drive.id)}/root:/${encodeGraphPath(uploadPath)}:/content`,
        accessToken,
        {
          method: 'PUT',
          headers: { 'Content-Type': file.type || 'application/octet-stream' },
          body: fileBytes
        },
        'invoice file upload',
        120000
      );
    }
  );

  if (!fileData || typeof fileData.id !== 'string' || !fileData.id) {
    throw new Error('Microsoft Graph upload response did not include a file item ID.');
  }

  await runStep(
    'עדכון שדות החשבונית והזמנת הרכש',
    () => graphRequest(
      `${graphBaseUrl}/drives/${encodeURIComponent(drive.id)}/items/${encodeURIComponent(fileData.id)}/listItem/fields`,
      accessToken,
      {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(getInvoiceMetadata(invoiceValues, supplierId, orderId))
      },
      'invoice metadata update'
    )
  );
  const invoiceListItemId = await runStep(
    'קבלת מזהה פריט החשבונית לקישור הצרופות',
    () => getDriveItemListItemId(drive.id, fileData.id, accessToken)
  );

  return {
    listItemId: invoiceListItemId,
    fileName: safeFileName,
    fileUrl: fileData.webUrl || null,
    folderServerRelativeUrl: resolveProjectFolderPath(projectFolderUrl)
  };
}

async function uploadSupportingDocuments(
  files,
  siteId,
  accessToken,
  runStep,
  invoiceValues,
  supplierId,
  orderId,
  invoiceListItemId
) {
  if (!files.length) {
    return [];
  }

  const drive = await runStep(
    `איתור ספריית ${additionalDocumentsLibraryName}`,
    () => getDocumentLibraryDrive(siteId, additionalDocumentsLibraryName, accessToken)
  );
  const usedNames = new Set();
  const uploadedFiles = [];

  for (const [index, file] of files.entries()) {
    const uploaded = await runStep(
      `העלאת הצרופה ${file.name} לספריית ${additionalDocumentsLibraryName}`,
      () => uploadFileToDrive(
        drive.id,
        file,
        accessToken,
        'additional document upload',
        index + 1,
        usedNames
      )
    );
    await runStep(
      `עדכון פרטי הצרופה וקישור לחשבונית ${uploaded.fileName}`,
      () => graphRequest(
        `${graphBaseUrl}/drives/${encodeURIComponent(drive.id)}/items/${encodeURIComponent(uploaded.fileData.id)}/listItem/fields`,
        accessToken,
        {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            ...getInvoiceMetadata(invoiceValues, supplierId, orderId),
            InvoiceLookupLookupId: invoiceListItemId
          })
        },
        'additional document metadata update'
      )
    );
    uploadedFiles.push(uploaded.fileName);
  }

  return uploadedFiles;
}

app.http('submitInvoice', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const completedSteps = [];
    let currentStep = 'בדיקת הגדרות SharePoint';
    const failureResponse = (status, error, diagnostic) => jsonResponse(status, {
      error,
      completedSteps,
      failedStep: currentStep,
      ...(diagnostic ? { diagnostic } : {})
    });
    const runStep = async (label, operation) => {
      currentStep = label;
      const result = await operation();
      completedSteps.push(label);
      return result;
    };

    const credentials = getCredentials();
    if (!credentials) {
      context.error('Invoice validation is not configured: SharePoint credentials are missing.');
      return failureResponse(503, 'בדיקת התאמת החשבונית אינה זמינה כרגע.');
    }
    completedSteps.push('הגדרות SharePoint קיימות');

    try {
      const formData = await runStep('קריאת נתוני הטופס', () => request.formData());
      currentStep = 'בדיקת שדות החשבונית והקובץ';
      const supplierNumber = formData.get('verifiedNumber');
      if (typeof supplierNumber !== 'string' || !supplierNumber.trim()) {
        return failureResponse(400, 'מספר הספק חסר.');
      }

      const invoiceFile = formData.get('invoiceFile');
      if (!isFile(invoiceFile)) {
        return failureResponse(400, 'יש לצרף את קובץ החשבונית.');
      }

      const supportingFiles = formData.getAll('supportingFiles')
        .filter((file) => isFile(file) && file.name);
      if (supportingFiles.some((file) => file.size > maxUploadBytes)) {
        return failureResponse(400, 'גודל אחת מהצרופות הנוספות עולה על 50 MB.');
      }

      if (invoiceFile.size > maxUploadBytes) {
        return failureResponse(400, 'גודל קובץ החשבונית עולה על 50 MB.');
      }

      const invoiceValues = getInvoiceValues(formData);
      if (!invoiceValues) {
        return failureResponse(400, 'פרטי החשבונית אינם תקינים.');
      }

      const projectFolderUrl = getTextField(formData, 'projectFolderUrl');
      const supplierId = getTextField(formData, 'supplierId');
      const orderId = invoiceValues.purchaseOrder;
      if (!projectFolderUrl) {
        return failureResponse(400, 'יש לבחור הזמנת רכש תקינה לפני העלאת החשבונית.');
      }
      if (
        !/^[1-9]\d*$/.test(supplierId)
        || !/^[1-9]\d*$/.test(orderId)
        || !Number.isSafeInteger(Number(supplierId))
        || !Number.isSafeInteger(Number(orderId))
      ) {
        return failureResponse(400, 'פרטי הספק או הזמנת הרכש אינם תקינים.');
      }
      completedSteps.push(currentStep);

      const graphToken = await runStep(
        'קבלת אסימון Microsoft Graph',
        () => getAccessToken(credentials, 'https://graph.microsoft.com/.default')
      );
      const site = await runStep(
        'איתור אתר SharePoint',
        () => getSharePointSite(graphToken)
      );
      const verifiedSupplierId = await runStep(
        'אימות הספק מול רשימת הספקים',
        () => getVerifiedSupplierId(site.id, invoiceValues, graphToken)
      );
      if (verifiedSupplierId !== supplierId) {
        currentStep = 'אימות מזהה הספק שנבחר';
        return failureResponse(403, 'פרטי הספק לא אומתו מול רשימת הספקים.');
      }
      if (
        invoiceValues.documentType === 'חשבונית זיכוי'
        && !await runStep(
          'אימות מספר החשבונית המקורית מול החשבוניות של הספק',
          () => supplierInvoiceExists(
            site.id,
            invoiceValues.supplierName,
            invoiceValues.creditInvoiceNumber,
            graphToken
          )
        )
      ) {
        currentStep = 'אימות מספר החשבונית המקורית';
        return failureResponse(400, 'מספר החשבונית המקורית שנבחרה אינו מופיע בחשבוניות של הספק.');
      }
      const verifiedPurchaseOrder = await runStep(
        'אימות הזמנת הרכש ויתרת הסכום מול הספק',
        () => getVerifiedPurchaseOrder(site.id, orderId, invoiceValues.supplierName, graphToken)
      );
      if (!verifiedPurchaseOrder.belongsToSupplier) {
        currentStep = 'אימות הזמנת הרכש שנבחרה';
        return failureResponse(403, 'הזמנת הרכש אינה משויכת לספק המאומת.');
      }
      if (invoiceValues.totalAmount > verifiedPurchaseOrder.remainingAmount) {
        currentStep = 'בדיקת סכום החשבונית מול יתרת הזמנת הרכש';
        return failureResponse(
          400,
          `סכום החשבונית גבוה מהיתרה שנותרה להזמנה (${verifiedPurchaseOrder.remainingAmount.toFixed(2)} ₪).`
        );
      }
      const matchingTaskExists = await runStep(
        'בדיקת חשבונית כפולה ברשימת המשימות',
        () => financeTaskExists(
          site.id,
          { invoiceNumber: invoiceValues.invoiceNumber, supplierName: invoiceValues.supplierName },
          graphToken
        )
      );
      if (matchingTaskExists) {
        return jsonResponse(409, {
          completedSteps,
          error: `חשבונית מספר ${invoiceValues.invoiceNumber} כבר הוגשה עבור ספק ${invoiceValues.supplierName}.\nאי אפשר להעלות את החשבונית.`
        });
      }

      const uploadResult = await uploadInvoiceFile(
        projectFolderUrl,
        invoiceFile,
        site.id,
        graphToken,
        runStep,
        invoiceValues,
        verifiedSupplierId,
        orderId
      );
      const uploadedSupportingFiles = await uploadSupportingDocuments(
        supportingFiles,
        site.id,
        graphToken,
        runStep,
        invoiceValues,
        verifiedSupplierId,
        orderId,
        uploadResult.listItemId
      );

      return jsonResponse(200, {
        success: true,
        validated: true,
        uploaded: true,
        fileName: uploadResult.fileName,
        fileUrl: uploadResult.fileUrl,
        folderServerRelativeUrl: uploadResult.folderServerRelativeUrl,
        supportingFiles: uploadedSupportingFiles,
        completedSteps
      });
    } catch (error) {
      context.error('Invoice validation failed.', error);
      const statusMatch = error && typeof error.message === 'string'
        ? error.message.match(/\((\d{3})\)/)
        : null;
      const diagnostic = error instanceof GraphRequestError
        ? `HTTP ${error.status}`
        : statusMatch
          ? `HTTP ${statusMatch[1]}`
          : error && error.name === 'TypeError'
            ? 'שגיאת רשת או כתובת'
            : 'שגיאה פנימית; יש לבדוק את יומן השרת';

      if (error instanceof GraphRequestError) {
        return failureResponse(
          502,
          `בדיקת SharePoint נכשלה בשלב ${error.operation} (HTTP ${error.status}).`,
          diagnostic
        );
      }

      return failureResponse(
        502,
        'לא ניתן לבדוק את התאמת החשבונית כרגע. נא לבדוק את הגדרות SharePoint ולנסות שוב.',
        diagnostic
      );
    }
  }
});
