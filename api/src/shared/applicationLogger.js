const { randomUUID } = require('node:crypto');

const graphBaseUrl = 'https://graph.microsoft.com/v1.0';
const applicationName = 'SupplierInvoices';
const logListName = 'External Application Logs';
const environments = new Set(['Development', 'Test', 'Staging', 'Production']);
const severities = new Set(['Debug', 'Information', 'Warning', 'Error', 'Critical']);

function getEnvironment() {
  const environment = process.env.APP_ENVIRONMENT || 'Production';
  if (!environments.has(environment)) {
    throw new Error('APP_ENVIRONMENT must be Development, Test, Staging, or Production.');
  }
  return environment;
}

async function graphRequest(url, accessToken, options, operation) {
  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options && options.headers)
    },
    signal: AbortSignal.timeout(15000)
  });

  if (!response.ok) {
    let graphCode;
    let graphMessage;
    try {
      const body = await response.json();
      const graphError = body && body.error;
      graphCode = typeof (graphError && graphError.code) === 'string'
        ? graphError.code.slice(0, 100)
        : undefined;
      graphMessage = typeof (graphError && graphError.message) === 'string'
        ? graphError.message.replace(/[\r\n\t]+/g, ' ').slice(0, 500)
        : undefined;
    } catch {
      // The HTTP status is still useful when Graph does not return a JSON error.
    }

    const error = new Error(
      `Application log ${operation} failed (HTTP ${response.status})`
      + (graphCode ? `: ${graphCode}` : '')
      + (graphMessage ? ` — ${graphMessage}` : '')
    );
    error.status = response.status;
    error.graphCode = graphCode;
    error.graphMessage = graphMessage;
    error.operation = operation;
    throw error;
  }

  if (response.status === 204) {
    return null;
  }
  return response.json();
}

async function getLogListId(siteId, accessToken) {
  const listsUrl = new URL(`${graphBaseUrl}/sites/${encodeURIComponent(siteId)}/lists`);
  listsUrl.searchParams.set('$filter', `displayName eq '${logListName}'`);
  listsUrl.searchParams.set('$select', 'id,displayName');
  const result = await graphRequest(listsUrl, accessToken, null, 'list lookup');
  const list = result.value.find((item) => item.displayName === logListName);
  if (!list || typeof list.id !== 'string' || !list.id) {
    throw new Error(`SharePoint list "${logListName}" was not found.`);
  }
  return list.id;
}

async function writeApplicationLog({
  siteId,
  accessToken,
  severity,
  message,
  details,
  correlationId,
  httpStatus
}) {
  let operation = 'validate log fields';
  try {
    if (!severities.has(severity)) {
      throw new Error(`Unsupported application log severity: ${severity}`);
    }

    operation = 'find log list';
    const listId = await getLogListId(siteId, accessToken);
    const fields = {
      Title: randomUUID(),
      Application: applicationName,
      Environment: 'Production',//getEnvironment(),
      Severity: severity,
      OccurredAt: new Date().toISOString(),
      Message: String(message).slice(0, 2000),
      Details: details === undefined ? '' : JSON.stringify(details).slice(0, 12000),
      CorrelationId: String(correlationId || randomUUID()).slice(0, 255)
    };
    if (Number.isInteger(httpStatus)) {
      fields.HttpStatus = httpStatus;
    }

    operation = 'create log item';
    await graphRequest(
      `${graphBaseUrl}/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items`,
      accessToken,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields })
      },
      'write'
    );
  } catch (error) {
    console.error('[ApplicationLogger] Failed to write SharePoint log.', {
      operation: error.operation || operation,
      message: error.message,
      httpStatus: error.status,
      graphCode: error.graphCode,
      graphMessage: error.graphMessage
    });
    throw error;
  }
}

module.exports = { writeApplicationLog };
