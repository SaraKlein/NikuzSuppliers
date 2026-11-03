import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

import styles from './MyInvoicesWebPart.module.scss';

interface IMyInvoicesWebPartProps { }

interface ICurrentUser {
  Email?: string;
}

interface IInvoiceFileItem {
  Id: number;
  File?: IInvoiceFile;
}

interface IInvoiceFile {
  ServerRelativeUrl?: string;
}

interface IInvoice {
  Id: number;
  Title?: string;
  InvoiceNumber?: string;
  InvoiceDate?: string;
  InvoiceAmount?: string | number;
  InvoiceItemId?: number;
  PurchaseOrderTitle?: string;
  SupplierName?: string;
  DocumentType?: string;
  ApprovalStatus?: string;
  File?: IInvoiceFile;
}

interface IVerboseCollection<T> {
  results: T[];
  __next?: string;
}

interface IODataEnvelope<T> {
  d: T | IVerboseCollection<T>;
}

export default class MyInvoicesWebPart extends BaseClientSideWebPart<IMyInvoicesWebPartProps> {
  public render(): void {
    this.domElement.innerHTML = `
      <section class="${styles.myInvoices}" aria-labelledby="my-invoices-title">
        <h1 class="${styles.heading}" id="my-invoices-title">החשבוניות שלי</h1>
        <p class="${styles.subtitle}">חשבוניות שהוגשו עבור הזמנות רכש שיצרת.</p>
        <div class="${styles.loading}" id="invoices-loading" role="status" aria-live="polite">
          טוען את החשבוניות...
        </div>
        <div class="${styles.error}" id="invoices-error" role="alert" hidden></div>
        <p class="${styles.empty}" id="invoices-empty" hidden>לא נמצאו חשבוניות להזמנות שיצרת.</p>
        <ul class="${styles.invoiceList}" id="invoice-list" aria-label="רשימת החשבוניות" hidden></ul>
      </section>`;

    this._loadInvoices().catch((error: unknown) => {
      console.error('Could not load invoices for the current user.', error);
      this._showError(error instanceof Error ? error.message : 'לא ניתן לטעון את החשבוניות כרגע.');
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private async _loadInvoices(): Promise<void> {
    const currentUser = await this._getJson<ICurrentUser>(
      this._restUrl('currentuser?$select=Email')
    );
    const currentUserEmail = currentUser.Email?.trim();
    if (!currentUserEmail) {
      throw new Error('לא ניתן לזהות את המשתמש המחובר.');
    }

    const invoicesUrl = this._collectionUrl(
      'UploadedInvoices',
      'Id,Title,InvoiceNumber,InvoiceDate,InvoiceAmount,InvoiceItemId,' +
      'PurchaseOrderTitle,SupplierName,DocumentType,ApprovalStatus',
      `InitiatorEmail eq '${currentUserEmail.replace(/'/g, "''")}'`
    );
    const invoices = await this._getAll<IInvoice>(invoicesUrl);
    if (!invoices.length) {
      this._showEmpty();
      return;
    }

    const fileUrls = await this._getInvoiceFileUrls(
      invoices
        .map((invoice) => invoice.InvoiceItemId)
        .filter((id): id is number =>
          typeof id === 'number' && Number.isSafeInteger(id) && id > 0
        )
    );
    invoices.forEach((invoice) => {
      const fileUrl = invoice.InvoiceItemId
        ? fileUrls.get(invoice.InvoiceItemId)
        : undefined;
      if (fileUrl) {
        invoice.File = { ServerRelativeUrl: fileUrl };
      }
    });

    invoices.sort((left, right) => {
      const leftDate = left.InvoiceDate ? Date.parse(left.InvoiceDate) : 0;
      const rightDate = right.InvoiceDate ? Date.parse(right.InvoiceDate) : 0;
      return rightDate - leftDate;
    });

    this._element('invoices-loading').hidden = true;
    if (!invoices.length) {
      this._showEmpty();
      return;
    }

    this._renderInvoices(invoices);
  }

  private async _getInvoiceFileUrls(invoiceItemIds: number[]): Promise<Map<number, string>> {
    const fileUrls = new Map<number, string>();
    const uniqueIds = Array.from(new Set(invoiceItemIds));

    for (let offset = 0; offset < uniqueIds.length; offset += 50) {
      const batch = uniqueIds.slice(offset, offset + 50);
      const filter = batch.map((id) => `Id eq ${id}`).join(' or ');
      const url = this._collectionUrl(
        'מסמכי פרויקט',
        'Id,File/ServerRelativeUrl',
        filter,
        'File'
      );
      const items = await this._getAll<IInvoiceFileItem>(url);
      items.forEach((item) => {
        if (item.File?.ServerRelativeUrl) {
          fileUrls.set(item.Id, item.File.ServerRelativeUrl);
        }
      });
    }

    return fileUrls;
  }

  private _collectionUrl(
    listTitle: string,
    select: string,
    filter: string,
    expand?: string
  ): string {
    const url = new URL(
      this._restUrl(`lists/getbytitle('${listTitle.replace(/'/g, "''")}')/items`)
    );
    url.searchParams.set('$select', select);
    url.searchParams.set('$filter', filter);
    url.searchParams.set('$top', '5000');
    if (expand) {
      url.searchParams.set('$expand', expand);
    }
    return url.toString();
  }

  private _renderInvoices(invoices: IInvoice[]): void {
    const list = this._element<HTMLUListElement>('invoice-list');
    list.replaceChildren();

    invoices.forEach((invoice) => {
      const item = document.createElement('li');
      item.className = styles.invoice;

      const main = document.createElement('div');
      main.className = styles.invoiceMain;
      const title = document.createElement('strong');
      title.className = styles.invoiceTitle;
      title.textContent = invoice.InvoiceNumber
        ? `חשבונית ${invoice.InvoiceNumber}`
        : invoice.Title || `חשבונית ${invoice.Id}`;
      const supplier = document.createElement('span');
      supplier.className = styles.invoiceMeta;
      supplier.textContent = invoice.SupplierName || invoice.DocumentType || '';
      main.append(title, supplier);

      item.append(
        main,
        this._field('הזמנת רכש', invoice.PurchaseOrderTitle || '—'),
        this._field('תאריך', this._formatDate(invoice.InvoiceDate)),
        this._field('סכום', this._formatAmount(invoice.InvoiceAmount))
      );

      const status = document.createElement('span');
      status.className = styles.status;
      status.textContent = this._formatStatus(invoice.ApprovalStatus);
      const normalizedStatus = (invoice.ApprovalStatus || '').trim().toLowerCase();
      if (['approve', 'approved', 'מאושר', 'מאושרת'].indexOf(normalizedStatus) !== -1) {
        status.classList.add(styles.statusApproved);
      } else if (['reject', 'rejected', 'נדחה', 'נדחתה'].indexOf(normalizedStatus) !== -1) {
        status.classList.add(styles.statusRejected);
      }
      status.setAttribute('aria-label', `סטטוס: ${status.textContent}`);
      const statusActions = document.createElement('div');
      statusActions.className = styles.statusActions;
      statusActions.append(status);

      if (
        invoice.ApprovalStatus?.toLowerCase() === 'חדש' &&
        typeof invoice.InvoiceItemId === 'number' &&
        Number.isSafeInteger(invoice.InvoiceItemId) &&
        invoice.InvoiceItemId > 0
      ) {
        const approvalButton = document.createElement('a');
        approvalButton.className = styles.approvalButton;
        const approvalUrl = new URL(
          `${this.context.pageContext.web.absoluteUrl}/SitePages/InvoiceApprovalPage.aspx`
        );
        approvalUrl.searchParams.set('itemId', String(invoice.InvoiceItemId));
        approvalButton.href = approvalUrl.toString();
        approvalButton.target = '_blank';
        approvalButton.rel = 'noopener noreferrer';
        approvalButton.textContent = 'אישור / דחייה';
        statusActions.append(approvalButton);
      }
      item.append(statusActions);

      if (invoice.File?.ServerRelativeUrl) {
        const link = document.createElement('a');
        link.className = styles.openFile;
        link.href = new URL(invoice.File.ServerRelativeUrl, window.location.origin).toString();
        link.target = '_blank';
        link.rel = 'noopener noreferrer';
        link.textContent = 'פתיחת החשבונית';
        item.append(link);
      }

      list.append(item);
    });

    list.hidden = false;
  }

  private _field(label: string, value: string): HTMLDivElement {
    const field = document.createElement('div');
    field.className = styles.field;
    const fieldLabel = document.createElement('span');
    fieldLabel.className = styles.fieldLabel;
    fieldLabel.textContent = label;
    const fieldValue = document.createElement('span');
    fieldValue.className = styles.fieldValue;
    fieldValue.textContent = value;
    field.append(fieldLabel, fieldValue);
    return field;
  }

  private _formatDate(value?: string): string {
    if (!value) {
      return '—';
    }
    const date = new Date(value);
    return Number.isNaN(date.getTime())
      ? value
      : new Intl.DateTimeFormat('he-IL', {
        year: 'numeric',
        month: 'numeric',
        day: 'numeric'
      }).format(date);
  }

  private _formatAmount(value?: string | number): string {
    if (value === undefined || value === null || String(value).trim() === '') {
      return '—';
    }
    const amount = Number(value);
    return Number.isFinite(amount)
      ? new Intl.NumberFormat('he-IL', { style: 'currency', currency: 'ILS' }).format(amount)
      : String(value);
  }

  private _formatStatus(value?: string): string {
    const normalized = (value || '').trim().toLowerCase();
    const labels: Record<string, string> = {
      approve: 'אושרה',
      approved: 'אושרה',
      reject: 'נדחתה',
      rejected: 'נדחתה',
      'חדש': 'חדשה',
      'מאושר': 'אושרה',
      'מאושרת': 'אושרה',
      'נדחה': 'נדחתה',
      'נדחתה': 'נדחתה'
    };
    return labels[normalized] || value || 'ללא סטטוס';
  }

  private _restUrl(path: string): string {
    return `${this.context.pageContext.web.absoluteUrl}/_api/web/${path}`;
  }

  private async _getJson<T>(url: string): Promise<T> {
    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: new Headers({
          Accept: 'application/json;odata=verbose',
          'OData-Version': '3.0'
        })
      }
    );
    if (!response.ok) {
      throw new Error(`SharePoint request failed (${response.status}).`);
    }

    const payload = await response.json() as T | IODataEnvelope<T>;
    return typeof payload === 'object' && payload !== null && 'd' in payload
      ? payload.d as T
      : payload;
  }

  private async _getAll<T>(url: string): Promise<T[]> {
    const results: T[] = [];
    let nextUrl: string | undefined = url;
    const siteOrigin = new URL(this.context.pageContext.web.absoluteUrl).origin;

    while (nextUrl) {
      const pageUrl = new URL(nextUrl, this.context.pageContext.web.absoluteUrl);
      if (pageUrl.origin !== siteOrigin) {
        throw new Error('SharePoint returned an unexpected paging URL.');
      }
      const payload = await this._getJson<unknown>(pageUrl.toString());
      if (!payload || typeof payload !== 'object') {
        throw new Error('SharePoint returned an invalid invoice list response.');
      }

      const page = payload as Partial<IVerboseCollection<T>> & {
        value?: T[];
        '@odata.nextLink'?: string;
      };
      const pageResults = Array.isArray(page.results)
        ? page.results
        : page.value;
      if (!Array.isArray(pageResults)) {
        throw new Error('SharePoint returned an invalid invoice list response.');
      }

      results.push(...pageResults);
      nextUrl = page.__next || page['@odata.nextLink'];
    }

    return results;
  }

  private _element<T extends HTMLElement>(id: string): T {
    const element = this.domElement.querySelector<T>(`#${id}`);
    if (!element) {
      throw new Error(`Required invoice list element "${id}" was not found.`);
    }
    return element;
  }

  private _showEmpty(): void {
    this._element('invoices-loading').hidden = true;
    this._element('invoices-empty').hidden = false;
  }

  private _showError(message: string): void {
    this._element('invoices-loading').hidden = true;
    const error = this._element<HTMLDivElement>('invoices-error');
    error.textContent = message;
    error.hidden = false;
  }
}
