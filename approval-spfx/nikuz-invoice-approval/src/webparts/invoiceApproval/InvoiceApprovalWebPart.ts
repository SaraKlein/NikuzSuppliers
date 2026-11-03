import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import styles from './InvoiceApprovalWebPart.module.scss';

interface IInvoiceApprovalWebPartProps { }

interface ILookup {
  Id: number;
  Title?: string;
}

interface IFile {
  ServerRelativeUrl?: string;
}

interface IUser {
  Id: number;
  Title: string;
  EMail?: string;
}

interface IInvoiceItem {
  Id: number;
  Title?: string;
  InvoiceNumber?: string;
  InvoiceDate?: string;
  invoiceSum?: string | number;
  SupplierName?: string;
  AllocationNumber?: string;
  DocumentType?: string;
  LinkedInvoiceNumber?: string;
  OData__Comments?: string;
  InvoiceStatus?: string;
  OrderInitiatorComments?: string;
  order0?: ILookup;
  File?: IFile;
  '@odata.etag'?: string;
  supplier: ILookup;
}

interface IUploadedInvoiceItem {
  Id: number;
  InvoiceItemId: number;
}

interface IUploadedInvoiceCollection {
  results?: IUploadedInvoiceItem[];
  value?: IUploadedInvoiceItem[];
}

interface IODataVerboseResponse<T> {
  d: T;
}

interface IOrderItem {
  Id: number;
  Title?: string;
  Author?: IUser;
}

export default class InvoiceApprovalWebPart extends BaseClientSideWebPart<IInvoiceApprovalWebPartProps> {
  private _invoice?: IInvoiceItem;
  private _uploadedInvoiceItemId?: number;
  private _submitting: boolean = false;

  public render(): void {
    this.domElement.innerHTML = `
      <section class="${styles.approval}" dir="rtl" aria-labelledby="approval-title">
        <div class="${styles.card}">
          <div class="${styles.loading}" id="approval-loading" role="status">טוען את פרטי החשבונית...</div>
          <div class="${styles.message}" id="approval-message" role="status" hidden></div>
          <div id="approval-content" hidden>
            <h1 id="approval-title">אישור חשבונית</h1>
            <p class="${styles.subtitle}">בדקו את פרטי החשבונית והקובץ המצורף, ולאחר מכן בחרו את ההחלטה.</p>
            <div class="${styles.layout}">
              <section class="${styles.previewPanel}" aria-labelledby="preview-title">
                <h2 id="preview-title">תצוגת החשבונית</h2>
                <div class="${styles.preview}" id="invoice-preview"></div>
                <a class="${styles.openFile}" id="open-invoice" target="_blank" rel="noopener noreferrer" hidden>
                  פתיחת הקובץ בחלון חדש
                </a>
              </section>
              <section aria-labelledby="details-title">
                <h2 id="details-title">פרטי החשבונית</h2>
                <dl class="${styles.details}" id="invoice-details"></dl>
                <p class="${styles.decisionMessage}" id="existing-decision" role="status" hidden></p>
                <form id="decision-form">
                  <fieldset class="${styles.fieldset}" id="decision-options">
                    <legend>החלטה</legend>
                    <div class="${styles.choices}">
                      <label class="${styles.choice}"><input type="radio" name="decision" value="approve" required> לאשר</label>
                      <label class="${styles.choice}"><input type="radio" name="decision" value="reject" required> לדחות</label>
                    </div>
                  </fieldset>
                  <label class="${styles.commentLabel}" id="comments-label" for="comments" hidden>הערה — נדרשת בדחייה</label>
                  <textarea class="${styles.comments}" id="comments" name="comments" maxlength="10000" hidden></textarea>
                  <div class="${styles.actions}" id="decision-actions">
                    <button class="${styles.submit}" type="submit" id="submit-decision">שליחת החלטה</button>
                  </div>
                  <p class="${styles.notice}" id="decision-notice" role="status" aria-live="polite"></p>
                </form>
              </section>
            </div>
          </div>
        </div>
      </section>`;

    this._bindForm();
    this._loadInvoice().catch((error: unknown) => {
      console.error('Invoice approval initialization failed.', error);
      this._showMessage(
        error instanceof Error ? error.message : 'לא ניתן לטעון את החשבונית כרגע.'
      );
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _element<T extends HTMLElement>(id: string): T {
    const element = this.domElement.querySelector<T>(`#${id}`);
    if (!element) {
      throw new Error(`Required approval page element "${id}" was not found.`);
    }
    return element;
  }

  private _bindForm(): void {
    const form = this._element<HTMLFormElement>('decision-form');
    const comments = this._element<HTMLTextAreaElement>('comments');
    const commentsLabel = this._element<HTMLLabelElement>('comments-label');
    const choices = Array.from(form.querySelectorAll<HTMLInputElement>('input[name="decision"]'));

    choices.forEach((choice) => {
      choice.addEventListener('change', () => {
        const rejecting = choices.some((input) => input.checked && input.value === 'reject');
        comments.hidden = !rejecting;
        commentsLabel.hidden = !rejecting;
        comments.required = rejecting;
        if (!rejecting) {
          comments.value = '';
        }
      });
    });

    form.addEventListener('submit', (event) => {
      event.preventDefault();
      this._submitDecision().catch((error: unknown) => {
        console.error('Invoice decision submission failed unexpectedly.', error);
        this._element<HTMLParagraphElement>('decision-notice').textContent =
          error instanceof Error ? error.message : 'שליחת ההחלטה נכשלה.';
      });
    });
  }

  private async _loadInvoice(): Promise<void> {
    console.log('Loading invoice for approval...');
    const itemId = new URL(window.location.href).searchParams.get('itemId');
    if (!/^[1-9]\d*$/.test(itemId || '')) {
      this._showMessage('קישור החשבונית אינו תקין.');
      return;
    }

    try {
      const invoiceReferenceUrl = this._restUrl(
        `lists/getbytitle('מסמכי פרויקט')/items(${itemId})` +
        '?$select=Id,order0/Id&$expand=order0'
      );
      console.log('Fetching invoice order reference from SharePoint:', invoiceReferenceUrl);
      const invoiceReference = await this._getJson<IInvoiceItem>(invoiceReferenceUrl);
      if (!invoiceReference.order0 || !Number.isSafeInteger(invoiceReference.order0.Id)) {
        throw new Error('לחשבונית לא משויכת הזמנת רכש תקינה.');
      }

      const orderUrl = this._restUrl(
        `lists/getbytitle('הזמנות')/items(${invoiceReference.order0.Id})` +
        '?$select=Id,Title,Author/Id,Author/Title,Author/EMail&$expand=Author'
      );
      console.log('Fetching order from SharePoint:', orderUrl);
      const order = await this._getJson<IOrderItem>(orderUrl);
      console.log('Fetched order:', order);
      const currentUser = await this._getJson<{ Id: number }>(
        this._restUrl('currentuser?$select=Id')
      );
      console.log('Current user:', currentUser);
      if (!order.Author || order.Author.Id !== currentUser.Id) {
        this._showMessage('המשתמש הנוכחי אינו מורשה לצפות בחשבונית זו.');
        return;
      }

      const invoiceUrl = this._restUrl(
        `lists/getbytitle('מסמכי פרויקט')/items(${itemId})` +
        '?$select=Id,Title,InvoiceNumber,OData__Comments,InvoiceDate,invoiceSum,AllocationNumber,DocumentType,' +
        'LinkedInvoiceNumber,InvoiceStatus,OrderInitiatorComments,order0/Id,order0/Title,' +
        'supplier/Id,supplier/Title,' +
        'File/ServerRelativeUrl&$expand=order0,File,supplier'
      );
      console.log('Fetching authorized invoice details from SharePoint:', invoiceUrl);
      const invoice = await this._getJson<IInvoiceItem>(invoiceUrl);
      this._uploadedInvoiceItemId = await this._getUploadedInvoiceItemId(invoice.Id);

      this._invoice = invoice;
      this._renderInvoice(invoice);
      this._showExistingDecision(invoice.InvoiceStatus || '');
      this._element('approval-loading').hidden = true;
      this._element('approval-content').hidden = false;
    } catch (error) {
      console.error('Could not load invoice for approval.', error);
      this._showMessage(
        error instanceof Error ? error.message : 'לא ניתן לטעון את החשבונית כרגע.'
      );
    }
  }

  private _renderInvoice(invoice: IInvoiceItem): void {
    const details = this._element<HTMLDListElement>('invoice-details');
    const rows: Array<[string, string]> = [
      ['מספר חשבונית', invoice.InvoiceNumber || '—'],
      ['תאריך חשבונית', this._formatDate(invoice.InvoiceDate)],
      ['סכום חשבונית', invoice.invoiceSum === undefined || invoice.invoiceSum === null
        ? '—'
        : String(invoice.invoiceSum)],
      ['שם הספק', invoice.supplier?.Title || '—'],
      ['מספר הזמנת רכש', invoice.order0?.Title || '—'],
      ['מספר הקצאה', invoice.AllocationNumber || '—'],
      ['סוג מסמך', invoice.DocumentType || '—'],
      ['חשבונית מקורית', invoice.LinkedInvoiceNumber || '—'],
      ['סטטוס נוכחי', invoice.InvoiceStatus || '—'],
      ['הערות הספק', invoice.OData__Comments || '—'],
      ['הערת החלטה', invoice.OrderInitiatorComments || '—']
    ];

    details.replaceChildren();
    rows.forEach(([label, value]) => {
      const row = document.createElement('div');
      row.className = styles.detail;
      const heading = document.createElement('dt');
      heading.textContent = label;
      const content = document.createElement('dd');
      content.textContent = value;
      row.append(heading, content);
      details.append(row);
    });

    const preview = this._element<HTMLDivElement>('invoice-preview');
    const openLink = this._element<HTMLAnchorElement>('open-invoice');
    const fileUrl = invoice.File?.ServerRelativeUrl;
    if (!fileUrl) {
      preview.textContent = 'קישור לתצוגת החשבונית אינו זמין.';
      return;
    }

    const absoluteFileUrl = new URL(fileUrl, window.location.origin).toString();
    const frame = document.createElement('iframe');
    frame.src = absoluteFileUrl;
    frame.title = invoice.Title || invoice.InvoiceNumber || 'תצוגת חשבונית';
    preview.append(frame);
    openLink.href = absoluteFileUrl;
    openLink.hidden = false;
  }

  private _showExistingDecision(status: string): void {
    const normalized = status.trim().toLowerCase();
    const approved = ['approve', 'approved', 'מאושר', 'מאושרת'];
    const rejected = ['reject', 'rejected', 'נדחה', 'נדחתה'];
    const message = approved.indexOf(normalized) !== -1
      ? 'החשבונית כבר אושרה.'
      : rejected.indexOf(normalized) !== -1
        ? 'החשבונית כבר נדחתה.'
        : '';
    if (!message) {
      return;
    }

    this._element('decision-options').hidden = true;
    this._element('comments-label').hidden = true;
    this._element<HTMLTextAreaElement>('comments').hidden = true;
    this._element('decision-actions').hidden = true;
    const decisionMessage = this._element<HTMLParagraphElement>('existing-decision');
    decisionMessage.textContent = message;
    decisionMessage.hidden = false;
  }

  private async _submitDecision(): Promise<void> {
    console.log("submit decision");
    if (this._submitting) {
      return;
    }

    if (!this._invoice) {
      return;
    }

    const notice = this._element<HTMLParagraphElement>('decision-notice');
    if (!this._uploadedInvoiceItemId) {
      notice.textContent =
        'לא נמצא פריט מעקב לחשבונית. לא ניתן לשמור את ההחלטה.';
      return;
    }

    const selected = this.domElement.querySelector<HTMLInputElement>('input[name="decision"]:checked');
    const comments = this._element<HTMLTextAreaElement>('comments').value.trim();
    if (!selected || (selected.value !== 'approve' && selected.value !== 'reject')) {
      notice.textContent = 'יש לבחור אישור או דחייה.';
      return;
    }
    if (selected.value === 'reject' && !comments) {
      notice.textContent = 'נדרש להוסיף הערה בעת דחיית החשבונית.';
      this._element<HTMLTextAreaElement>('comments').focus();
      return;
    }
    const decision = selected.value === 'approve' ? 'approve' : 'reject';

    this._submitting = true;
    const button = this._element<HTMLButtonElement>('submit-decision');
    button.disabled = true;
    button.textContent = 'שולח החלטה...';
    notice.textContent = '';
    let invoiceUpdated = false;

    try {
     /* const etag = this._invoice['@odata.etag'];
      if (!etag) {
        throw new Error('לא ניתן לאמת את גרסת החשבונית. טענו מחדש את הדף ונסו שוב.');
      }*/

      const response = await this.context.spHttpClient.post(
        this._restUrl(`lists/getbytitle('מסמכי פרויקט')/items(${this._invoice.Id})`),
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: JSON.stringify({
            InvoiceStatus: decision,
            OrderInitiatorComments: comments
          })
        }
      );
      if (!response.ok) {
        throw new Error(await this._responseError(response));
      }
      invoiceUpdated = true;

      const trackingResponse = await this.context.spHttpClient.post(
        this._restUrl(
          `lists/getbytitle('UploadedInvoices')/items(${this._uploadedInvoiceItemId})`
        ),
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: JSON.stringify({
            ApprovalStatus: decision,
            DecisionComment: comments
          })
        }
      );
      if (!trackingResponse.ok) {
        throw new Error(await this._responseError(trackingResponse));
      }

      const message = decision === 'approve'
        ? 'החשבונית אושרה בהצלחה.'
        : 'החשבונית נדחתה בהצלחה וההערה נשמרה.';
      notice.textContent = message;
      notice.classList.add(styles.success);
      this._element('decision-options').hidden = true;
      this._element('comments-label').hidden = true;
      this._element<HTMLTextAreaElement>('comments').hidden = true;
      this._element('decision-actions').hidden = true;
      this._element('invoice-details').querySelectorAll('dt').forEach((term) => {
        if (term.textContent === 'סטטוס נוכחי') {
          const value = term.nextElementSibling;
          if (value) {
            value.textContent = decision;
          }
        }
      });
    } catch (error) {
      console.error('Could not save invoice decision.', error);
      if (invoiceUpdated) {
        notice.textContent =
          `ההחלטה נשמרה בחשבונית, אך עדכון רשימת UploadedInvoices נכשל. ${error instanceof Error ? error.message : ''} פנו למנהל המערכת כדי לסנכרן את הרשומה.`;
        this._updateDisplayedInvoiceStatus(decision);
        this._element('decision-options').hidden = true;
        this._element('comments-label').hidden = true;
        this._element<HTMLTextAreaElement>('comments').hidden = true;
        this._element('decision-actions').hidden = true;
      } else {
        notice.textContent = error instanceof Error
          ? error.message
          : 'שליחת ההחלטה נכשלה.';
        button.disabled = false;
        button.textContent = 'שליחת החלטה';
        this._submitting = false;
      }
    }
  }

  private _updateDisplayedInvoiceStatus(decision: string): void {
    this._element('invoice-details').querySelectorAll('dt').forEach((term) => {
      if (term.textContent === 'סטטוס נוכחי') {
        const value = term.nextElementSibling;
        if (value) {
          value.textContent = decision;
        }
      }
    });
  }

  private async _getUploadedInvoiceItemId(invoiceItemId: number): Promise<number> {
    const url = new URL(
      this._restUrl("lists/getbytitle('UploadedInvoices')/items")
    );
    url.searchParams.set('$select', 'Id,InvoiceItemId');
    url.searchParams.set('$filter', `InvoiceItemId eq ${invoiceItemId}`);
    const payload = await this._getJson<IUploadedInvoiceCollection>(url.toString());
    const items = payload.results || payload.value;
    if (!Array.isArray(items) || items.length !== 1) {
      throw new Error(
        'לא נמצא פריט מעקב יחיד עבור החשבונית ברשימת UploadedInvoices.'
      );
    }

    const item = items[0];
    if (
      !Number.isSafeInteger(item.Id) ||
      item.Id <= 0 ||
      item.InvoiceItemId !== invoiceItemId
    ) {
      throw new Error('פריט המעקב של החשבונית אינו תקין.');
    }

    return item.Id;
  }

  private _restUrl(path: string): string {
    return `${this.context.pageContext.web.absoluteUrl}/_api/web/${path}`;
  }

  private async _getJson<T>(url: string): Promise<T> {
    console.log('Fetching JSON from URL:', url);
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
    console.log('Received response:', response);
    if (!response.ok) {
      throw new Error(await this._responseError(response));
    }
    // console.log('Response body:', await response.json());

    const body = await response.json() as T | IODataVerboseResponse<T>;
    return typeof body === 'object' && body !== null && 'd' in body
      ? body.d
      : body;
  }

  private async _responseError(response: SPHttpClientResponse): Promise<string> {
    if (response.status === 403) {
      return 'אין לך הרשאה לבצע פעולה זו.';
    }
    if (response.status === 412) {
      return 'החשבונית עודכנה על ידי משתמש אחר. טענו מחדש את הדף לפני שתנסו שוב.';
    }
    return `בקשת SharePoint נכשלה (${response.status}).`;
  }

  private _showMessage(message: string): void {
    this._element('approval-loading').hidden = true;
    const element = this._element<HTMLDivElement>('approval-message');
    element.textContent = message;
    element.hidden = false;
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
}
