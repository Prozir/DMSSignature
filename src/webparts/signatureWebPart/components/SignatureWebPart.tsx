import * as React from 'react';
import SignatureCanvas from 'react-signature-canvas';
import { PDFDocument } from 'pdf-lib';
import * as pdfJsLib from 'pdfjs-dist/legacy/build/pdf.mjs';
import 'pdfjs-dist/legacy/build/pdf.worker.mjs';
import { SPHttpClient } from '@microsoft/sp-http';
import { Dialog, DialogFooter, DefaultButton, PrimaryButton } from '@fluentui/react';
import styles from './SignatureWebPart.module.scss';
import DocumentsList, { IDocumentListItem } from './DocumentsList';
import type { ISignatureWebPartProps } from './ISignatureWebPartProps';

interface ISignaturePlacement {
  pageIndex: number;
  x: number;
  y: number;
  width: number;
}

interface IPdfPageSize {
  width: number;
  height: number;
}

interface ISignatureWebPartState {
  fileName: string;
  pdfBytes?: Uint8Array;
  pdfDocument?: unknown;
  pageCount: number;
  currentPageIndex: number;
  pageSize?: IPdfPageSize;
  placement?: ISignaturePlacement;
  signatureDataUrl: string;
  signatureAspectRatio: number;
  signatureWidth: number;
  isLoading: boolean;
  isRendering: boolean;
  statusMessage: string;
  documents: IDocumentListItem[];
  isDocumentsLoading: boolean;
  isDialogOpen: boolean;
  activeDocument?: IDocumentListItem;
}

export default class SignatureWebPart extends React.Component<ISignatureWebPartProps, ISignatureWebPartState> {
  private readonly _canvasWrapRef: React.RefObject<HTMLDivElement> = React.createRef<HTMLDivElement>();
  private readonly _previewCanvasRef: React.RefObject<HTMLCanvasElement> = React.createRef<HTMLCanvasElement>();
  private readonly _signatureRef: React.RefObject<SignatureCanvas> = React.createRef<SignatureCanvas>();
  private _resizeTimer: number | undefined;

  public constructor(props: ISignatureWebPartProps) {
    super(props);

    this.state = {
      fileName: '',
      pageCount: 0,
      currentPageIndex: 0,
      signatureDataUrl: '',
      signatureAspectRatio: 2.8,
      signatureWidth: 110,
      isLoading: false,
      isRendering: false,
      statusMessage: 'Loading documents from SharePoint... waiting for selection.',
      documents: [],
      isDocumentsLoading: false,
      isDialogOpen: false
    };
  }

  public componentDidUpdate(
    previousProps: ISignatureWebPartProps,
    previousState: ISignatureWebPartState
  ): void {
    if (
      previousProps.siteUrl !== this.props.siteUrl ||
      previousProps.taskListName !== this.props.taskListName
    ) {
      this._loadDocumentListFromSharePoint().catch(() => {
        this.setState({
          isDocumentsLoading: false,
          statusMessage: 'The SharePoint documents could not be loaded.'
        });
      });
    }

    if (
      previousState.pdfDocument !== this.state.pdfDocument ||
      previousState.currentPageIndex !== this.state.currentPageIndex
    ) {
      this._renderCurrentPage().catch(() => {
        this.setState({
          isRendering: false,
          statusMessage: 'The selected PDF page could not be rendered.'
        });
      });
    }
  }

  public componentDidMount(): void {
    window.addEventListener('resize', this._handleWindowResize);
    this._loadDocumentListFromSharePoint().catch(() => {
      this.setState({
        isDocumentsLoading: false,
        statusMessage: 'The SharePoint documents could not be loaded.'
      });
    });
  }

  public componentWillUnmount(): void {
    window.removeEventListener('resize', this._handleWindowResize);

    if (this._resizeTimer !== undefined) {
      window.clearTimeout(this._resizeTimer);
    }
  }

  public render(): React.ReactElement<ISignatureWebPartProps> {
    const {
      currentPageIndex,
      fileName,
      isLoading,
      isRendering,
      pageCount,
      pageSize,
      placement,
      signatureAspectRatio,
      signatureDataUrl,
      signatureWidth,
    } = this.state;

    const canPlaceSignature: boolean = !!pageSize && signatureDataUrl.length > 0 && !isRendering;
    const canDownload: boolean = !!this.state.pdfBytes && !!placement && signatureDataUrl.length > 0 && !isLoading;
    const signatureHeight: number = this._getSignatureHeight(signatureWidth, signatureAspectRatio);

    return (
      <section className={`${styles.signatureWebPart} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.listContainer}>
          <DocumentsList
            items={this.state.documents}
            isLoading={this.state.isDocumentsLoading}
            onOpenItem={this._openDocument}
          />
        </div>

        <Dialog
          hidden={!this.state.isDialogOpen}
          onDismiss={this._closeDialog}
          dialogContentProps={{
            title: this.state.activeDocument ? `Sign ${this.state.activeDocument.documentName || this.state.activeDocument.title || 'document'}` : 'Sign document'
          }}
          modalProps={{ 
            isBlocking: true,
            className: styles.dialogModal
          }}
           minWidth="70vw"
           maxWidth="1100px"
        >
          <div className={styles.dialogContent}>
            <div className={styles.workspace}>
              <aside className={styles.controls}>
                <div className={styles.panel}>
                  <h3>Signature</h3>
                  <SignatureCanvas
                    ref={this._signatureRef}
                    penColor="#111827"
                    clearOnResize={false}
                    canvasProps={{
                      className: styles.signatureCanvas,
                      'aria-label': 'Draw signature'
                    }}
                    onEnd={this._captureSignature}
                  />
                  <div className={styles.buttonRow}>
                    <button type="button" onClick={this._captureSignature}>Use signature</button>
                    <button type="button" onClick={this._clearSignature}>Clear</button>
                  </div>
                </div>

                <div className={styles.panel}>
                  <h3>Placement</h3>
                  <label className={styles.rangeLabel} htmlFor="signatureWidth">
                    Width: {signatureWidth}px
                  </label>
                  <input
                    id="signatureWidth"
                    type="range"
                    min="20"
                    max="220"
                    step="10"
                    value={signatureWidth}
                    onChange={this._handleSignatureWidthChange}
                  />
                  <p className={styles.helpText}>
                    {canPlaceSignature ? 'Click the PDF preview where the top-left of the signature should appear.' : 'Draw a signature.'}
                  </p>
                </div>

                <button
                  type="button"
                  className={styles.downloadButton}
                  disabled={!canDownload}
                  onClick={this._downloadSignedPdf}
                >
                  Download signed PDF
                </button>
              </aside>

              <main className={styles.previewArea}>
                <div className={styles.previewToolbar}>
                  <span>{fileName || 'No PDF selected'}</span>
                  <div className={styles.pageControls}>
                    <button type="button" disabled={currentPageIndex === 0 || pageCount === 0} onClick={this._goToPreviousPage}>
                      Previous
                    </button>
                    <span>{pageCount > 0 ? `${currentPageIndex + 1} / ${pageCount}` : '0 / 0'}</span>
                    <button type="button" disabled={currentPageIndex >= pageCount - 1 || pageCount === 0} onClick={this._goToNextPage}>
                      Next
                    </button>
                  </div>
                </div>

                <div className={styles.canvasWrap} ref={this._canvasWrapRef}>
                  <div className={styles.pageSurface}>
                    <canvas
                      ref={this._previewCanvasRef}
                      className={styles.pdfCanvas}
                      onClick={this._handlePreviewClick}
                    />
                    {placement && placement.pageIndex === currentPageIndex && signatureDataUrl && (
                      <img
                        alt="Signature placement"
                        className={styles.signaturePreview}
                        src={signatureDataUrl}
                        style={{
                          left: `${placement.x}px`,
                          top: `${placement.y}px`,
                          width: `${placement.width}px`,
                          height: `${signatureHeight}px`
                        }}
                      />
                    )}
                  </div>
                  {(isLoading || isRendering || pageCount === 0) && (
                    <div className={styles.emptyState}>{isLoading || isRendering ? 'Working...' : 'No document available from SharePoint.'}</div>
                  )}
                </div>
              </main>
            </div>
          </div>
          <DialogFooter>
            <PrimaryButton text="Close" onClick={this._closeDialog} />
            <DefaultButton text="Cancel" onClick={this._closeDialog} />
          </DialogFooter>
        </Dialog>
      </section>
    );
  }

  private readonly _captureSignature = (): void => {
    const signaturePad: SignatureCanvas | null = this._signatureRef.current;

    if (!signaturePad || signaturePad.isEmpty()) {
      this.setState({
        signatureDataUrl: '',
        signatureAspectRatio: 2.8,
        placement: undefined,
        statusMessage: 'Draw a signature before placing it on the PDF.'
      });
      return;
    }

    const signatureCanvas: HTMLCanvasElement | undefined = this._getTrimmedSignatureCanvas(signaturePad);

    if (!signatureCanvas) {
      this.setState({
        signatureDataUrl: '',
        signatureAspectRatio: 2.8,
        placement: undefined,
        statusMessage: 'The signature could not be captured. Please draw it again.'
      });
      return;
    }

    this.setState({
      signatureDataUrl: signatureCanvas.toDataURL('image/png'),
      signatureAspectRatio: signatureCanvas.width / signatureCanvas.height,
      placement: undefined,
      statusMessage: 'Signature captured. Click the PDF page to place it.'
    });
  };

  private readonly _clearSignature = (): void => {
    const signaturePad: SignatureCanvas | null = this._signatureRef.current;

    if (signaturePad) {
      signaturePad.clear();
    }

    this.setState({
      signatureDataUrl: '',
      signatureAspectRatio: 2.8,
      placement: undefined,
      statusMessage: 'Signature cleared.'
    });
  };

  private readonly _handleSignatureWidthChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const width: number = Number(event.target.value);
    const placement: ISignaturePlacement | undefined = this.state.placement
      ? {
        ...this.state.placement,
        width
      }
      : undefined;

    this.setState({
      signatureWidth: width,
      placement
    });
  };

  private readonly _handlePreviewClick = (event: React.MouseEvent<HTMLCanvasElement>): void => {
    const canvas: HTMLCanvasElement | null = this._previewCanvasRef.current;

    if (!canvas || !this.state.signatureDataUrl) {
      this.setState({ statusMessage: 'Draw and use a signature before placing it.' });
      return;
    }

    const rect: DOMRect = canvas.getBoundingClientRect();
    const scaleX: number = canvas.width / rect.width;
    const scaleY: number = canvas.height / rect.height;
    const x: number = Math.max(0, (event.clientX - rect.left) * scaleX);
    const y: number = Math.max(0, (event.clientY - rect.top) * scaleY);

    this.setState({
      placement: {
        pageIndex: this.state.currentPageIndex,
        x,
        y,
        width: this.state.signatureWidth
      },
      statusMessage: 'Signature placed. Download the signed PDF when ready.'
    });
  };

  private readonly _goToPreviousPage = (): void => {
    this.setState(previousState => ({
      currentPageIndex: Math.max(0, previousState.currentPageIndex - 1)
    }));
  };

  private readonly _goToNextPage = (): void => {
    this.setState(previousState => ({
      currentPageIndex: Math.min(previousState.pageCount - 1, previousState.currentPageIndex + 1)
    }));
  };

  private readonly _handleWindowResize = (): void => {
    if (!this.state.pdfDocument) {
      return;
    }

    if (this._resizeTimer !== undefined) {
      window.clearTimeout(this._resizeTimer);
    }

    this._resizeTimer = window.setTimeout(() => {
      this._renderCurrentPage().catch(() => {
        this.setState({
          isRendering: false,
          statusMessage: 'The selected PDF page could not be resized.'
        });
      });
    }, 150);
  };

  private readonly _openDocument = async (item: IDocumentListItem): Promise<void> => {
    if (!item.attachmentServerRelativeUrl) {
      this.setState({
        statusMessage: 'The selected document does not contain a PDF attachment.'
      });
      return;
    }

    this.setState({
      activeDocument: item,
      isDialogOpen: true,
      fileName: '',
      pdfBytes: undefined,
      pdfDocument: undefined,
      pageCount: 0,
      currentPageIndex: 0,
      pageSize: undefined,
      placement: undefined,
      signatureDataUrl: '',
      isLoading: true,
      statusMessage: 'Loading PDF for selected document...'
    });

    await this._loadDocumentFromSharePoint(item.id, item.attachmentServerRelativeUrl);
  };

  private readonly _closeDialog = (): void => {
    this.setState({
      isDialogOpen: false,
      activeDocument: undefined,
      statusMessage: 'Document viewer closed.'
    });
  };

  private _getListApiPath(listName: string, webAbsoluteUrl: string): { apiPath: string; displayName: string } {
    const trimmed = listName.trim();
    const normalizeServerRelativeUrl = (url: string): string => url.replace(/ /g, '%20');

    if (/^https?:\/\//i.test(trimmed)) {
      try {
        const listUrl = new URL(trimmed);
        const serverRelativeUrl = listUrl.pathname;
        return {
          apiPath: `web/GetList('${normalizeServerRelativeUrl(serverRelativeUrl).replace(/'/g, "''")}')`,
          displayName: serverRelativeUrl.split('/').pop() || trimmed
        };
      } catch {
        // fallback to title-based access
      }
    }

    if (trimmed.indexOf('/') === 0) {
      const serverRelativeUrl = normalizeServerRelativeUrl(trimmed);
      return {
        apiPath: `web/GetList('${serverRelativeUrl.replace(/'/g, "''")}')`,
        displayName: serverRelativeUrl.split('/').pop() || trimmed
      };
    }

    return {
      apiPath: `web/lists/getbytitle('${trimmed.replace(/'/g, "''")}')`,
      displayName: trimmed
    };
  }

  private readonly _loadDocumentListFromSharePoint = async (): Promise<void> => {
    const { siteUrl, taskListName, spHttpClient, webAbsoluteUrl } = this.props;
    const trimmedSiteUrl: string = siteUrl ? siteUrl.trim() : webAbsoluteUrl;
    const listName: string | undefined = taskListName ? taskListName.trim() : undefined;

    if (!trimmedSiteUrl) {
      this.setState({
        statusMessage: 'The site URL is not configured.',
        isDocumentsLoading: false
      });
      return;
    }

    if (!listName) {
      this.setState({
        statusMessage: 'The SharePoint list name is not configured.',
        isDocumentsLoading: false
      });
      return;
    }

    this.setState({
      isDocumentsLoading: true,
      statusMessage: 'Loading signed documents from SharePoint...',
      documents: []
    });

    const apiBaseUrl: string = trimmedSiteUrl.replace(/\/$/, '');
    const listInfo = this._getListApiPath(listName, trimmedSiteUrl);

    try {
      const requestUrl = `${apiBaseUrl}/_api/${listInfo.apiPath}/items?$select=Id,Title,DocumentName,DocumentNumber,Trader,AccountCode,ApprovalStatus,Created,Modified,AttachmentFiles&$expand=AttachmentFiles&$top=500`;
      const itemsResponse = await spHttpClient.get(
        requestUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        }
      );

      if (!itemsResponse.ok) {
        throw new Error(`Unable to query the list. Status ${itemsResponse.status} for ${requestUrl}`);
      }

      const itemsJson = await itemsResponse.json();

      interface ISharePointDocumentItem {
        Id: number;
        Title?: string;
        DocumentName?: string;
        DocumentNumber?: string;
        Trader?: string;
        AccountCode?: string;
        ApprovalStatus?: string;
        Created?: string;
        Modified?: string;
        AttachmentFiles?: Array<{ FileName: string; ServerRelativeUrl: string }>;
      }

      const items = (itemsJson && (itemsJson.value as Array<ISharePointDocumentItem>)) || [];

      const documents: IDocumentListItem[] = items
        .map((item: ISharePointDocumentItem) => {
          const attachmentFiles: Array<{ FileName: string; ServerRelativeUrl: string }> =
            item.AttachmentFiles || [];
          let pdfAttachment: { FileName: string; ServerRelativeUrl: string } | undefined;

          for (let i = 0; i < attachmentFiles.length; i++) {
            const file = attachmentFiles[i];
            const fileName = file.FileName.toLowerCase();
            if (fileName.length > 4 && fileName.substr(fileName.length - 4) === '.pdf') {
              pdfAttachment = file;
              break;
            }
          }

          if (!pdfAttachment && attachmentFiles.length > 0) {
            pdfAttachment = attachmentFiles[0];
          }

          return {
            id: item.Id,
            title: item.Title,
            documentName: item.DocumentName,
            documentNumber: item.DocumentNumber,
            trader: item.Trader,
            accountCode: item.AccountCode,
            approvalStatus: item.ApprovalStatus,
            created: item.Created,
            modified: item.Modified,
            attachmentFileName: pdfAttachment ? pdfAttachment.FileName : undefined,
            attachmentServerRelativeUrl: pdfAttachment ? pdfAttachment.ServerRelativeUrl : undefined
          } as IDocumentListItem;
        })
        .filter((item: IDocumentListItem) => !!item.attachmentServerRelativeUrl);

      this.setState({
        documents,
        isDocumentsLoading: false,
        statusMessage: documents.length === 0
          ? `No signed documents with PDF attachments were found in '${listName}'. Request URL: ${requestUrl}`
          : 'Select a document to open for signature.'
      });
    } catch (error) {
      this.setState({
        documents: [],
        isDocumentsLoading: false,
        statusMessage: `The SharePoint documents could not be loaded. ${this._getErrorMessage(error)}`
      });
    }
  };

  private readonly _loadDocumentFromSharePoint = async (itemId?: number, attachmentRelativeUrl?: string): Promise<void> => {
    const { siteUrl, taskListName, spHttpClient, webAbsoluteUrl } = this.props;
    const trimmedSiteUrl: string = siteUrl ? siteUrl.trim() : webAbsoluteUrl;
    const listName: string | undefined = taskListName ? taskListName.trim() : undefined;

    if (!trimmedSiteUrl) {
      this.setState({
        statusMessage: 'The site URL is not configured.',
        isLoading: false
      });
      return;
    }

    if (!listName) {
      this.setState({
        statusMessage: 'The SharePoint list name is not configured.',
        isLoading: false
      });
      return;
    }

    this.setState({
      isLoading: true,
      statusMessage: 'Loading document from SharePoint... ',
      placement: undefined
    });

    const apiBaseUrl: string = trimmedSiteUrl.replace(/\/$/, '');
    const listTitle: string = listName.replace(/'/g, "''");

    try {
      let itemIdToUse = itemId;

      if (itemIdToUse === undefined) {
        const itemsResponse = await spHttpClient.get(
          `${apiBaseUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=Id,Title&$filter=Attachments eq true&$top=1`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=nometadata'
            }
          }
        );

        if (!itemsResponse.ok) {
          throw new Error(`Unable to query the list. Status ${itemsResponse.status}`);
        }

        const itemsJson = await itemsResponse.json();
        const items = (itemsJson && (itemsJson.value as Array<{ Id: number; Title?: string }>)) || [];

        if (items.length === 0) {
          this.setState({
            fileName: '',
            pdfBytes: undefined,
            pdfDocument: undefined,
            pageCount: 0,
            pageSize: undefined,
            placement: undefined,
            isLoading: false,
            statusMessage: `No items with attachments were found in the list '${listName}'.`
          });
          return;
        }

        itemIdToUse = items[0].Id;
      }

      let pdfAttachment: { FileName: string; ServerRelativeUrl: string } | undefined;

      if (attachmentRelativeUrl) {
        const fileName = attachmentRelativeUrl.substring(attachmentRelativeUrl.lastIndexOf('/') + 1);
        pdfAttachment = {
          FileName: fileName,
          ServerRelativeUrl: attachmentRelativeUrl
        };
      } else {
        const attachmentsResponse = await spHttpClient.get(
          `${apiBaseUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${itemIdToUse})/AttachmentFiles?$select=FileName,ServerRelativeUrl`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=nometadata'
            }
          }
        );

        if (!attachmentsResponse.ok) {
          throw new Error(`Unable to read attachments. Status ${attachmentsResponse.status}`);
        }

        const attachmentsJson = await attachmentsResponse.json();
        const attachmentFiles: Array<{ FileName: string; ServerRelativeUrl: string }> =
          (attachmentsJson && attachmentsJson.value as Array<{ FileName: string; ServerRelativeUrl: string }>) || [];

        for (let i = 0; i < attachmentFiles.length; i++) {
          const file = attachmentFiles[i];
          const fileName = file.FileName.toLowerCase();
          if (fileName.length > 4 && fileName.substr(fileName.length - 4) === '.pdf') {
            pdfAttachment = file;
            break;
          }
        }

        if (!pdfAttachment && attachmentFiles.length > 0) {
          pdfAttachment = attachmentFiles[0];
        }
      }

      if (!pdfAttachment) {
        this.setState({
          fileName: '',
          pdfBytes: undefined,
          pdfDocument: undefined,
          pageCount: 0,
          pageSize: undefined,
          placement: undefined,
          isLoading: false,
          statusMessage: `No PDF attachment was found for item ${itemIdToUse}.`
        });
        return;
      }

      const fileResponse = await spHttpClient.get(
        `${apiBaseUrl}/_api/web/GetFileByServerRelativeUrl('${pdfAttachment.ServerRelativeUrl}')/$value`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/pdf'
          }
        }
      );

      if (!fileResponse.ok) {
        throw new Error(`Unable to download file. Status ${fileResponse.status}`);
      }

      const arrayBuffer = await fileResponse.arrayBuffer();
      const bytes = new Uint8Array(arrayBuffer);
      const loadingTask = pdfJsLib.getDocument({ data: bytes, disableWorker: true });
      const pdfDocument = await loadingTask.promise;
      const pageCount = (pdfDocument as { numPages: number }).numPages;

      this.setState({
        fileName: pdfAttachment.FileName,
        pdfBytes: bytes,
        pdfDocument,
        pageCount,
        currentPageIndex: 0,
        isLoading: false,
        statusMessage: `Loaded ${pageCount} page${pageCount === 1 ? '' : 's'} from SharePoint.`
      });
    } catch (error) {
      this.setState({
        fileName: '',
        pdfBytes: undefined,
        pdfDocument: undefined,
        pageCount: 0,
        pageSize: undefined,
        placement: undefined,
        isLoading: false,
        statusMessage: `The SharePoint document could not be loaded. ${this._getErrorMessage(error)}`
      });
    }
  };

  private readonly _downloadSignedPdf = async (): Promise<void> => {
    const { fileName, pdfBytes, placement, signatureDataUrl } = this.state;
    const canvas: HTMLCanvasElement | null = this._previewCanvasRef.current;

    if (!pdfBytes || !placement || !signatureDataUrl || !canvas) {
      this.setState({ statusMessage: 'Load a SharePoint PDF, draw a signature, and place it before downloading.' });
      return;
    }

    this.setState({
      isLoading: true,
      statusMessage: 'Creating signed PDF...'
    });

    try {
      const pdfDoc: PDFDocument = await PDFDocument.load(pdfBytes);
      const signatureImage = await pdfDoc.embedPng(signatureDataUrl);
      const page = pdfDoc.getPages()[placement.pageIndex];
      const pageWidth: number = page.getWidth();
      const pageHeight: number = page.getHeight();
      const signaturePdfWidth: number = (placement.width / canvas.width) * pageWidth;
      const signaturePdfHeight: number = signaturePdfWidth / (signatureImage.width / signatureImage.height);
      const x: number = (placement.x / canvas.width) * pageWidth;
      const y: number = pageHeight - ((placement.y / canvas.height) * pageHeight) - signaturePdfHeight;

      page.drawImage(signatureImage, {
        x,
        y,
        width: signaturePdfWidth,
        height: signaturePdfHeight
      });

      const signedPdfBytes: Uint8Array = await pdfDoc.save();
      this._downloadBlob(
        signedPdfBytes,
        fileName.replace(/\.pdf$/i, '') + '-signed.pdf'
      );

      this.setState({
        isLoading: false,
        statusMessage: 'Signed PDF downloaded.'
      });
    } catch {
      this.setState({
        isLoading: false,
        statusMessage: 'The signed PDF could not be created.'
      });
    }
  };

  private async _renderCurrentPage(): Promise<void> {
    const { pdfDocument, currentPageIndex } = this.state;
    const canvasWrap: HTMLDivElement | null = this._canvasWrapRef.current;
    const canvas: HTMLCanvasElement | null = this._previewCanvasRef.current;

    if (!pdfDocument || !canvas || !canvasWrap) {
      return;
    }

    this.setState({ isRendering: true });

    const page = await (pdfDocument as { getPage: (pageNumber: number) => Promise<unknown> }).getPage(currentPageIndex + 1);
    const pdfPage = page as {
      getViewport: (options: { scale: number }) => { width: number; height: number };
      render: (options: { canvasContext: CanvasRenderingContext2D; viewport: unknown }) => { promise: Promise<void> };
    };
    const unscaledViewport = pdfPage.getViewport({ scale: 1 });
    const wrapStyle: CSSStyleDeclaration = window.getComputedStyle(canvasWrap);
    const horizontalPadding: number = parseFloat(wrapStyle.paddingLeft) + parseFloat(wrapStyle.paddingRight);
    const availableWidth: number = Math.max(280, canvasWrap.clientWidth - horizontalPadding);
    const scale: number = availableWidth / unscaledViewport.width;
    const viewport = pdfPage.getViewport({ scale });
    const context: CanvasRenderingContext2D | null = canvas.getContext('2d');

    if (!context) {
      return;
    }

    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);

    await pdfPage.render({
      canvasContext: context,
      viewport
    }).promise;

    this.setState({
      pageSize: {
        width: canvas.width,
        height: canvas.height
      },
      isRendering: false
    });
  }


  private _getErrorMessage(error: unknown): string {
    return error instanceof Error && error.message
      ? error.message
      : '';
  }

  private _getSignatureHeight(width: number, aspectRatio: number): number {
    return Math.max(32, Math.round(width / aspectRatio));
  }

  private _getTrimmedSignatureCanvas(signaturePad: SignatureCanvas): HTMLCanvasElement | undefined {
    try {
      const sourceCanvas: HTMLCanvasElement = signaturePad.getCanvas();
      const context: CanvasRenderingContext2D | null = sourceCanvas.getContext('2d');

      if (!context) {
        return undefined;
      }

      const imageData: ImageData = context.getImageData(0, 0, sourceCanvas.width, sourceCanvas.height);
      const data: Uint8ClampedArray = imageData.data;
      let top: number = sourceCanvas.height;
      let right: number = 0;
      let bottom: number = 0;
      let left: number = sourceCanvas.width;

      for (let y: number = 0; y < sourceCanvas.height; y++) {
        for (let x: number = 0; x < sourceCanvas.width; x++) {
          const alphaIndex: number = ((y * sourceCanvas.width) + x) * 4 + 3;

          if (data[alphaIndex] > 0) {
            if (x < left) {
              left = x;
            }
            if (x > right) {
              right = x;
            }
            if (y < top) {
              top = y;
            }
            if (y > bottom) {
              bottom = y;
            }
          }
        }
      }

      if (right < left || bottom < top) {
        return undefined;
      }

      const trimmedCanvas: HTMLCanvasElement = document.createElement('canvas');
      trimmedCanvas.width = right - left + 1;
      trimmedCanvas.height = bottom - top + 1;

      const trimmedContext: CanvasRenderingContext2D | null = trimmedCanvas.getContext('2d');

      if (!trimmedContext) {
        return undefined;
      }

      trimmedContext.drawImage(
        sourceCanvas,
        left,
        top,
        trimmedCanvas.width,
        trimmedCanvas.height,
        0,
        0,
        trimmedCanvas.width,
        trimmedCanvas.height
      );

      return trimmedCanvas;
    } catch {
      return undefined;
    }
  }

  private _downloadBlob(bytes: Uint8Array, fileName: string): void {
    const blob: Blob = new Blob([bytes], { type: 'application/pdf' });
    const url: string = URL.createObjectURL(blob);
    const link: HTMLAnchorElement = document.createElement('a');

    link.href = url;
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(url);
  }
}
