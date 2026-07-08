import * as React from 'react';
import SignatureCanvas from 'react-signature-canvas';
import * as pdfJsLib from 'pdfjs-dist/legacy/build/pdf.mjs';
import 'pdfjs-dist/legacy/build/pdf.worker.mjs';
import { PDFDocument } from 'pdf-lib';
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
  renderedWidth: number;
  renderedHeight: number;
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
  isRejectDialogOpen: boolean;
  isApproving: boolean;
  isRejecting: boolean;
  activeDocument?: IDocumentListItem;
  rejectionComments: string;
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
      isDialogOpen: false,
      isRejectDialogOpen: false,
      isApproving: false,
      isRejecting: false,
      rejectionComments: ''
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
    const canApprove: boolean = !!this.state.activeDocument && !!placement && signatureDataUrl.length > 0 && !isLoading && !isRendering && !this.state.isApproving;
    const canReject: boolean = !!this.state.activeDocument && !this.state.isRejecting && !this.state.isApproving;
    const canSubmitReject: boolean = this.state.rejectionComments.trim().length > 0 && !this.state.isRejecting;
    const approveButtonTitle: string = canApprove
      ? 'Save the signed PDF and approve the current document'
      : 'Place the signature on the PDF before approving';
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

                <div className={styles.status} role="status" aria-live="polite">
                  {this.state.statusMessage}
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
                    <div className={styles.emptyState}>{isLoading || isRendering ? 'Working...' : 'No document found'}</div>
                  )}
                </div>
              </main>
            </div>
          </div>
          <DialogFooter>
            <PrimaryButton text={this.state.isApproving ? 'Approving...' : 'Approve'} onClick={this._approveDocument} disabled={!canApprove} title={approveButtonTitle} />
            <DefaultButton text={this.state.isRejecting ? 'Rejecting...' : 'Reject'} onClick={this._openRejectDialog} disabled={!canReject} />
          </DialogFooter>
        </Dialog>

        <Dialog
          hidden={!this.state.isRejectDialogOpen}
          onDismiss={this._closeRejectDialog}
          dialogContentProps={{
            title: 'Reject Document',
            subText: 'Enter rejection comments to submit this rejection.'
          }}
          modalProps={{
            isBlocking: true
          }}
        >
          <label className={styles.rangeLabel} htmlFor="rejectionCommentsDialog">
            Rejection comments
          </label>
          <textarea
            id="rejectionCommentsDialog"
            value={this.state.rejectionComments}
            onChange={this._handleRejectionCommentsChange}
            rows={5}
            style={{ width: '100%', resize: 'vertical', padding: '8px', boxSizing: 'border-box' }}
            placeholder="Type the rejection comments here"
          />
          <DialogFooter>
            <PrimaryButton
              text={this.state.isRejecting ? 'Rejecting...' : 'Submit Rejection'}
              onClick={this._rejectDocument}
              disabled={!canSubmitReject}
            />
            <DefaultButton text="Cancel" onClick={this._closeRejectDialog} disabled={this.state.isRejecting} />
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
        width: this.state.signatureWidth,
        renderedWidth: canvas.width,
        renderedHeight: canvas.height
      },
      statusMessage: 'Signature placed successfully.'
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

    this._blurActiveElement();

    this.setState({
      activeDocument: item,
      isDialogOpen: true,
      isRejectDialogOpen: false,
      fileName: '',
      pdfBytes: undefined,
      pdfDocument: undefined,
      pageCount: 0,
      currentPageIndex: 0,
      pageSize: undefined,
      placement: undefined,
      signatureDataUrl: '',
      rejectionComments: '',
      isLoading: true,
      statusMessage: 'Loading PDF for selected document...'
    });

    await this._loadDocumentFromSharePoint(item.id, item.attachmentServerRelativeUrl);
  };

  private readonly _closeDialog = (): void => {
    this._blurActiveElement();

    this.setState({
      isDialogOpen: false,
      isRejectDialogOpen: false,
      activeDocument: undefined,
      rejectionComments: '',
      statusMessage: 'Document viewer closed.'
    });
  };

  private readonly _openRejectDialog = (): void => {
    if (!this.state.activeDocument) {
      this.setState({
        statusMessage: 'Select a document before rejecting it.'
      });
      return;
    }

    this.setState({
      isRejectDialogOpen: true,
      rejectionComments: ''
    });
  };

  private readonly _closeRejectDialog = (): void => {
    this.setState({
      isRejectDialogOpen: false,
      rejectionComments: ''
    });
  };

  private readonly _approveDocument = async (): Promise<void> => {
    const activeDocument = this.state.activeDocument;

    if (!activeDocument) {
      this.setState({
        statusMessage: 'Select a document before approving it.'
      });
      return;
    }

    const approvalUpdate = this._getApprovalUpdate(activeDocument.approvalStatus);

    if (!approvalUpdate) {
      this.setState({
        statusMessage: `The current approval status '${activeDocument.approvalStatus || ''}' cannot be approved.`
      });
      return;
    }

    const { siteUrl, taskListName, webAbsoluteUrl } = this.props;
    const trimmedSiteUrl: string = siteUrl ? siteUrl.trim() : webAbsoluteUrl;
    const listName: string | undefined = taskListName ? taskListName.trim() : undefined;

    if (!trimmedSiteUrl) {
      this.setState({
        statusMessage: 'The site URL is not configured.'
      });
      return;
    }

    if (!listName) {
      this.setState({
        statusMessage: 'The SharePoint list name is not configured.'
      });
      return;
    }

    this.setState({
      isApproving: true,
      statusMessage: 'Approving document...'
    });

    const apiBaseUrl: string = trimmedSiteUrl.replace(/\/$/, '');
    const listTitle: string = listName.replace(/'/g, "''");

    try {
      this.setState({
        statusMessage: 'Creating signed PDF...'
      });

      const signedPdfBytes = await this._createSignedPdfBytes(activeDocument);

      this.setState({
        statusMessage: 'Replacing the current attachment with the signed PDF...'
      });

      await this._replaceCurrentAttachmentWithPdf(apiBaseUrl, listTitle, activeDocument, signedPdfBytes);

      this.setState({
        statusMessage: 'Updating approval status...'
      });

      await this._updateSharePointListItem(apiBaseUrl, listTitle, activeDocument.id, {
        ApprovalStatus: approvalUpdate.approvalStatus,
        Comments: approvalUpdate.comments,
        IsTaskActive: false
      });

      if (activeDocument.documentId && activeDocument.documentNumber) {
        const relatedItems = await this._getRelatedDocumentItems(
          apiBaseUrl,
          listTitle,
          activeDocument.documentId,
          activeDocument.documentNumber,
          activeDocument.id
        );

        for (let i = 0; i < relatedItems.length; i++) {
          await this._updateSharePointListItem(apiBaseUrl, listTitle, relatedItems[i], {
            ApprovalStatus: '',
            IsTaskActive: false
          });
        }
      }

      await this._loadDocumentListFromSharePoint();

      this.setState({
        isDialogOpen: false,
        isRejectDialogOpen: false,
        activeDocument: undefined,
        isApproving: false,
        rejectionComments: '',
        statusMessage: `Document approved as ${approvalUpdate.approvalStatus}.`
      });
    } catch (error) {
      this.setState({
        isApproving: false,
        statusMessage: `The document could not be approved. ${this._getErrorMessage(error)}`
      });
    }
  };

  private readonly _rejectDocument = async (): Promise<void> => {
    const activeDocument = this.state.activeDocument;
    const rejectionComments = this.state.rejectionComments.trim();

    if (!activeDocument) {
      this.setState({
        statusMessage: 'Select a document before rejecting it.'
      });
      return;
    }

    if (!rejectionComments) {
      this.setState({
        statusMessage: 'Enter rejection comments before submitting the rejection.'
      });
      return;
    }

    const rejectionUpdate = this._getRejectionUpdate(activeDocument.approvalStatus);

    if (!rejectionUpdate) {
      this.setState({
        statusMessage: `The current approval status '${activeDocument.approvalStatus || ''}' cannot be rejected.`
      });
      return;
    }

    const { siteUrl, taskListName, webAbsoluteUrl } = this.props;
    const trimmedSiteUrl: string = siteUrl ? siteUrl.trim() : webAbsoluteUrl;
    const listName: string | undefined = taskListName ? taskListName.trim() : undefined;

    if (!trimmedSiteUrl) {
      this.setState({
        statusMessage: 'The site URL is not configured.'
      });
      return;
    }

    if (!listName) {
      this.setState({
        statusMessage: 'The SharePoint list name is not configured.'
      });
      return;
    }

    this.setState({
        isRejectDialogOpen: false,
      isRejecting: true,
      statusMessage: 'Submitting rejection...'
    });

    const apiBaseUrl: string = trimmedSiteUrl.replace(/\/$/, '');
    const listTitle: string = listName.replace(/'/g, "''");

    try {
      await this._updateSharePointListItem(apiBaseUrl, listTitle, activeDocument.id, {
        ApprovalStatus: rejectionUpdate.approvalStatus,
        Comments: rejectionComments,
        IsTaskActive: false
      });

      if (activeDocument.documentId && activeDocument.documentNumber) {
        const relatedItems = await this._getRelatedDocumentItems(
          apiBaseUrl,
          listTitle,
          activeDocument.documentId,
          activeDocument.documentNumber,
          activeDocument.id
        );

        for (let i = 0; i < relatedItems.length; i++) {
          await this._updateSharePointListItem(apiBaseUrl, listTitle, relatedItems[i], {
            ApprovalStatus: '',
            IsTaskActive: false
          });
        }
      }

      await this._loadDocumentListFromSharePoint();

      this.setState({
        isDialogOpen: false,
        isRejectDialogOpen: false,
        activeDocument: undefined,
        isRejecting: false,
        rejectionComments: '',
        statusMessage: `Document rejected as ${rejectionUpdate.approvalStatus}.`
      });
    } catch (error) {
      this.setState({
        isRejectDialogOpen: false,
        isRejecting: false,
        statusMessage: `The document could not be rejected. ${this._getErrorMessage(error)}`
      });
    }
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
      statusMessage: 'Loading tasks...',
      documents: []
    });

    const apiBaseUrl: string = trimmedSiteUrl.replace(/\/$/, '');
    const listInfo = this._getListApiPath(listName, trimmedSiteUrl);
    const currentUserEmail = this.props.userEmail?.trim() || '';
    const signatoryFilter = currentUserEmail
      ? `&$filter=L1Signatory eq '${currentUserEmail.replace(/'/g, "''")}'`
      : '';

    try {
      const requestUrl = `${apiBaseUrl}/_api/${listInfo.apiPath}/items?$select=Id,Title,DocumentID,DocumentName,DocumentNumber,Trader,AccountCode,ApprovalStatus,Comments,IsTaskActive,Created,Modified,L1Signatory,AttachmentFiles&$expand=AttachmentFiles&$top=500${signatoryFilter}`;
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
        DocumentID?: string;
        DocumentName?: string;
        DocumentNumber?: string;
        Trader?: string;
        AccountCode?: string;
        ApprovalStatus?: string;
        Comments?: string;
        IsTaskActive?: boolean;
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
            documentId: item.DocumentID,
            documentName: item.DocumentName,
            documentNumber: item.DocumentNumber,
            trader: item.Trader,
            accountCode: item.AccountCode,
            approvalStatus: item.ApprovalStatus,
            comments: item.Comments,
            isTaskActive: item.IsTaskActive,
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
        `${apiBaseUrl}/_api/web/GetFileByServerRelativeUrl('${this._normalizeServerRelativeUrlForApi(pdfAttachment.ServerRelativeUrl)}')/$value`,
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

  private _blurActiveElement(): void {
    const activeElement = document.activeElement;

    if (activeElement instanceof HTMLElement) {
      activeElement.blur();
    }
  }

  private _getApprovalUpdate(currentStatus?: string): { approvalStatus: string; comments: string } | undefined {
    const normalizedStatus = (currentStatus || '').trim();

    if (normalizedStatus === 'L1 Pending For Signature') {
      return {
        approvalStatus: 'L1 Signed',
        comments: 'Approved by Level 1'
      };
    }

    if (normalizedStatus === 'L2 Pending For Signature') {
      return {
        approvalStatus: 'L2 Signed',
        comments: 'Approved by L2'
      };
    }

    return undefined;
  }

  private _getRejectionUpdate(currentStatus?: string): { approvalStatus: string } | undefined {
    const normalizedStatus = (currentStatus || '').trim();

    if (normalizedStatus === 'L1 Pending For Signature') {
      return {
        approvalStatus: 'L1 Rejected'
      };
    }

    if (normalizedStatus === 'L2 Pending For Signature') {
      return {
        approvalStatus: 'L2 Rejected'
      };
    }

    return undefined;
  }

  private readonly _createSignedPdfBytes = async (activeDocument: IDocumentListItem): Promise<Uint8Array> => {
    const { pdfBytes, placement, signatureDataUrl, signatureAspectRatio } = this.state;

    if (!placement || !signatureDataUrl) {
      throw new Error('A placed signature is required before approving the document.');
    }

    let sourcePdfBytes: Uint8Array | undefined = this._isPdfByteArray(pdfBytes) ? pdfBytes : undefined;

    if (!sourcePdfBytes) {
      sourcePdfBytes = await this._loadPdfBytesForApproval(activeDocument);

      if (!sourcePdfBytes) {
        throw new Error('Unable to load the source PDF for approval.');
      }

      if (!this._isPdfByteArray(sourcePdfBytes)) {
        throw new Error('Unable to load a valid PDF for approval.');
      }
    }

    try {
      return await this._stampSignatureOnPdf(sourcePdfBytes, placement, signatureDataUrl, signatureAspectRatio);
    } catch (error) {
      throw new Error(`Unable to create the signed PDF. ${this._getErrorMessage(error)}`);
    }
  };

  private readonly _stampSignatureOnPdf = async (
    sourcePdfBytes: Uint8Array,
    placement: ISignaturePlacement,
    signatureDataUrl: string,
    signatureAspectRatio: number
  ): Promise<Uint8Array> => {
    const pdfDocument = await PDFDocument.load(sourcePdfBytes);
    const page = pdfDocument.getPages()[placement.pageIndex];

    if (!page) {
      throw new Error('The selected PDF page could not be found.');
    }

    const renderedWidth: number = placement.renderedWidth || this.state.pageSize?.width || 0;
    const renderedHeight: number = placement.renderedHeight || this.state.pageSize?.height || 0;

    if (renderedWidth <= 0 || renderedHeight <= 0) {
      throw new Error('The PDF preview size is not available for signature placement.');
    }

    const pngBytes = this._dataUrlToUint8Array(signatureDataUrl);
    const signatureImage = await pdfDocument.embedPng(pngBytes);
    const pageWidth = page.getWidth();
    const pageHeight = page.getHeight();
    const signatureHeight = this._getSignatureHeight(placement.width, signatureAspectRatio);
    const scaleX = pageWidth / renderedWidth;
    const scaleY = pageHeight / renderedHeight;
    const x = placement.x * scaleX;
    const y = pageHeight - ((placement.y + signatureHeight) * scaleY);
    const width = placement.width * scaleX;
    const height = signatureHeight * scaleY;

    page.drawImage(signatureImage, {
      x,
      y,
      width,
      height
    });

    return await pdfDocument.save();
  };

  private readonly _loadPdfBytesForApproval = async (activeDocument: IDocumentListItem): Promise<Uint8Array | undefined> => {
    if (this.state.pdfBytes && this.state.pdfBytes.length > 0) {
      return this.state.pdfBytes;
    }

    const { siteUrl, spHttpClient, webAbsoluteUrl } = this.props;
    const trimmedSiteUrl: string = siteUrl ? siteUrl.trim() : webAbsoluteUrl;

    if (!trimmedSiteUrl || !activeDocument.attachmentServerRelativeUrl) {
      return undefined;
    }

    const apiBaseUrl: string = trimmedSiteUrl.replace(/\/$/, '');
    const normalizedAttachmentUrl = this._normalizeServerRelativeUrlForApi(activeDocument.attachmentServerRelativeUrl);

    const fileResponse = await spHttpClient.get(
      `${apiBaseUrl}/_api/web/GetFileByServerRelativeUrl('${normalizedAttachmentUrl}')/$value`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/pdf'
        }
      }
    );

    if (!fileResponse.ok) {
      throw new Error(`Unable to reload the source PDF. Status ${fileResponse.status}`);
    }

    return new Uint8Array(await fileResponse.arrayBuffer());
  };

  private readonly _replaceCurrentAttachmentWithPdf = async (
    apiBaseUrl: string,
    listTitle: string,
    activeDocument: IDocumentListItem,
    signedPdfBytes: Uint8Array
  ): Promise<void> => {
    const currentAttachmentName = activeDocument.attachmentFileName || this._getFileNameFromServerRelativeUrl(activeDocument.attachmentServerRelativeUrl) || `${activeDocument.documentName || activeDocument.title || 'document'}.pdf`;
    const encodedAttachmentName = encodeURIComponent(currentAttachmentName);

    await this._deleteAttachmentFile(apiBaseUrl, listTitle, activeDocument.id, encodedAttachmentName);
    await this._addAttachmentFile(apiBaseUrl, listTitle, activeDocument.id, currentAttachmentName, signedPdfBytes);
  };

  private readonly _deleteAttachmentFile = async (
    apiBaseUrl: string,
    listTitle: string,
    itemId: number,
    encodedAttachmentName: string
  ): Promise<void> => {
    const response = await this.props.spHttpClient.post(
      `${apiBaseUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})/AttachmentFiles('${encodedAttachmentName}')`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE'
        }
      }
    );

    if (!response.ok) {
      throw new Error(`Unable to delete the existing attachment. Status ${response.status}`);
    }
  };

  private readonly _addAttachmentFile = async (
    apiBaseUrl: string,
    listTitle: string,
    itemId: number,
    fileName: string,
    fileBytes: Uint8Array
  ): Promise<void> => {
    const encodedFileName = encodeURIComponent(fileName);
    const response = await this.props.spHttpClient.post(
      `${apiBaseUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})/AttachmentFiles/add(FileName='${encodedFileName}')`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/pdf'
        },
        body: new Blob([fileBytes], { type: 'application/pdf' })
      }
    );

    if (!response.ok) {
      throw new Error(`Unable to upload the signed attachment. Status ${response.status}`);
    }
  };

  private _dataUrlToUint8Array(dataUrl: string): Uint8Array {
    const base64 = dataUrl.split(',')[1] || '';
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);

    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }

    return bytes;
  }

  private _isPdfByteArray(bytes?: Uint8Array): boolean {
    return !!bytes &&
      bytes.length >= 4 &&
      bytes[0] === 0x25 &&
      bytes[1] === 0x50 &&
      bytes[2] === 0x44 &&
      bytes[3] === 0x46;
  }

  private _getFileNameFromServerRelativeUrl(serverRelativeUrl?: string): string | undefined {
    if (!serverRelativeUrl) {
      return undefined;
    }

    return serverRelativeUrl.substring(serverRelativeUrl.lastIndexOf('/') + 1) || undefined;
  }

  private _normalizeServerRelativeUrlForApi(serverRelativeUrl: string): string {
    const normalizedSegments = serverRelativeUrl
      .trim()
      .split('/')
      .map((segment: string, index: number) => {
        if (index === 0 && segment === '') {
          return '';
        }

        let decodedSegment = segment;

        try {
          decodedSegment = decodeURIComponent(segment);
        } catch {
          // Keep the original segment if it is not valid URI encoding.
        }

        return encodeURIComponent(decodedSegment);
      })
      .join('/');

    return normalizedSegments.replace(/'/g, "''");
  }

  private readonly _updateSharePointListItem = async (
    apiBaseUrl: string,
    listTitle: string,
    itemId: number,
    updates: Record<string, unknown>
  ): Promise<void> => {
    const response = await this.props.spHttpClient.post(
      `${apiBaseUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify(updates)
      }
    );

    if (!response.ok) {
      throw new Error(`Unable to update list item ${itemId}. Status ${response.status}`);
    }
  };

  private readonly _handleRejectionCommentsChange = (
    event: React.ChangeEvent<HTMLTextAreaElement>
  ): void => {
    this.setState({
      rejectionComments: event.target.value
    });
  };

  private readonly _getRelatedDocumentItems = async (
    apiBaseUrl: string,
    listTitle: string,
    documentId: string,
    documentNumber: string,
    currentItemId: number
  ): Promise<number[]> => {
    const safeDocumentId = documentId.replace(/'/g, "''");
    const safeDocumentNumber = documentNumber.replace(/'/g, "''");
    const response = await this.props.spHttpClient.get(
      `${apiBaseUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=Id&$filter=DocumentID eq '${safeDocumentId}' and DocumentNumber eq '${safeDocumentNumber}' and Id ne ${currentItemId}&$top=500`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );

    if (!response.ok) {
      throw new Error(`Unable to query related list items. Status ${response.status}`);
    }

    const json = await response.json();
    const items = (json && (json.value as Array<{ Id: number }>)) || [];
    return items.map((item: { Id: number }) => item.Id);
  };

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

}
