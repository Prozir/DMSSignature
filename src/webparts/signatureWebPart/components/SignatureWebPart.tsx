import * as React from 'react';
import SignatureCanvas from 'react-signature-canvas';
import * as pdfJsLib from 'pdfjs-dist/legacy/build/pdf.mjs';
import 'pdfjs-dist/legacy/build/pdf.worker.mjs';
import { PDFDocument } from 'pdf-lib';
import { SPHttpClient } from '@microsoft/sp-http';
import { Dialog, DialogFooter, DefaultButton, PrimaryButton } from '@fluentui/react';
import styles from './SignatureWebPart.module.scss';
import DocumentsList, { IDocumentListItem } from './DocumentsList';
import ApprovalHistory from './ApprovalHistory';
import type { ISignatureWebPartProps } from './ISignatureWebPartProps';
import type { IApprovalHistoryItem, ISignaturePlacement, ISignatureWebPartState } from './SignatureWebPartInterfaces';
import {
  getApprovalUpdate,
  getFileNameFromServerRelativeUrl,
  getListApiPath,
  getRejectionUpdate,
  normalizeServerRelativeUrlForApi
} from './SignatureWebPartSharePointHelper';
import {
  dataUrlToUint8Array,
  getSignatureHeight,
  getTrimmedSignatureCanvas,
  uint8ArrayToDataUrl
} from './SignatureWebPartSignatureHelper';

interface ISharePointDocumentItem {
  Id: number;
  Title?: string;
  DocumentID?: string;
  DocumentName?: string;
  DocumentNumber?: string;
  Trader?: string;
  AccountCode?: string;
  L1Signatory?: string;
  ApprovalStatus?: string;
  Comments?: string;
  IsTaskActive?: boolean;
  Created?: string;
  Modified?: string;
  ActionTakenOn?: string;
  AttachmentFiles?: Array<{ FileName: string; ServerRelativeUrl: string }>;
}

export default class SignatureWebPart extends React.Component<ISignatureWebPartProps, ISignatureWebPartState> {
  private static readonly _signatureMasterListName: string = 'Signature Master';
  private static readonly _taskIdQueryParamName: string = 'TaskID';
  private readonly _canvasWrapRef: React.RefObject<HTMLDivElement> = React.createRef<HTMLDivElement>();
  private readonly _previewCanvasRef: React.RefObject<HTMLCanvasElement> = React.createRef<HTMLCanvasElement>();
  private readonly _signatureRef: React.RefObject<SignatureCanvas> = React.createRef<SignatureCanvas>();
  private _isSignaturePointerDown: boolean = false;
  private _resizeTimer: number | undefined;
  private _hasHandledTaskIdDeepLink: boolean = false;
  private _isHandlingTaskIdDeepLink: boolean = false;
  private _approvalHistoryRequestId: number = 0;

  // Set default values used by the component when it first loads.
  public constructor(props: ISignatureWebPartProps) {
    super(props);

    this.state = {
      fileName: '',
      pageCount: 0,
      currentPageIndex: 0,
      signatureDataUrl: '',
      signatureAspectRatio: 2.8,
      signatureWidth: 60,
      isLoading: false,
      isRendering: false,
      statusMessage: 'Loading documents...',
      documents: [],
      isDocumentsLoading: false,
      isApprovalHistoryOpen: false,
      isApprovalHistoryLoading: false,
      approvalHistoryError: '',
      approvalHistory: [],
      isDialogOpen: false,
      isRejectDialogOpen: false,
      isApprovalSuccessDialogOpen: false,
      isTaskLinkMessageDialogOpen: false,
      taskLinkMessageDialogTitle: '',
      taskLinkMessageDialogText: '',
      isApproving: false,
      isRejecting: false,
      isSavingSignature: false,
      isLoadingSavedSignature: false,
      rejectionComments: '',
      savedSignatureDataUrl: '',
      savedSignatureAspectRatio: 2.8
    };
  }

  // Reload data or page preview when key props or state values change.
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

  // Add listeners and load the initial document list after mount.
  public componentDidMount(): void {
    window.addEventListener('resize', this._handleWindowResize);

    this._loadDocumentListFromSharePoint().catch(() => {
      this.setState({
        isDocumentsLoading: false,
        statusMessage: 'The SharePoint documents could not be loaded.'
      });
    });
  }

  // Remove listeners and timers before the component is destroyed.
  public componentWillUnmount(): void {
    window.removeEventListener('resize', this._handleWindowResize);

    if (this._resizeTimer !== undefined) {
      window.clearTimeout(this._resizeTimer);
    }
  }

  // Render the document list, dialogs, controls, and PDF preview area.
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
    const isActiveDocumentPending: boolean = (this.state.activeDocument?.approvalStatus || '').toLowerCase().indexOf('pending') !== -1;
    const canApprove: boolean = !!this.state.activeDocument && isActiveDocumentPending && !!placement && signatureDataUrl.length > 0 && !isLoading && !isRendering && !this.state.isApproving;
    const canReject: boolean = !!this.state.activeDocument && isActiveDocumentPending && !this.state.isRejecting && !this.state.isApproving;
    const canSaveSignature: boolean = !this.state.isSavingSignature && !this.state.isLoadingSavedSignature;
    const canSubmitReject: boolean = this.state.rejectionComments.trim().length > 0 && !this.state.isRejecting;
    const approveButtonTitle: string = !isActiveDocumentPending
      ? 'Available only for Pending items'
      : canApprove
        ? 'Save the signed PDF and approve the current document'
        : 'Place the signature on the PDF before approving';
    const signatureHeight: number = getSignatureHeight(signatureWidth, signatureAspectRatio);

    return (
      <section className={`${styles.signatureWebPart} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.listContainer}>
          <DocumentsList
            items={this.state.documents}
            isLoading={this.state.isDocumentsLoading}
            displayMode={this.props.displayMode}
            onOpenItem={this._openDocument}
            onShowHistory={this._showApprovalHistory}
          />
        </div>

        <ApprovalHistory
          isOpen={this.state.isApprovalHistoryOpen}
          isLoading={this.state.isApprovalHistoryLoading}
          errorMessage={this.state.approvalHistoryError}
          document={this.state.approvalHistoryDocument}
          items={this.state.approvalHistory}
          onDismiss={this._closeApprovalHistory}
        />

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
                      'aria-label': 'Draw signature',
                      onMouseDown: this._handleSignaturePointerDown,
                      onMouseUp: this._handleSignaturePointerUp,
                      onMouseLeave: this._handleSignaturePointerUp,
                      onTouchStart: this._handleSignaturePointerDown,
                      onTouchEnd: this._handleSignaturePointerUp,
                      onTouchCancel: this._handleSignaturePointerUp,
                      onMouseMoveCapture: this._handleSignatureHoverGuard
                    }}
                    onEnd={this._captureSignature}
                  />
                  <div className={styles.buttonRow}>
                    <PrimaryButton text="Use Signature" onClick={this._captureSignature} disabled={!isActiveDocumentPending} />
                    <PrimaryButton
                      text={this.state.isSavingSignature ? 'Saving signature...' : 'Save Signature'}
                      onClick={this._saveSignatureToSharePoint}
                      disabled={!canSaveSignature || !isActiveDocumentPending}
                    />
                    <DefaultButton text="Clear" onClick={this._clearSignature} disabled={!isActiveDocumentPending} />
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
            <DefaultButton text={this.state.isRejecting ? 'Rejecting...' : 'Reject'} onClick={this._openRejectDialog} disabled={!canReject} title={isActiveDocumentPending ? undefined : 'Available only for Pending items'} />
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

        <Dialog
          hidden={!this.state.isApprovalSuccessDialogOpen}
          onDismiss={this._closeApprovalSuccessDialog}
          dialogContentProps={{
            title: 'Document is approved'
          }}
          modalProps={{
            isBlocking: true
          }}
        >
          <DialogFooter>
            <PrimaryButton text="OK" onClick={this._closeApprovalSuccessDialog} />
          </DialogFooter>
        </Dialog>

        <Dialog
          hidden={!this.state.isTaskLinkMessageDialogOpen}
          onDismiss={this._acknowledgeTaskLinkMessageDialog}
          dialogContentProps={{
            title: this.state.taskLinkMessageDialogTitle,
            subText: this.state.taskLinkMessageDialogText
          }}
          modalProps={{
            isBlocking: true
          }}
        >
          <DialogFooter>
            <PrimaryButton text="OK" onClick={this._acknowledgeTaskLinkMessageDialog} />
          </DialogFooter>
        </Dialog>
      </section>
    );
  }

  // Mark that drawing has started on the signature pad.
  private readonly _handleSignaturePointerDown = (): void => {
    this._isSignaturePointerDown = true;
  };

  // Mark that drawing has ended and reset signature pad interaction.
  private readonly _handleSignaturePointerUp = (): void => {
    this._isSignaturePointerDown = false;
    this._resetSignaturePadInteraction();
  };

  // Reset hover state if the mouse moves without drawing pressed.
  private readonly _handleSignatureHoverGuard = (event: React.MouseEvent<HTMLCanvasElement>): void => {
    if (!this._isSignaturePointerDown && event.buttons === 0) {
      this._resetSignaturePadInteraction();
    }
  };

  // Rebind signature pad handlers to avoid stuck drawing behavior.
  private readonly _resetSignaturePadInteraction = (): void => {
    const signaturePad = this._signatureRef.current?.getSignaturePad?.();

    if (!signaturePad) {
      return;
    }

    signaturePad.off();
    signaturePad.on();
  };

  // Capture and trim the signature image from the signature pad.
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

    const signatureCanvas: HTMLCanvasElement | undefined = getTrimmedSignatureCanvas(signaturePad);

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

  // Clear the signature pad and remove current signature placement.
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

  // Update signature width and keep placement width in sync.
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

  // Place the signature on the PDF preview where the user clicks.
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

  // Move to the previous PDF page.
  private readonly _goToPreviousPage = (): void => {
    this.setState(previousState => ({
      currentPageIndex: Math.max(0, previousState.currentPageIndex - 1)
    }));
  };

  // Move to the next PDF page.
  private readonly _goToNextPage = (): void => {
    this.setState(previousState => ({
      currentPageIndex: Math.min(previousState.pageCount - 1, previousState.currentPageIndex + 1)
    }));
  };

  // Re-render the current page and restore signature preview after resize.
  private readonly _handleWindowResize = (): void => {
    if (!this.state.pdfDocument) {
      return;
    }

    const signatureDataUrlBeforeResize: string | undefined = this._getCurrentSignatureDataUrl();

    if (this._resizeTimer !== undefined) {
      window.clearTimeout(this._resizeTimer);
    }

    this._resizeTimer = window.setTimeout(() => {
      this._renderCurrentPage().then(async () => {
        if (!signatureDataUrlBeforeResize) {
          return;
        }

        try {
          await this._populateSignaturePad(signatureDataUrlBeforeResize);
        } catch {
          this.setState({
            statusMessage: 'The signature preview could not be restored after resizing.'
          });
        }
      }).catch(() => {
        this.setState({
          isRendering: false,
          statusMessage: 'The selected PDF page could not be resized.'
        });
      });
    }, 150);
  };

  // Open a selected document and load its PDF and saved signature.
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
      signatureAspectRatio: 2.8,
      rejectionComments: '',
      isLoadingSavedSignature: true,
      isLoading: true,
      statusMessage: 'Loading PDF and saved signature...'
    });

    await Promise.all([
      this._loadSavedSignatureFromSharePoint(),
      this._loadDocumentFromSharePoint(item.id, item.attachmentServerRelativeUrl)
    ]);

    if (this.state.savedSignatureDataUrl) {
      try {
        await this._populateSignaturePad(this.state.savedSignatureDataUrl);
      } catch {
        this.setState({
          statusMessage: 'The saved signature preview could not be loaded into the canvas.'
        });
      }
    }
  };

  // Open the approval history for a selected document.
  private readonly _showApprovalHistory = async (item: IDocumentListItem): Promise<void> => {
    const requestId: number = ++this._approvalHistoryRequestId;

    this.setState({
      approvalHistoryDocument: item,
      approvalHistory: [],
      approvalHistoryError: '',
      isApprovalHistoryLoading: true,
      isApprovalHistoryOpen: true
    });

    try {
      const items = await this._getApprovalHistoryItems(item);

      if (requestId !== this._approvalHistoryRequestId) {
        return;
      }

      this.setState({
        approvalHistory: items,
        isApprovalHistoryLoading: false
      });
    } catch (error) {
      if (requestId !== this._approvalHistoryRequestId) {
        return;
      }

      this.setState({
        approvalHistory: [],
        isApprovalHistoryLoading: false,
        approvalHistoryError: `The approval history could not be loaded. ${this._getErrorMessage(error)}`.trim()
      });
    }
  };

  private readonly _closeApprovalHistory = (): void => {
    this._blurActiveElement();
    this._approvalHistoryRequestId++;

    this.setState({
      isApprovalHistoryOpen: false,
      isApprovalHistoryLoading: false,
      approvalHistoryError: '',
      approvalHistory: [],
      approvalHistoryDocument: undefined
    });
  };

  // Close the main dialog and clear active document state.
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

  // Open the rejection dialog for the active document.
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

  // Close the rejection dialog and clear rejection comments.
  private readonly _closeRejectDialog = (): void => {
    this.setState({
      isRejectDialogOpen: false,
      rejectionComments: ''
    });
  };

  // Close the approval success message dialog.
  private readonly _closeApprovalSuccessDialog = (): void => {
    this.setState({
      isApprovalSuccessDialogOpen: false
    });
  };

  // Confirm the task link message and clean up URL state.
  private readonly _acknowledgeTaskLinkMessageDialog = (): void => {
    this._removeTaskIdFromUrl();

    this.setState({
      isTaskLinkMessageDialogOpen: false,
      taskLinkMessageDialogTitle: '',
      taskLinkMessageDialogText: ''
    });
  };

  // Sign the PDF, update task status, and complete approval flow.
  private readonly _approveDocument = async (): Promise<void> => {
    const activeDocument = this.state.activeDocument;

    if (!activeDocument) {
      this.setState({
        statusMessage: 'Select a document before approving it.'
      });
      return;
    }

    const approvalUpdate = getApprovalUpdate(activeDocument.approvalStatus);

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
        IsTaskActive: false,
        ActionTakenOn: new Date().toISOString()
      });

      if (activeDocument.documentId && activeDocument.documentNumber) {
        const relatedItems = await this._getRelatedDocumentItems(
          apiBaseUrl,
          listTitle,
          activeDocument.documentId,
          activeDocument.documentNumber,
          activeDocument.id,
          activeDocument.approvalStatus || ''
        );

        for (let i = 0; i < relatedItems.length; i++) {
          await this._updateSharePointListItem(apiBaseUrl, listTitle, relatedItems[i], {
            ApprovalStatus: '',
            IsTaskActive: false
          });
        }
      }

      await this._loadDocumentListFromSharePoint();
      this._removeTaskIdFromUrl();

      this.setState({
        isDialogOpen: false,
        isRejectDialogOpen: false,
        isApprovalSuccessDialogOpen: true,
        activeDocument: undefined,
        isApproving: false,
        rejectionComments: '',
        statusMessage: 'Document is approved.'
      });
    } catch (error) {
      this.setState({
        isApproving: false,
        statusMessage: `The document could not be approved. ${this._getErrorMessage(error)}`
      });
    }
  };

  // Submit rejection comments and update task status in SharePoint.
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

    const rejectionUpdate = getRejectionUpdate(activeDocument.approvalStatus);

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
        IsTaskActive: false,
        ActionTakenOn: new Date().toISOString()
      });

      if (activeDocument.documentId && activeDocument.documentNumber) {
        const relatedItems = await this._getRelatedDocumentItems(
          apiBaseUrl,
          listTitle,
          activeDocument.documentId,
          activeDocument.documentNumber,
          activeDocument.id,
          activeDocument.approvalStatus || ''
        );

        for (let i = 0; i < relatedItems.length; i++) {
          await this._updateSharePointListItem(apiBaseUrl, listTitle, relatedItems[i], {
            ApprovalStatus: '',
            IsTaskActive: false
          });
        }
      }

      await this._loadDocumentListFromSharePoint();
      this._removeTaskIdFromUrl();

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

  // Load assigned documents from SharePoint for the current user.
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
    const listInfo = getListApiPath(listName);
    const currentUserEmail = this.props.userEmail?.trim() || '';
    const lastDateToFetchFrom = new Date(Date.now() - 45 * 24 * 60 * 60 * 1000).toISOString();
    const signatoryFilter = currentUserEmail ? ` and L1Signatory eq '${currentUserEmail.replace(/'/g, "''")}'` : '';

    try {
      const requestUrl = `${apiBaseUrl}/_api/${listInfo.apiPath}/items?$select=Id,Title,DocumentID,DocumentName,DocumentNumber,Trader,AccountCode,ApprovalStatus,Comments,IsTaskActive,Created,Modified,L1Signatory,AttachmentFiles&$expand=AttachmentFiles&$top=500&$filter=Created ge datetime'${lastDateToFetchFrom}'${signatoryFilter}`;
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
      const items = (itemsJson && (itemsJson.value as Array<ISharePointDocumentItem>)) || [];

      const documents: IDocumentListItem[] = items
        .map((item: ISharePointDocumentItem) => this._mapSharePointDocumentItem(item))
        .filter((item: IDocumentListItem | undefined): item is IDocumentListItem => !!item);

      this.setState({
        documents,
        isDocumentsLoading: false,
        statusMessage: documents.length === 0
          ? `No signed documents with PDF attachments were found in '${listName}'. Request URL: ${requestUrl}`
          : 'Select a document to open for signature.'
      });

      await this._processTaskIdDeepLink(documents);
    } catch (error) {
      this.setState({
        documents: [],
        isDocumentsLoading: false,
        statusMessage: `The documents could not be loaded. ${this._getErrorMessage(error)}`
      });
    }
  };

  // Convert one SharePoint item to the UI document model.
  private readonly _mapSharePointDocumentItem = (item: ISharePointDocumentItem): IDocumentListItem | undefined => {
    const attachmentFiles: Array<{ FileName: string; ServerRelativeUrl: string }> = item.AttachmentFiles || [];
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

    if (!pdfAttachment) {
      return undefined;
    }

    return {
      id: item.Id,
      title: item.Title,
      documentId: item.DocumentID,
      documentName: item.DocumentName,
      documentNumber: item.DocumentNumber,
      trader: item.Trader,
      accountCode: item.AccountCode,
      approverEmail: item.L1Signatory,
      approvalStatus: item.ApprovalStatus,
      comments: item.Comments,
      isTaskActive: item.IsTaskActive,
      created: item.Created,
      modified: item.Modified,
      attachmentFileName: pdfAttachment.FileName,
      attachmentServerRelativeUrl: pdfAttachment.ServerRelativeUrl
    };
  };

  // Read and validate TaskID from the page query string.
  private readonly _getTaskIdFromQuery = (): { isPresent: boolean; value?: number } => {
    if (typeof window === 'undefined') {
      return { isPresent: false };
    }

    const queryString = window.location.search ? window.location.search.substring(1) : '';

    if (!queryString) {
      return { isPresent: false };
    }

    const queryPairs = queryString.split('&');
    const taskIdQueryKey = SignatureWebPart._taskIdQueryParamName.toLowerCase();
    let rawValue: string | undefined;

    for (let i = 0; i < queryPairs.length; i++) {
      const pair = queryPairs[i];

      if (!pair) {
        continue;
      }

      const equalSignIndex = pair.indexOf('=');
      const rawKey = equalSignIndex === -1 ? pair : pair.substring(0, equalSignIndex);
      const normalizedKey = this._safeDecodeQueryValue(rawKey).toLowerCase();

      if (normalizedKey === taskIdQueryKey) {
        rawValue = equalSignIndex === -1 ? '' : this._safeDecodeQueryValue(pair.substring(equalSignIndex + 1));
        break;
      }
    }

    if (rawValue === undefined) {
      return { isPresent: false };
    }

    const normalizedRawValue = rawValue.trim();

    if (!normalizedRawValue || !/^\d+$/.test(normalizedRawValue)) {
      return { isPresent: true };
    }

    const parsedValue = parseInt(normalizedRawValue, 10);

    if (parsedValue <= 0) {
      return { isPresent: true };
    }

    return {
      isPresent: true,
      value: parsedValue
    };
  };

  // Remove TaskID from the URL without reloading the page.
  private readonly _removeTaskIdFromUrl = (): void => {
    if (typeof window === 'undefined') {
      return;
    }

    const currentUrl = new URL(window.location.href);
    const queryString = currentUrl.search ? currentUrl.search.substring(1) : '';

    if (!queryString) {
      return;
    }

    const queryPairs = queryString.split('&');
    const taskIdQueryKey = SignatureWebPart._taskIdQueryParamName.toLowerCase();
    const nextPairs: string[] = [];
    let removedTaskId = false;

    for (let i = 0; i < queryPairs.length; i++) {
      const pair = queryPairs[i];

      if (!pair) {
        continue;
      }

      const equalSignIndex = pair.indexOf('=');
      const rawKey = equalSignIndex === -1 ? pair : pair.substring(0, equalSignIndex);
      const normalizedKey = this._safeDecodeQueryValue(rawKey).toLowerCase();

      if (normalizedKey === taskIdQueryKey) {
        removedTaskId = true;
        continue;
      }

      nextPairs.push(pair);
    }

    if (!removedTaskId) {
      return;
    }

    const nextQuery = nextPairs.join('&');
    const nextUrl = `${currentUrl.pathname}${nextQuery ? `?${nextQuery}` : ''}${currentUrl.hash}`;
    window.history.replaceState(window.history.state, document.title, nextUrl);
  };

  // Decode query text safely and handle invalid encoded values.
  private readonly _safeDecodeQueryValue = (value: string): string => {
    const normalized = value.replace(/\+/g, ' ');

    try {
      return decodeURIComponent(normalized);
    } catch {
      return normalized;
    }
  };

  // Show a simple dialog message for task link status.
  private readonly _showTaskLinkMessageDialog = (title: string, text: string): void => {
    this.setState({
      isTaskLinkMessageDialogOpen: true,
      taskLinkMessageDialogTitle: title,
      taskLinkMessageDialogText: text
    });
  };

  // Process TaskID link, validate user, and auto-open the task.
  private readonly _processTaskIdDeepLink = async (documents: IDocumentListItem[]): Promise<void> => {
    if (this._hasHandledTaskIdDeepLink || this._isHandlingTaskIdDeepLink) {
      return;
    }

    const taskIdFromQuery = this._getTaskIdFromQuery();

    // No TaskID means normal page load path with no deep-link behavior.
    if (!taskIdFromQuery.isPresent) {
      return;
    }

    this._isHandlingTaskIdDeepLink = true;

    try {
      if (!taskIdFromQuery.value) {
        this._hasHandledTaskIdDeepLink = true;
        this._showTaskLinkMessageDialog('Invalid Task Link', 'The TaskID query string value is invalid.');
        return;
      }

      const taskId = taskIdFromQuery.value;
      const currentUserEmail = (this.props.userEmail || '').trim().toLowerCase();

      if (!currentUserEmail) {
        this._hasHandledTaskIdDeepLink = true;
        this._showTaskLinkMessageDialog('Unauthorized', 'Unable to validate approver because the current user email is not available.');
        return;
      }

      let targetTask: IDocumentListItem | undefined;

      for (let i = 0; i < documents.length; i++) {
        if (documents[i].id === taskId) {
          targetTask = documents[i];
          break;
        }
      }

      if (!targetTask) {
        targetTask = await this._getTaskItemByIdFromSharePoint(taskId);
      }

      if (!targetTask) {
        this._hasHandledTaskIdDeepLink = true;
        this._showTaskLinkMessageDialog('Task Not Found', `TaskID ${taskId} was not found.`);
        return;
      }

      const approverEmail = (targetTask.approverEmail || '').trim().toLowerCase();

      if (!approverEmail || approverEmail !== currentUserEmail) {
        this._hasHandledTaskIdDeepLink = true;
        this._showTaskLinkMessageDialog('Unauthorized', 'You are not the approver for this task.');
        return;
      }

      this._hasHandledTaskIdDeepLink = true;
      this._removeTaskIdFromUrl();
      await this._openDocument(targetTask);
    } catch (error) {
      this._hasHandledTaskIdDeepLink = true;
      this._showTaskLinkMessageDialog('Task Link Error', `The task link could not be processed. ${this._getErrorMessage(error)}`.trim());
    } finally {
      this._isHandlingTaskIdDeepLink = false;
    }
  };

  // Get one task item by ID directly from SharePoint.
  private readonly _getTaskItemByIdFromSharePoint = async (taskId: number): Promise<IDocumentListItem | undefined> => {
    const { siteUrl, taskListName, spHttpClient, webAbsoluteUrl } = this.props;
    const trimmedSiteUrl: string = siteUrl ? siteUrl.trim() : webAbsoluteUrl;
    const listName: string | undefined = taskListName ? taskListName.trim() : undefined;

    if (!trimmedSiteUrl || !listName) {
      return undefined;
    }

    const apiBaseUrl: string = trimmedSiteUrl.replace(/\/$/, '');
    const listInfo = getListApiPath(listName);
    const requestUrl = `${apiBaseUrl}/_api/${listInfo.apiPath}/items(${taskId})?$select=Id,Title,DocumentID,DocumentName,DocumentNumber,Trader,AccountCode,ApprovalStatus,Comments,IsTaskActive,Created,Modified,L1Signatory,AttachmentFiles&$expand=AttachmentFiles`;
    const response = await spHttpClient.get(
      requestUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );

    if (response.status === 404) {
      return undefined;
    }

    if (!response.ok) {
      throw new Error(`Unable to read task ${taskId}. Status ${response.status}`);
    }

    const item = await response.json() as ISharePointDocumentItem;
    return this._mapSharePointDocumentItem(item);
  };

  // Load PDF bytes from SharePoint and prepare PDF viewer state.
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
      statusMessage: 'Loading document... ',
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
        `${apiBaseUrl}/_api/web/GetFileByServerRelativeUrl('${normalizeServerRelativeUrlForApi(pdfAttachment.ServerRelativeUrl)}')/$value`,
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
        statusMessage: ""//`Loaded ${pageCount} page${pageCount === 1 ? '' : 's'} from SharePoint.`
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

  // Render the selected PDF page to the preview canvas.
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


  // Return a readable error message from unknown error values.
  private _getErrorMessage(error: unknown): string {
    return error instanceof Error && error.message
      ? error.message
      : '';
  }

  // Remove focus from current input to avoid stale UI focus.
  private _blurActiveElement(): void {
    const activeElement = document.activeElement;

    if (activeElement instanceof HTMLElement) {
      activeElement.blur();
    }
  }

  // Build signed PDF bytes by stamping the placed signature.
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

  // Draw signature image on the selected PDF page coordinates.
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

    const pngBytes = dataUrlToUint8Array(signatureDataUrl);
    const signatureImage = await pdfDocument.embedPng(pngBytes);
    const pageWidth = page.getWidth();
    const pageHeight = page.getHeight();
    const signatureHeight = getSignatureHeight(placement.width, signatureAspectRatio);
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

  // Reload source PDF bytes when they are not available in state.
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
    const normalizedAttachmentUrl = normalizeServerRelativeUrlForApi(activeDocument.attachmentServerRelativeUrl);

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

  // Replace existing list attachment with newly signed PDF file.
  private readonly _replaceCurrentAttachmentWithPdf = async (
    apiBaseUrl: string,
    listTitle: string,
    activeDocument: IDocumentListItem,
    signedPdfBytes: Uint8Array
  ): Promise<void> => {
    const currentAttachmentName = activeDocument.attachmentFileName || getFileNameFromServerRelativeUrl(activeDocument.attachmentServerRelativeUrl) || `${activeDocument.documentName || activeDocument.title || 'document'}.pdf`;
    const encodedAttachmentName = encodeURIComponent(currentAttachmentName);

    // Overwrite the existing attachment's bytes in place (single request) instead of delete+add,
    // which avoids bumping the list item version twice for one signing action.
    await this._updateAttachmentFileContent(apiBaseUrl, listTitle, activeDocument.id, encodedAttachmentName, signedPdfBytes);
  };

  // Save current signature image to Signature Master list.
  private readonly _saveSignatureToSharePoint = async (): Promise<void> => {
    const { siteUrl, webAbsoluteUrl } = this.props;
    const trimmedSiteUrl: string = siteUrl ? siteUrl.trim() : webAbsoluteUrl;
    const userEmail: string = (this.props.userEmail || '').trim();

    if (!trimmedSiteUrl) {
      this.setState({
        statusMessage: 'The site URL is not configured.'
      });
      return;
    }

    if (!userEmail) {
      this.setState({
        statusMessage: 'The current user email is not available. Unable to save the signature.'
      });
      return;
    }

    const signatureDataUrl = this._getCurrentSignatureDataUrl();

    if (!signatureDataUrl) {
      this.setState({
        statusMessage: 'Draw or load a signature before saving it.'
      });
      return;
    }

    this.setState({
      isSavingSignature: true,
      statusMessage: 'Saving signature to Signature Master...'
    });

    const apiBaseUrl: string = trimmedSiteUrl.replace(/\/$/, '');

    try {
      const signatureBytes = dataUrlToUint8Array(signatureDataUrl);
      const aspectRatio = await this._getImageAspectRatio(signatureDataUrl);
      const item = await this._getOrCreateSignatureMasterItem(apiBaseUrl, userEmail);
      const fileName = this._getSignatureFileName(userEmail);

      for (let i = 0; i < item.attachments.length; i++) {
        await this._deleteAttachmentFile(
          apiBaseUrl,
          SignatureWebPart._signatureMasterListName.replace(/'/g, "''"),
          item.id,
          encodeURIComponent(item.attachments[i].FileName)
        );
      }

      await this._addAttachmentFile(
        apiBaseUrl,
        SignatureWebPart._signatureMasterListName.replace(/'/g, "''"),
        item.id,
        fileName,
        signatureBytes,
        'image/png'
      );

      this.setState({
        isSavingSignature: false,
        signatureDataUrl,
        signatureAspectRatio: aspectRatio,
        savedSignatureDataUrl: signatureDataUrl,
        savedSignatureAspectRatio: aspectRatio,
        statusMessage: 'Signature saved to Signature Master.'
      });
    } catch (error) {
      this.setState({
        isSavingSignature: false,
        statusMessage: `Unable to save signature to Signature Master. ${this._getErrorMessage(error)}`
      });
    }
  };

  // Delete one attachment file from a SharePoint list item.
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

  // Upload one attachment file to a SharePoint list item.
  private readonly _addAttachmentFile = async (
    apiBaseUrl: string,
    listTitle: string,
    itemId: number,
    fileName: string,
    fileBytes: Uint8Array,
    contentType: string = 'application/pdf'
  ): Promise<void> => {
    const encodedFileName = encodeURIComponent(fileName);
    const response = await this.props.spHttpClient.post(
      `${apiBaseUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})/AttachmentFiles/add(FileName='${encodedFileName}')`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': contentType
        },
        body: new Blob([fileBytes], { type: contentType })
      }
    );

    if (!response.ok) {
      throw new Error(`Unable to upload the signed attachment. Status ${response.status}`);
    }
  };

  // Overwrite an existing attachment's content in place, without a delete+add round trip.
  private readonly _updateAttachmentFileContent = async (
    apiBaseUrl: string,
    listTitle: string,
    itemId: number,
    encodedAttachmentName: string,
    fileBytes: Uint8Array,
    contentType: string = 'application/pdf'
  ): Promise<void> => {
    const response = await this.props.spHttpClient.post(
      `${apiBaseUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})/AttachmentFiles('${encodedAttachmentName}')/$value`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': contentType,
          'X-HTTP-Method': 'PUT'
        },
        body: new Blob([fileBytes], { type: contentType })
      }
    );

    if (!response.ok) {
      throw new Error(`Unable to update the existing attachment. Status ${response.status}`);
    }
  };

  // Get signature image from pad if possible, else use saved state.
  private _getCurrentSignatureDataUrl(): string | undefined {
    const signaturePad: SignatureCanvas | null = this._signatureRef.current;

    if (signaturePad && !signaturePad.isEmpty()) {
      const trimmedCanvas = getTrimmedSignatureCanvas(signaturePad);

      if (trimmedCanvas) {
        return trimmedCanvas.toDataURL('image/png');
      }
    }

    return this.state.signatureDataUrl || undefined;
  }

  // Build a safe PNG file name using the user email.
  private _getSignatureFileName(userEmail: string): string {
    const sanitizedEmail = userEmail.replace(/[^a-zA-Z0-9._-]/g, '_');
    return `signature_${sanitizedEmail}.png`;
  }

  // Read image size and return width-to-height ratio.
  private async _getImageAspectRatio(dataUrl: string): Promise<number> {
    const image = await this._loadImageFromDataUrl(dataUrl);

    if (!image.naturalWidth || !image.naturalHeight) {
      return 2.8;
    }

    return image.naturalWidth / image.naturalHeight;
  }

  // Load an image element from a data URL string.
  private _loadImageFromDataUrl(dataUrl: string): Promise<HTMLImageElement> {
    return new Promise((resolve, reject) => {
      const image = new Image();

      image.onload = () => resolve(image);
      image.onerror = () => reject(new Error('Unable to decode image data URL.'));
      image.src = dataUrl;
    });
  }

  // Resize signature pad canvas to match display and device ratio.
  private _syncSignaturePadCanvasSize(signaturePad: SignatureCanvas): HTMLCanvasElement {
    const targetCanvas = signaturePad.getCanvas();
    const ratio = Math.max(window.devicePixelRatio || 1, 1);
    const nextWidth = Math.max(1, Math.round(targetCanvas.offsetWidth * ratio));
    const nextHeight = Math.max(1, Math.round(targetCanvas.offsetHeight * ratio));

    if (targetCanvas.width !== nextWidth || targetCanvas.height !== nextHeight) {
      targetCanvas.width = nextWidth;
      targetCanvas.height = nextHeight;

      const resizedContext = targetCanvas.getContext('2d');

      if (resizedContext) {
        resizedContext.scale(ratio, ratio);
      }
    }

    return targetCanvas;
  }

  // Draw saved signature image into the signature pad canvas.
  private async _populateSignaturePad(signatureDataUrl: string): Promise<void> {
    const signaturePad: SignatureCanvas | null = this._signatureRef.current;

    if (!signaturePad) {
      return;
    }

    const sourceImage = await this._loadImageFromDataUrl(signatureDataUrl);
    const targetCanvas = this._syncSignaturePadCanvasSize(signaturePad);
    const context = targetCanvas.getContext('2d');

    if (!context) {
      return;
    }

    signaturePad.clear();

    const padding = 8;
    const canvasWidth = Math.max(1, targetCanvas.offsetWidth);
    const canvasHeight = Math.max(1, targetCanvas.offsetHeight);
    const availableWidth = Math.max(1, canvasWidth - (padding * 2));
    const availableHeight = Math.max(1, canvasHeight - (padding * 2));
    const imageAspectRatio = sourceImage.naturalWidth > 0 && sourceImage.naturalHeight > 0
      ? sourceImage.naturalWidth / sourceImage.naturalHeight
      : 2.8;

    let drawWidth = availableWidth;
    let drawHeight = drawWidth / imageAspectRatio;

    if (drawHeight > availableHeight) {
      drawHeight = availableHeight;
      drawWidth = drawHeight * imageAspectRatio;
    }

    const offsetX = (canvasWidth - drawWidth) / 2;
    const offsetY = (canvasHeight - drawHeight) / 2;

    context.drawImage(sourceImage, offsetX, offsetY, drawWidth, drawHeight);
  }

  // Load saved signature for current user from Signature Master.
  private readonly _loadSavedSignatureFromSharePoint = async (): Promise<void> => {
    const { siteUrl, webAbsoluteUrl } = this.props;
    const userEmail: string = (this.props.userEmail || '').trim();
    const trimmedSiteUrl: string = siteUrl ? siteUrl.trim() : webAbsoluteUrl;

    if (!trimmedSiteUrl || !userEmail) {
      this.setState({
        isLoadingSavedSignature: false,
        savedSignatureDataUrl: '',
        savedSignatureAspectRatio: 2.8
      });
      return;
    }

    this.setState({
      isLoadingSavedSignature: true
    });

    const apiBaseUrl: string = trimmedSiteUrl.replace(/\/$/, '');

    try {
      const item = await this._tryGetSignatureMasterItem(apiBaseUrl, userEmail);

      if (!item || item.attachments.length === 0) {
        this.setState({
          isLoadingSavedSignature: false,
          savedSignatureDataUrl: '',
          savedSignatureAspectRatio: 2.8
        });
        return;
      }

      let attachment = item.attachments[0];

      for (let i = 0; i < item.attachments.length; i++) {
        const fileName = item.attachments[i].FileName.toLowerCase();
        if (fileName.length > 4 && fileName.substr(fileName.length - 4) === '.png') {
          attachment = item.attachments[i];
          break;
        }
      }

      const fileResponse = await this.props.spHttpClient.get(
        `${apiBaseUrl}/_api/web/GetFileByServerRelativeUrl('${normalizeServerRelativeUrlForApi(attachment.ServerRelativeUrl)}')/$value`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'image/png'
          }
        }
      );

      if (!fileResponse.ok) {
        throw new Error(`Unable to read saved signature attachment. Status ${fileResponse.status}`);
      }

      const bytes = new Uint8Array(await fileResponse.arrayBuffer());
      const signatureDataUrl = uint8ArrayToDataUrl(bytes, 'image/png');
      const aspectRatio = await this._getImageAspectRatio(signatureDataUrl);

      this.setState({
        isLoadingSavedSignature: false,
        savedSignatureDataUrl: signatureDataUrl,
        savedSignatureAspectRatio: aspectRatio,
        signatureDataUrl: signatureDataUrl,
        signatureAspectRatio: aspectRatio
      });
    } catch (error) {
      this.setState({
        isLoadingSavedSignature: false,
        savedSignatureDataUrl: '',
        savedSignatureAspectRatio: 2.8,
        statusMessage: `Unable to load signature from Signature Master. ${this._getErrorMessage(error)}`
      });
    }
  };

  // Try to find the user's Signature Master item.
  private async _tryGetSignatureMasterItem(
    apiBaseUrl: string,
    userEmail: string
  ): Promise<{ id: number; attachments: Array<{ FileName: string; ServerRelativeUrl: string }> } | undefined> {
    const listTitle = SignatureWebPart._signatureMasterListName.replace(/'/g, "''");
    const safeEmail = userEmail.replace(/'/g, "''");
    const queryUrl = `${apiBaseUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=Id,Title,AttachmentFiles&$expand=AttachmentFiles&$filter=Title eq '${safeEmail}'&$top=1`;
    const response = await this.props.spHttpClient.get(
      queryUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );

    if (!response.ok) {
      throw new Error(`Unable to query Signature Master list. Status ${response.status}`);
    }

    const json = await response.json();
    const items = (json && (json.value as Array<{
      Id: number;
      Title?: string;
      AttachmentFiles?: Array<{ FileName: string; ServerRelativeUrl: string }>;
    }>)) || [];

    if (items.length === 0) {
      return undefined;
    }

    return {
      id: items[0].Id,
      attachments: items[0].AttachmentFiles || []
    };
  }

  // Get existing Signature Master item or create a new one.
  private async _getOrCreateSignatureMasterItem(
    apiBaseUrl: string,
    userEmail: string
  ): Promise<{ id: number; attachments: Array<{ FileName: string; ServerRelativeUrl: string }> }> {
    const existingItem = await this._tryGetSignatureMasterItem(apiBaseUrl, userEmail);

    if (existingItem) {
      return existingItem;
    }

    const listTitle = SignatureWebPart._signatureMasterListName.replace(/'/g, "''");
    const createResponse = await this.props.spHttpClient.post(
      `${apiBaseUrl}/_api/web/lists/getbytitle('${listTitle}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify({
          Title: userEmail
        })
      }
    );

    if (!createResponse.ok) {
      throw new Error(`Unable to create Signature Master list item. Status ${createResponse.status}`);
    }

    const createJson = await createResponse.json() as { Id?: number };

    if (!createJson.Id) {
      throw new Error('Signature Master list item was created but no item ID was returned.');
    }

    return {
      id: createJson.Id,
      attachments: []
    };
  }

  // Check if bytes start with the PDF file signature header.
  private _isPdfByteArray(bytes?: Uint8Array): boolean {
    return !!bytes &&
      bytes.length >= 4 &&
      bytes[0] === 0x25 &&
      bytes[1] === 0x50 &&
      bytes[2] === 0x44 &&
      bytes[3] === 0x46;
  }

  // Update fields on a SharePoint list item.
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

  // Keep rejection comment text in component state.
  private readonly _handleRejectionCommentsChange = (
    event: React.ChangeEvent<HTMLTextAreaElement>
  ): void => {
    this.setState({
      rejectionComments: event.target.value
    });
  };

  private readonly _getApprovalHistoryItems = async (document: IDocumentListItem): Promise<IApprovalHistoryItem[]> => {
    const { siteUrl, taskListName, spHttpClient, webAbsoluteUrl } = this.props;
    const trimmedSiteUrl: string = siteUrl ? siteUrl.trim() : webAbsoluteUrl;
    const listName: string | undefined = taskListName ? taskListName.trim() : undefined;

    if (!trimmedSiteUrl || !listName) {
      throw new Error('The SharePoint site URL and task list name are required.');
    }

    if (!document.documentId || !document.documentNumber) {
      throw new Error('Document ID and Document Number are required to find related approval history.');
    }

    const apiBaseUrl: string = trimmedSiteUrl.replace(/\/$/, '');
    const listInfo = getListApiPath(listName);
    const safeDocumentId = document.documentId.replace(/'/g, "''");
    const safeDocumentNumber = document.documentNumber.replace(/'/g, "''");
    const requestUrl = `${apiBaseUrl}/_api/${listInfo.apiPath}/items?$select=Id,DocumentID,DocumentNumber,L1Signatory,ApprovalStatus,Comments,ActionTakenOn,Created&$filter=DocumentID eq '${safeDocumentId}' and DocumentNumber eq '${safeDocumentNumber}'&$top=500`;
    const response = await spHttpClient.get(
      requestUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );

    if (!response.ok) {
      throw new Error(`Unable to query approval history. Status ${response.status}`);
    }

    const json = await response.json();
    const historyItems = (json && (json.value as Array<ISharePointDocumentItem>)) || [];
    const approvalHistory: IApprovalHistoryItem[] = historyItems.map((item: ISharePointDocumentItem): IApprovalHistoryItem => {
      const approvalStatus = item.ApprovalStatus || '';
      const actionTakenBy = approvalStatus === 'L1 Rejected' || approvalStatus === 'L2 Rejected' ||
        approvalStatus === 'L1 Signed' || approvalStatus === 'L2 Signed'
        ? item.L1Signatory
        : '';

      return {
        id: item.Id,
        documentId: item.DocumentID,
        documentNumber: item.DocumentNumber,
        approvalRecipient: item.L1Signatory,
        actionTakenBy,
        approvalStatus: item.ApprovalStatus,
        comments: item.Comments,
        actionTakenOn: item.ActionTakenOn,
        created: item.Created
      };
    });

    approvalHistory.sort((first: IApprovalHistoryItem, second: IApprovalHistoryItem) => {
      const firstTime = first.created ? new Date(first.created).getTime() : 0;
      const secondTime = second.created ? new Date(second.created).getTime() : 0;
      return secondTime - firstTime || second.id - first.id;
    });

    return approvalHistory;
  };

  // Find other items that share the same document IDs.
  private readonly _getRelatedDocumentItems = async (
    apiBaseUrl: string,
    listTitle: string,
    documentId: string,
    documentNumber: string,
    currentItemId: number,
    currentApprovalStatus: string
  ): Promise<number[]> => {
    const safeDocumentId = documentId.replace(/'/g, "''");
    const safeDocumentNumber = documentNumber.replace(/'/g, "''");
    const safeApprovalStatus = currentApprovalStatus.replace(/'/g, "''");
    const response = await this.props.spHttpClient.get(
      `${apiBaseUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=Id&$filter=DocumentID eq '${safeDocumentId}' and DocumentNumber eq '${safeDocumentNumber}' and ApprovalStatus eq '${safeApprovalStatus}' and Id ne ${currentItemId}&$top=500`,
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

}
