import type { IDocumentListItem } from './DocumentsList';

export interface IApprovalHistoryItem {
  id: number;
  documentId?: string;
  documentNumber?: string;
  approvalRecipient?: string;
  actionTakenBy?: string;
  approvalStatus?: string;
  comments?: string;
  actionTakenOn?: string;
  created?: string;
}

export interface ISignaturePlacement {
  pageIndex: number;
  x: number;
  y: number;
  width: number;
  renderedWidth: number;
  renderedHeight: number;
}

export interface IPdfPageSize {
  width: number;
  height: number;
}

// Canvas-pixel bounding box of a detected hidden signature anchor (e.g. "Sig_es_:signer1").
export interface ISignatureAnchorRect {
  pageIndex: number;
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface IPdfTextItem {
  str: string;
  transform: number[];
  width: number;
  height: number;
}

export interface IPdfTextContent {
  items: IPdfTextItem[];
}

export interface IPdfViewport {
  width: number;
  height: number;
  convertToViewportPoint(x: number, y: number): number[];
}

export interface IPdfPage {
  getViewport(options: { scale: number }): IPdfViewport;
  render(options: { canvasContext: CanvasRenderingContext2D; viewport: unknown }): { promise: Promise<void> };
  getTextContent(): Promise<IPdfTextContent>;
}

export interface ISignatureWebPartState {
  fileName: string;
  pdfBytes?: Uint8Array;
  pdfDocument?: unknown;
  pageCount: number;
  currentPageIndex: number;
  pageSize?: IPdfPageSize;
  placement?: ISignaturePlacement;
  signatureAnchorRect?: ISignatureAnchorRect;
  isPlacementAutoDetected?: boolean;
  signatureDataUrl: string;
  signatureAspectRatio: number;
  signatureWidth: number;
  isLoading: boolean;
  isRendering: boolean;
  statusMessage: string;
  documents: IDocumentListItem[];
  isDocumentsLoading: boolean;
  isApprovalHistoryOpen: boolean;
  isApprovalHistoryLoading: boolean;
  approvalHistoryError: string;
  approvalHistory: IApprovalHistoryItem[];
  approvalHistoryDocument?: IDocumentListItem;
  isDialogOpen: boolean;
  isRejectDialogOpen: boolean;
  isApprovalSuccessDialogOpen: boolean;
  isTaskLinkMessageDialogOpen: boolean;
  taskLinkMessageDialogTitle: string;
  taskLinkMessageDialogText: string;
  isApproving: boolean;
  isRejecting: boolean;
  isSavingSignature: boolean;
  isLoadingSavedSignature: boolean;
  isSignatureCanvasEmpty: boolean;
  isSaveSignatureConfirmOpen: boolean;
  activeDocument?: IDocumentListItem;
  rejectionComments: string;
  savedSignatureDataUrl: string;
  savedSignatureAspectRatio: number;
}

export interface IApprovalUpdate {
  approvalStatus: string;
  comments: string;
}

export interface IRejectionUpdate {
  approvalStatus: string;
}