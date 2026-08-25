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

export interface ISignatureWebPartState {
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