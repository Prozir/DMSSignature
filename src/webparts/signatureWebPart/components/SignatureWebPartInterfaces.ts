import type { IDocumentListItem } from './DocumentsList';

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