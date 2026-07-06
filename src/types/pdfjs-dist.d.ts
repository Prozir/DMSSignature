declare module 'pdfjs-dist/legacy/build/pdf.mjs' {
  export interface PDFDocumentLoadingTask {
    promise: Promise<unknown>;
  }

  export function getDocument(options: {
    data: Uint8Array;
    disableWorker?: boolean;
  }): PDFDocumentLoadingTask;
}

declare module 'pdfjs-dist/legacy/build/pdf.worker.mjs';
