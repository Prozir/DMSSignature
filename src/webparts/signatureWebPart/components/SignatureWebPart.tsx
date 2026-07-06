import * as React from 'react';
import SignatureCanvas from 'react-signature-canvas';
import { PDFDocument } from 'pdf-lib';
import * as pdfJsLib from 'pdfjs-dist/legacy/build/pdf.mjs';
import 'pdfjs-dist/legacy/build/pdf.worker.mjs';
import styles from './SignatureWebPart.module.scss';
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
      statusMessage: 'Upload a PDF to begin.'
    };
  }

  public componentDidUpdate(
    _previousProps: ISignatureWebPartProps,
    previousState: ISignatureWebPartState
  ): void {
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
      statusMessage
    } = this.state;

    const canPlaceSignature: boolean = !!pageSize && signatureDataUrl.length > 0 && !isRendering;
    const canDownload: boolean = !!this.state.pdfBytes && !!placement && signatureDataUrl.length > 0 && !isLoading;
    const signatureHeight: number = this._getSignatureHeight(signatureWidth, signatureAspectRatio);

    return (
      <section className={`${styles.signatureWebPart} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.header}>
          <div>
            <h2>PDF signature</h2>
            <p>Upload a PDF, draw a signature, click the page to place it, then download the signed copy.</p>
          </div>
          <label className={styles.fileButton}>
            <input type="file" accept="application/pdf" onChange={this._handlePdfUpload} />
            Upload PDF
          </label>
        </div>

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
                min="60"
                max="220"
                step="10"
                value={signatureWidth}
                onChange={this._handleSignatureWidthChange}
              />
              <p className={styles.helpText}>
                {canPlaceSignature ? 'Click the PDF preview where the top-left of the signature should appear.' : 'Draw a signature and load a PDF before placing it.'}
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
                <div className={styles.emptyState}>{isLoading || isRendering ? 'Working...' : 'Upload a PDF to preview it here.'}</div>
              )}
            </div>

            <div className={styles.status} role="status">{statusMessage}</div>
          </main>
        </div>
      </section>
    );
  }

  private readonly _handlePdfUpload = async (event: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file: File | undefined = event.target.files && event.target.files.length > 0
      ? event.target.files[0]
      : undefined;

    if (!file) {
      return;
    }

    this.setState({
      isLoading: true,
      statusMessage: 'Loading PDF...',
      placement: undefined
    });

    const input: HTMLInputElement = event.target;

    try {
      const buffer: ArrayBuffer = await this._readFileAsArrayBuffer(file);
      const bytes: Uint8Array = new Uint8Array(buffer);
      const loadingTask = pdfJsLib.getDocument({
        data: new Uint8Array(bytes),
        disableWorker: true
      });
      const pdfDocument: unknown = await loadingTask.promise;
      const pageCount: number = (pdfDocument as { numPages: number }).numPages;

      this.setState({
        fileName: file.name,
        pdfBytes: bytes,
        pdfDocument,
        pageCount,
        currentPageIndex: 0,
        isLoading: false,
        statusMessage: `Loaded ${pageCount} page${pageCount === 1 ? '' : 's'}.`
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
        statusMessage: `Please choose a valid PDF file. ${this._getErrorMessage(error)}`
      });
    } finally {
      input.value = '';
    }
  };

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

  private readonly _downloadSignedPdf = async (): Promise<void> => {
    const { fileName, pdfBytes, placement, signatureDataUrl } = this.state;
    const canvas: HTMLCanvasElement | null = this._previewCanvasRef.current;

    if (!pdfBytes || !placement || !signatureDataUrl || !canvas) {
      this.setState({ statusMessage: 'Upload a PDF, draw a signature, and place it before downloading.' });
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

  private _readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise<ArrayBuffer>((resolve, reject) => {
      const reader: FileReader = new FileReader();

      reader.onload = (): void => {
        if (reader.result instanceof ArrayBuffer) {
          resolve(reader.result);
          return;
        }

        reject(new Error('The selected file could not be read.'));
      };
      reader.onerror = (): void => reject(reader.error || new Error('The selected file could not be read.'));
      reader.readAsArrayBuffer(file);
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
