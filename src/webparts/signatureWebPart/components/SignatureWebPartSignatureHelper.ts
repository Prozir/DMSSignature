import SignatureCanvas from 'react-signature-canvas';
import type { IPdfPage, IPdfViewport, ISignatureAnchorRect, ISignaturePlacement } from './SignatureWebPartInterfaces';

export function dataUrlToUint8Array(dataUrl: string): Uint8Array {
  const base64 = dataUrl.split(',')[1] || '';
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);

  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }

  return bytes;
}

export function uint8ArrayToDataUrl(bytes: Uint8Array, contentType: string): string {
  let binary = '';

  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }

  return `data:${contentType};base64,${btoa(binary)}`;
}

export function getSignatureHeight(width: number, aspectRatio: number): number {
  return Math.max(32, Math.round(width / aspectRatio));
}

// Matches e.g. "{{Sig_es_:signer1:signature:dimension(width=35mm, height=10mm)}}" with all whitespace stripped first,
// so incidental spaces/newlines inserted by PDF text extraction can never break the match.
const _anchorDimensionPattern: string = ':signature:dimension\\(width=([\\d.]+)mm,height=([\\d.]+)mm\\)\\}\\}';
const _defaultAnchorWidthMm: number = 35;
const _defaultAnchorHeightMm: number = 10;

export interface IAnchorPatternConfig {
  strictPattern: RegExp;
  loosePattern: RegExp;
}

// Map the approval level to the hidden anchor tag embedded in the PDF for that signer.
export function getAnchorPatternForApprovalStatus(approvalStatus?: string): IAnchorPatternConfig | undefined {
  const normalized: string = (approvalStatus || '').toLowerCase();

  if (normalized.indexOf('l1') !== -1) {
    return {
      strictPattern: new RegExp(`\\{\\{Sig_es_:signer1${_anchorDimensionPattern}`, 'i'),
      loosePattern: /sig_es_:?signer1/i
    };
  }

  if (normalized.indexOf('l2') !== -1) {
    return {
      strictPattern: new RegExp(`\\{\\{Sig_es_:signer2${_anchorDimensionPattern}`, 'i'),
      loosePattern: /sig_es_:?signer2/i
    };
  }

  return undefined;
}

// Locate the hidden anchor text on a rendered PDF page and return its canvas-pixel bounding box.
export async function findSignatureAnchorRect(
  page: IPdfPage,
  viewport: IPdfViewport,
  config: IAnchorPatternConfig,
  pageIndex: number
): Promise<ISignatureAnchorRect | undefined> {
  const textContent = await page.getTextContent();

  // Strip whitespace before matching so extraction quirks (extra spaces/newlines between glyph runs) never break the match.
  let strippedText: string = '';
  const strippedCharItemIndexes: number[] = [];

  for (let itemIndex = 0; itemIndex < textContent.items.length; itemIndex++) {
    const chars: string = textContent.items[itemIndex].str;

    for (let charIndex = 0; charIndex < chars.length; charIndex++) {
      if (/\s/.test(chars[charIndex])) {
        continue;
      }

      strippedText += chars[charIndex];
      strippedCharItemIndexes.push(itemIndex);
    }
  }

  let match: RegExpExecArray | null = config.strictPattern.exec(strippedText);
  let widthMm: number = _defaultAnchorWidthMm;
  let heightMm: number = _defaultAnchorHeightMm;

  if (match) {
    widthMm = parseFloat(match[1]);
    heightMm = parseFloat(match[2]);

    if (isNaN(widthMm) || isNaN(heightMm)) {
      widthMm = _defaultAnchorWidthMm;
      heightMm = _defaultAnchorHeightMm;
    }
  } else {
    // The dimension suffix didn't match exactly (formatting differences) - fall back to just locating the signer tag.
    match = config.loosePattern.exec(strippedText);
  }

  if (!match) {
    // eslint-disable-next-line no-console
    console.warn('Signature anchor tag not found in PDF text layer. Extracted text:', strippedText);
    return undefined;
  }

  // The anchor tag's own glyph run is invisible/tiny - use its position as the box origin, not its rendered size.
  const originItemIndex: number | undefined = strippedCharItemIndexes[match.index];

  if (originItemIndex === undefined) {
    return undefined;
  }

  const originItem = textContent.items[originItemIndex];
  const millimetersToPdfPoints: number = 72 / 25.4;
  const x0: number = originItem.transform[4];
  const y0: number = originItem.transform[5];
  const x1: number = x0 + (widthMm * millimetersToPdfPoints);
  const y1: number = y0 + (heightMm * millimetersToPdfPoints);
  const corner0: number[] = viewport.convertToViewportPoint(x0, y0);
  const corner1: number[] = viewport.convertToViewportPoint(x1, y1);
  const left: number = Math.min(corner0[0], corner1[0]);
  const top: number = Math.min(corner0[1], corner1[1]);
  const right: number = Math.max(corner0[0], corner1[0]);
  const bottom: number = Math.max(corner0[1], corner1[1]);

  return {
    pageIndex,
    left,
    top,
    width: right - left,
    height: bottom - top
  };
}

// Fit a signature of the given aspect ratio inside the anchor box, centered, preserving proportions.
export function getFittedPlacement(
  rect: ISignatureAnchorRect,
  aspectRatio: number,
  renderedWidth: number,
  renderedHeight: number
): ISignaturePlacement {
  const widthFromBoxWidth: number = rect.width;
  const widthFromBoxHeight: number = rect.height * aspectRatio;
  const width: number = Math.max(1, Math.min(widthFromBoxWidth, widthFromBoxHeight));
  const height: number = getSignatureHeight(width, aspectRatio);

  return {
    pageIndex: rect.pageIndex,
    x: rect.left + (rect.width - width) / 2,
    y: rect.top + (rect.height - height) / 2,
    width,
    renderedWidth,
    renderedHeight
  };
}

export function getTrimmedSignatureCanvas(signaturePad: SignatureCanvas): HTMLCanvasElement | undefined {
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