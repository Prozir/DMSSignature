import SignatureCanvas from 'react-signature-canvas';

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