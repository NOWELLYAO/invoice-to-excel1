import * as pdfjsLib from "pdfjs-dist";
// @ts-ignore - Vite resolves this to the worker file's built URL
import pdfjsWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorkerUrl;

// Budget for the JSON body sent to /api/extract. Vercel Serverless Functions reject
// requests over ~4.5MB, so we stay comfortably under that (base64 + JSON overhead).
const SAFE_RAW_PDF_BYTES = 3 * 1024 * 1024; // below this, send the PDF as-is (best fidelity)
const TARGET_PAYLOAD_BYTES = 3.5 * 1024 * 1024; // budget for rasterized fallback
const MAX_PAGES = 40; // very long catalogs: only the first N pages are sent

export interface PreparedFile {
  base64?: string;
  mimeType?: string;
  images?: { base64: string; mimeType: string }[];
  pagesUsed?: number;
  pagesTotal?: number;
}

async function readBytes(file: File): Promise<Uint8Array> {
  const buffer = await file.arrayBuffer();
  let bytes = new Uint8Array(buffer);

  // Corrige un défaut d'en-tête PDF observé sur certains exports (ex: "%%PDF-" au lieu de "%PDF-").
  if (file.type === "application/pdf") {
    const header = new TextDecoder().decode(bytes.slice(0, 8));
    if (header.startsWith("%%PDF-")) {
      bytes = bytes.slice(1);
    }
  }
  return bytes;
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

async function rasterizePdf(bytes: Uint8Array): Promise<PreparedFile> {
  const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
  const pagesTotal = pdf.numPages;
  const pagesUsed = Math.min(pagesTotal, MAX_PAGES);

  let quality = 0.6;
  let scale = 1.3;

  for (let attempt = 0; attempt < 3; attempt++) {
    const images: { base64: string; mimeType: string }[] = [];
    let totalBytes = 0;

    for (let i = 1; i <= pagesUsed; i++) {
      const page = await pdf.getPage(i);
      const viewport = page.getViewport({ scale });
      const canvas = document.createElement("canvas");
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext("2d")!;
      await page.render({ canvasContext: ctx, viewport }).promise;

      const dataUrl = canvas.toDataURL("image/jpeg", quality);
      const base64 = dataUrl.split(",")[1];
      totalBytes += base64.length;
      images.push({ base64, mimeType: "image/jpeg" });
    }

    if (totalBytes <= TARGET_PAYLOAD_BYTES || attempt === 2) {
      return { images, pagesUsed, pagesTotal };
    }
    // Trop lourd : on réduit qualité et résolution puis on réessaie.
    quality = Math.max(0.3, quality - 0.15);
    scale = Math.max(0.8, scale - 0.25);
  }

  // Ne devrait pas arriver, mais garde un filet de sécurité.
  return { images: [], pagesUsed, pagesTotal };
}

export async function prepareFileForApi(file: File): Promise<PreparedFile> {
  const bytes = await readBytes(file);

  if (file.type === "application/pdf" && bytes.byteLength > SAFE_RAW_PDF_BYTES) {
    return rasterizePdf(bytes);
  }

  return { base64: bytesToBase64(bytes), mimeType: file.type };
}
