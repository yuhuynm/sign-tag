"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { PDFDocument } from "pdf-lib";
import type {
  PDFDocumentProxy,
  PDFPageProxy,
} from "pdfjs-dist/types/src/display/api";

type EditorMode = "edit" | "review";
type DocumentType = "pdf" | "docx";
type PdfJsLib = typeof import("pdfjs-dist");

type SignatureAsset = {
  id: string;
  name: string;
  src: string;
  mimeType: string;
  bytes: Uint8Array;
};

type MockSignature = {
  id: string;
  name: string;
  src: string;
  mimeType: string;
};

type PageSize = {
  width: number;
  height: number;
};

type WordPageRect = PageSize & {
  left: number;
  top: number;
  virtual?: boolean;
};

type DocxMetadata = {
  pageCount: number | null;
  pageSize: PageSize | null;
};

type Placement = {
  id: string;
  signatureId: string;
  pageIndex: number;
  x: number;
  y: number;
  width: number;
  height: number;
};

type Interaction =
  | {
      kind: "move" | "resize";
      placementId: string;
      pointerId: number;
      startClientX: number;
      startClientY: number;
      startPlacement: Placement;
    }
  | null;

type DragAnchor = {
  offsetX: number;
  offsetY: number;
};

const MIN_SIGNATURE_RATIO = 0.035;
const DEFAULT_SIGNATURE_WIDTH_PX = 180;
const MAX_PDF_BYTES = 50 * 1024 * 1024;
const MAX_DOCX_BYTES = 30 * 1024 * 1024;
const MAX_PAGES = 200;
const MAX_PLACEMENTS = 300;
const DEFAULT_WORD_PAGE_SIZE: PageSize = { width: 816, height: 1056 };
const SIGNATURE_ASPECT_RATIO = 0.42;
const WORD_PAGE_GAP_PX = 24;
const MOCK_SIGNATURES: MockSignature[] = [
  {
    id: "signature-nguyen-van-a",
    name: "Nguyen Van A",
    mimeType: "image/svg+xml",
    src: "data:image/svg+xml,%3Csvg%20xmlns='http://www.w3.org/2000/svg'%20viewBox='0%200%20480%20180'%3E%3Crect%20width='480'%20height='180'%20fill='none'/%3E%3Cpath%20d='M42%20116c42-54%2080-72%20112-54%2029%2017%2015%2055-17%2052-33-3-32-48%204-70%2048-29%20103%2048%2064%2086-22%2022-42%2011-33-12%2013%2029%2051%2040%2092%2028%2026-8%2041-28%2041-28s-9%2040%2018%2039c35-1%2056-54%2056-54'%20fill='none'%20stroke='%230f766e'%20stroke-width='12'%20stroke-linecap='round'%20stroke-linejoin='round'/%3E%3Ctext%20x='58'%20y='156'%20font-family='Arial,Helvetica,sans-serif'%20font-size='24'%20font-weight='700'%20fill='%23172033'%3ENguyen%20Van%20A%3C/text%3E%3C/svg%3E",
  },
  {
    id: "signature-tran-thi-b",
    name: "Tran Thi B",
    mimeType: "image/svg+xml",
    src: "data:image/svg+xml,%3Csvg%20xmlns='http://www.w3.org/2000/svg'%20viewBox='0%200%20480%20180'%3E%3Crect%20width='480'%20height='180'%20fill='none'/%3E%3Cpath%20d='M46%2090c72-56%20121-54%20118-14-3%2037-77%2058-83%2026-6-31%2060-65%20111-27%2046%2034%205%2092-27%2068-22-17%2011-65%2045-51%2025%2010%2016%2054-15%2058-35%204-11-55%2031-50%2036%204%2049%2043%2085%2040%2029-2%2046-22%2062-45'%20fill='none'%20stroke='%231f2937'%20stroke-width='11'%20stroke-linecap='round'%20stroke-linejoin='round'/%3E%3Ctext%20x='58'%20y='156'%20font-family='Arial,Helvetica,sans-serif'%20font-size='24'%20font-weight='700'%20fill='%23172033'%3ETran%20Thi%20B%3C/text%3E%3C/svg%3E",
  },
  {
    id: "signature-company-seal",
    name: "Company Authorized",
    mimeType: "image/svg+xml",
    src: "data:image/svg+xml,%3Csvg%20xmlns='http://www.w3.org/2000/svg'%20viewBox='0%200%20480%20180'%3E%3Crect%20width='480'%20height='180'%20fill='none'/%3E%3Cellipse%20cx='153'%20cy='88'%20rx='102'%20ry='58'%20fill='none'%20stroke='%23b42318'%20stroke-width='9'/%3E%3Cellipse%20cx='153'%20cy='88'%20rx='73'%20ry='36'%20fill='none'%20stroke='%23b42318'%20stroke-width='5'/%3E%3Cpath%20d='M279%20113c24-39%2055-59%2092-47%2024%208%2030%2034%207%2046-21%2011-50-3-43-24%207-22%2050-28%2081%2010'%20fill='none'%20stroke='%23b42318'%20stroke-width='10'%20stroke-linecap='round'%20stroke-linejoin='round'/%3E%3Ctext%20x='88'%20y='95'%20font-family='Arial,Helvetica,sans-serif'%20font-size='22'%20font-weight='800'%20fill='%23b42318'%3EAPPROVED%3C/text%3E%3Ctext%20x='58'%20y='156'%20font-family='Arial,Helvetica,sans-serif'%20font-size='24'%20font-weight='700'%20fill='%23172033'%3ECompany%20Authorized%3C/text%3E%3C/svg%3E",
  },
];

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max);
}

function makeId(prefix: string) {
  return `${prefix}-${crypto.randomUUID()}`;
}

function copyBytes(bytes: Uint8Array) {
  return new Uint8Array(bytes);
}

function toArrayBuffer(bytes: Uint8Array) {
  const buffer = new ArrayBuffer(bytes.byteLength);

  new Uint8Array(buffer).set(bytes);
  return buffer;
}

async function svgDataUrlToPngBytes(dataUrl: string) {
  const image = new Image();

  image.decoding = "async";
  image.src = dataUrl;
  await image.decode();

  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");

  canvas.width = image.naturalWidth || 480;
  canvas.height = image.naturalHeight || 180;

  if (!context) {
    throw new Error("Unable to render mock signature.");
  }

  context.drawImage(image, 0, 0, canvas.width, canvas.height);
  return canvasToPngBytes(canvas);
}

function isPdf(bytes: Uint8Array) {
  return (
    bytes[0] === 0x25 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x44 &&
    bytes[3] === 0x46 &&
    bytes[4] === 0x2d
  );
}

function isZip(bytes: Uint8Array) {
  return bytes[0] === 0x50 && bytes[1] === 0x4b;
}

function twipsToPixels(twips: number) {
  return Math.round(twips / 15);
}

function getXmlAttribute(tag: string, attributeName: string) {
  const localName = attributeName.includes(":")
    ? attributeName.split(":").at(-1) ?? attributeName
    : attributeName;
  const escapeRegExp = (value: string) =>
    value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = tag.match(
    new RegExp(
      `\\s(?:${escapeRegExp(attributeName)}|(?:\\w+:)?${escapeRegExp(
        localName,
      )})="([^"]+)"`,
    ),
  );

  return match?.[1] ?? null;
}

async function isDocx(bytes: Uint8Array) {
  if (!isZip(bytes)) {
    return false;
  }

  try {
    const { default: JSZip } = await import("jszip");
    const zip = await JSZip.loadAsync(toArrayBuffer(bytes));

    return Boolean(zip.file("[Content_Types].xml") && zip.file("word/document.xml"));
  } catch {
    return false;
  }
}

async function getDocxMetadata(bytes: Uint8Array): Promise<DocxMetadata> {
  try {
    const { default: JSZip } = await import("jszip");
    const zip = await JSZip.loadAsync(toArrayBuffer(bytes));
    const appXml = await zip.file("docProps/app.xml")?.async("string");
    const documentXml = await zip.file("word/document.xml")?.async("string");
    const pageSizeTags = documentXml?.match(/<w:pgSz\b[^>]*\/?>/gi) ?? [];
    const pageSizeTag = pageSizeTags.at(-1)?.trim();
    const widthTwips = pageSizeTag
      ? Number.parseInt(getXmlAttribute(pageSizeTag, "w:w") ?? "", 10)
      : Number.NaN;
    const heightTwips = pageSizeTag
      ? Number.parseInt(getXmlAttribute(pageSizeTag, "w:h") ?? "", 10)
      : Number.NaN;
    const pageSize =
      Number.isFinite(widthTwips) &&
      Number.isFinite(heightTwips) &&
      widthTwips > 0 &&
      heightTwips > 0
        ? {
            height: twipsToPixels(heightTwips),
            width: twipsToPixels(widthTwips),
          }
        : null;

    if (!appXml) {
      return { pageCount: null, pageSize };
    }

    const match = appXml.match(/<Pages>(\d+)<\/Pages>/i);
    const pageCount = match ? Number.parseInt(match[1], 10) : Number.NaN;

    return {
      pageCount: Number.isFinite(pageCount) && pageCount > 0 ? pageCount : null,
      pageSize,
    };
  } catch {
    return { pageCount: null, pageSize: null };
  }
}

function sanitizeDownloadName(name: string) {
  return (
    name
      .replace(/\.(pdf|docx)$/i, "")
      .replace(/[^a-z0-9._-]+/gi, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 80) || "document"
  );
}

function downloadBytes({
  bytes,
  fileName,
  mimeType,
}: {
  bytes: Uint8Array;
  fileName: string;
  mimeType: string;
}) {
  const blob = new Blob([toArrayBuffer(bytes)], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function escapeXml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function pixelsToEmu(pixels: number) {
  return Math.round(pixels * 9525);
}

function getXmlMimeExtension(mimeType: string) {
  return mimeType === "image/png" ? "png" : "jpeg";
}

function getXmlMimeContentType(mimeType: string) {
  return mimeType === "image/png" ? "image/png" : "image/jpeg";
}

function ensureContentTypeDefault(
  contentTypesXml: XMLDocument,
  extension: string,
  contentType: string,
) {
  const existing = Array.from(contentTypesXml.getElementsByTagName("Default")).find(
    (node) => node.getAttribute("Extension") === extension,
  );

  if (existing) {
    return;
  }

  const defaultNode = contentTypesXml.createElementNS(
    "http://schemas.openxmlformats.org/package/2006/content-types",
    "Default",
  );

  defaultNode.setAttribute("Extension", extension);
  defaultNode.setAttribute("ContentType", contentType);
  contentTypesXml.documentElement.append(defaultNode);
}

function getPageAnchorMarkers(blocks: Element[]) {
  const pageStarts = [0];

  blocks.forEach((block, index) => {
    const hasManualBreak = block.getElementsByTagName("w:br").length > 0
      ? Array.from(block.getElementsByTagName("w:br")).some(
          (node) => node.getAttribute("w:type") === "page",
        )
      : false;
    const hasRenderedBreak =
      block.getElementsByTagName("w:lastRenderedPageBreak").length > 0;

    if ((hasManualBreak || hasRenderedBreak) && index + 1 < blocks.length) {
      pageStarts.push(index + 1);
    }
  });

  return pageStarts;
}

function ensureParagraphTarget(
  documentXml: XMLDocument,
  body: Element,
  block: Element | undefined,
) {
  if (!block) {
    const paragraph = documentXml.createElementNS(
      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "w:p",
    );

    body.insertBefore(paragraph, body.lastElementChild);
    return paragraph;
  }

  if (block.localName === "p") {
    return block;
  }

  const paragraph = documentXml.createElementNS(
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w:p",
  );

  body.insertBefore(paragraph, block);
  return paragraph;
}

function buildAnchoredDrawingXml({
  relationshipId,
  name,
  widthPx,
  heightPx,
  xPx,
  yPx,
  drawingId,
}: {
  relationshipId: string;
  name: string;
  widthPx: number;
  heightPx: number;
  xPx: number;
  yPx: number;
  drawingId: number;
}) {
  const widthEmu = pixelsToEmu(widthPx);
  const heightEmu = pixelsToEmu(heightPx);
  const xEmu = pixelsToEmu(xPx);
  const yEmu = pixelsToEmu(yPx);
  const escapedName = escapeXml(name);

  return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${xEmu}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${yEmu}</wp:posOffset></wp:positionV><wp:extent cx="${widthEmu}" cy="${heightEmu}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${drawingId}" name="${escapedName}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="${escapedName}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
}

function removeExternalStyles(container: HTMLElement) {
  container
    .querySelectorAll<HTMLLinkElement>('link[rel="stylesheet"]')
    .forEach((link) => link.remove());
  container.querySelectorAll<HTMLStyleElement>("style").forEach((style) => {
    style.textContent =
      style.textContent?.replace(/@import\s+url\([^)]*\)\s*;?/gi, "") ?? "";
  });
}

async function canvasToPngBytes(canvas: HTMLCanvasElement) {
  return new Promise<Uint8Array>((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error("Unable to encode rendered page."));
        return;
      }

      blob
        .arrayBuffer()
        .then((buffer) => resolve(new Uint8Array(buffer)))
        .catch(reject);
    }, "image/png");
  });
}

function createBlankPageCanvas(widthPx: number, heightPx: number) {
  const scale = 2;
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");

  canvas.width = Math.max(1, Math.round(widthPx * scale));
  canvas.height = Math.max(1, Math.round(heightPx * scale));

  if (!context) {
    throw new Error("Unable to create canvas.");
  }

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);

  return canvas;
}

function resizeCanvasToPageSize(
  sourceCanvas: HTMLCanvasElement,
  widthPx: number,
  heightPx: number,
) {
  const targetCanvas = createBlankPageCanvas(widthPx, heightPx);
  const context = targetCanvas.getContext("2d");

  if (!context) {
    throw new Error("Unable to normalize page.");
  }

  context.drawImage(sourceCanvas, 0, 0, targetCanvas.width, targetCanvas.height);

  return targetCanvas;
}

async function drawSignatureOnCanvas({
  canvas,
  placement,
  signature,
}: {
  canvas: HTMLCanvasElement;
  placement: Placement;
  signature: SignatureAsset;
}) {
  const context = canvas.getContext("2d");

  if (!context) {
    throw new Error("Unable to draw signature.");
  }

  const image = await createImageBitmap(
    new Blob([toArrayBuffer(signature.bytes)], { type: signature.mimeType }),
  );

  try {
    context.drawImage(
      image,
      placement.x * canvas.width,
      placement.y * canvas.height,
      placement.width * canvas.width,
      placement.height * canvas.height,
    );
  } finally {
    image.close();
  }
}

function arePageRectsEqual(current: WordPageRect[], next: WordPageRect[]) {
  if (current.length !== next.length) {
    return false;
  }

  return current.every((rect, index) => {
    const nextRect = next[index];

    return (
      Math.abs(rect.left - nextRect.left) < 0.5 &&
      Math.abs(rect.top - nextRect.top) < 0.5 &&
      Math.abs(rect.width - nextRect.width) < 0.5 &&
      Math.abs(rect.height - nextRect.height) < 0.5 &&
      Boolean(rect.virtual) === Boolean(nextRect.virtual)
    );
  });
}

function expandWordPageRects(
  renderedRects: WordPageRect[],
  expectedPageCount: number | null,
  expectedPageSize: PageSize | null,
) {
  const targetPageCount = expectedPageCount ?? renderedRects.length;
  const defaultSize = expectedPageSize ?? renderedRects[0] ?? DEFAULT_WORD_PAGE_SIZE;
  const nextRects = renderedRects.slice(0, targetPageCount).map((rect) => ({
    ...rect,
    height: expectedPageSize?.height ?? rect.height,
    width: expectedPageSize?.width ?? rect.width,
  }));
  const left = renderedRects[0]?.left ?? 0;
  let top =
    nextRects.length > 0
      ? Math.max(...nextRects.map((rect) => rect.top + rect.height)) +
        WORD_PAGE_GAP_PX
      : 0;

  while (nextRects.length < targetPageCount) {
    nextRects.push({
      height: defaultSize.height,
      left,
      top,
      virtual: true,
      width: defaultSize.width,
    });
    top += defaultSize.height + WORD_PAGE_GAP_PX;
  }

  return nextRects;
}

function getDefaultSignatureSize(pageWidth: number) {
  const width = Math.min(0.34, DEFAULT_SIGNATURE_WIDTH_PX / pageWidth);

  return {
    height: width * SIGNATURE_ASPECT_RATIO,
    width,
  };
}

function getDragAnchor(event: React.DragEvent) {
  const rawValue = event.dataTransfer.getData("application/pdf-sign-tag-anchor");

  if (!rawValue) {
    return null;
  }

  try {
    const parsed = JSON.parse(rawValue) as DragAnchor;

    return typeof parsed.offsetX === "number" &&
      typeof parsed.offsetY === "number"
      ? parsed
      : null;
  } catch {
    return null;
  }
}

function setSignatureDragImage({
  event,
  signature,
}: {
  event: React.DragEvent<HTMLElement>;
  signature: SignatureAsset;
}) {
  const dragImage = document.createElement("div");
  const image = document.createElement("img");
  const offsetX = DEFAULT_SIGNATURE_WIDTH_PX / 2;
  const offsetY = (DEFAULT_SIGNATURE_WIDTH_PX * SIGNATURE_ASPECT_RATIO) / 2;

  image.src = signature.src;
  image.alt = "";
  dragImage.className = "signature-drag-preview";
  dragImage.append(image);
  document.body.append(dragImage);
  event.dataTransfer.setDragImage(dragImage, offsetX, offsetY);
  event.dataTransfer.setData(
    "application/pdf-sign-tag-anchor",
    JSON.stringify({ offsetX, offsetY } satisfies DragAnchor),
  );
  window.requestAnimationFrame(() => dragImage.remove());
}

export default function PdfSignatureEditor() {
  const [mode, setMode] = useState<EditorMode>("edit");
  const [zoom, setZoom] = useState(1);
  const [pdfjs, setPdfjs] = useState<PdfJsLib | null>(null);
  const [documentType, setDocumentType] = useState<DocumentType | null>(null);
  const [documentBytes, setDocumentBytes] = useState<Uint8Array | null>(null);
  const [documentName, setDocumentName] = useState("");
  const [pdfDocument, setPdfDocument] = useState<PDFDocumentProxy | null>(null);
  const [docxMetadataPageCount, setDocxMetadataPageCount] = useState<
    number | null
  >(null);
  const [docxPageSize, setDocxPageSize] = useState<PageSize | null>(null);
  const [pageSizes, setPageSizes] = useState<PageSize[]>([]);
  const [signatures, setSignatures] = useState<SignatureAsset[]>([]);
  const [placements, setPlacements] = useState<Placement[]>([]);
  const [selectedPlacementId, setSelectedPlacementId] = useState<string | null>(
    null,
  );
  const [interaction, setInteraction] = useState<Interaction>(null);
  const [status, setStatus] = useState("Upload a PDF or DOCX to begin.");
  const [isExporting, setIsExporting] = useState(false);

  const pageSizesRef = useRef<PageSize[]>([]);
  const signaturesRef = useRef<SignatureAsset[]>([]);
  const pdfDocumentRef = useRef<PDFDocumentProxy | null>(null);
  const wordStageRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    let mounted = true;

    import("pdfjs-dist").then((loadedPdfjs) => {
      loadedPdfjs.GlobalWorkerOptions.workerSrc = new URL(
        "pdfjs-dist/build/pdf.worker.mjs",
        import.meta.url,
      ).toString();

      if (mounted) {
        setPdfjs(loadedPdfjs);
      }
    });

    return () => {
      mounted = false;
    };
  }, []);

  useEffect(() => {
    let mounted = true;

    Promise.all(
      MOCK_SIGNATURES.map(async (signature) => ({
        id: signature.id,
        name: signature.name,
        src: signature.src,
        mimeType: "image/png",
        bytes: await svgDataUrlToPngBytes(signature.src),
      })),
    )
      .then((mockSignatures) => {
        if (mounted) {
          setSignatures(mockSignatures);
        }
      })
      .catch(() => {
        if (mounted) {
          setStatus("Unable to load mock signatures.");
        }
      });

    return () => {
      mounted = false;
    };
  }, []);

  useEffect(() => {
    pageSizesRef.current = pageSizes;
  }, [pageSizes]);

  useEffect(() => {
    signaturesRef.current = signatures;
  }, [signatures]);

  useEffect(() => {
    pdfDocumentRef.current = pdfDocument;
  }, [pdfDocument]);

  useEffect(
    () => () => {
      pdfDocumentRef.current?.destroy().catch(() => undefined);
    },
    [],
  );

  const selectedPlacement = useMemo(
    () =>
      placements.find((placement) => placement.id === selectedPlacementId) ??
      null,
    [placements, selectedPlacementId],
  );

  const clearDocument = async () => {
    await pdfDocumentRef.current?.destroy();
    setPdfDocument(null);
    setDocumentType(null);
    setDocumentBytes(null);
    setDocumentName("");
    setDocxMetadataPageCount(null);
    setDocxPageSize(null);
    setPageSizes([]);
    setPlacements([]);
    setSelectedPlacementId(null);
  };

  const handleDocumentUpload = async (
    event: React.ChangeEvent<HTMLInputElement>,
  ) => {
    const file = event.target.files?.[0];

    if (!file) {
      return;
    }

    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const lowerName = file.name.toLowerCase();

      if (isPdf(bytes) || lowerName.endsWith(".pdf")) {
        if (!pdfjs) {
          setStatus("PDF renderer is still loading.");
          return;
        }

        if (file.size > MAX_PDF_BYTES) {
          setStatus("PDF rejected: maximum file size is 50 MB.");
          return;
        }

        if (!isPdf(bytes)) {
          setStatus("PDF rejected: file signature does not match a PDF.");
          return;
        }

        const loadingTask = pdfjs.getDocument({
          data: copyBytes(bytes),
          disableAutoFetch: true,
          disableStream: true,
          enableXfa: false,
          isEvalSupported: false,
          maxImageSize: 24_000_000,
          stopAtErrors: true,
        });
        const loadedPdf = await loadingTask.promise;

        if (loadedPdf.numPages > MAX_PAGES) {
          await loadedPdf.destroy();
          setStatus(`PDF rejected: maximum page count is ${MAX_PAGES}.`);
          return;
        }

        await clearDocument();

        setDocumentType("pdf");
        setDocumentBytes(bytes);
        setDocumentName(file.name);
        setPdfDocument(loadedPdf);
        setPageSizes(Array.from({ length: loadedPdf.numPages }));
        setStatus(`${file.name} loaded with ${loadedPdf.numPages} page(s).`);
        return;
      }

      if (!lowerName.endsWith(".docx")) {
        setStatus("File rejected: upload a PDF or DOCX file.");
        return;
      }

      if (file.size > MAX_DOCX_BYTES) {
        setStatus("DOCX rejected: maximum file size is 30 MB.");
        return;
      }

      if (!(await isDocx(bytes))) {
        setStatus("DOCX rejected: file structure does not match Word DOCX.");
        return;
      }
      await clearDocument();

      const docxMetadata = await getDocxMetadata(bytes);

      setDocumentType("docx");
      setDocumentBytes(bytes);
      setDocumentName(file.name);
      setDocxMetadataPageCount(docxMetadata.pageCount);
      setDocxPageSize(docxMetadata.pageSize);
      setPageSizes([]);
      setStatus("DOCX loaded in browser preview. Export will generate a PDF.");
    } catch (error) {
      setStatus(
        error instanceof Error
          ? `File rejected: ${error.message}`
          : "File rejected.",
      );
    } finally {
      event.target.value = "";
    }
  };

  const updatePageSize = useCallback((pageIndex: number, size: PageSize) => {
    setPageSizes((current) => {
      const next = [...current];
      const previous = next[pageIndex];

      if (
        previous &&
        Math.abs(previous.width - size.width) < 0.5 &&
        Math.abs(previous.height - size.height) < 0.5
      ) {
        return current;
      }

      next[pageIndex] = size;
      return next;
    });
  }, []);

  const addPlacement = (
    signatureId: string,
    pageIndex: number,
    pointX: number,
    pointY: number,
  ) => {
    if (placements.length >= MAX_PLACEMENTS) {
      setStatus(`Placement limit reached: maximum ${MAX_PLACEMENTS}.`);
      return;
    }

    const pageSize = pageSizesRef.current[pageIndex];
    const signature = signatures.find((asset) => asset.id === signatureId);

    if (!pageSize || !signature) {
      return;
    }

    const { width: defaultWidth, height: defaultHeight } =
      getDefaultSignatureSize(pageSize.width * zoom);
    const x = clamp(pointX / (pageSize.width * zoom), 0, 1 - defaultWidth);
    const y = clamp(pointY / (pageSize.height * zoom), 0, 1 - defaultHeight);
    const placement: Placement = {
      id: makeId("placement"),
      signatureId,
      pageIndex,
      x,
      y,
      width: defaultWidth,
      height: defaultHeight,
    };

    setPlacements((current) => [...current, placement]);
    setSelectedPlacementId(placement.id);
  };

  const handlePageDrop = (
    event: React.DragEvent<HTMLDivElement>,
    pageIndex: number,
  ) => {
    event.preventDefault();

    if (mode !== "edit") {
      return;
    }

    const signatureId = event.dataTransfer.getData("application/pdf-sign-tag");

    if (!signatureId) {
      return;
    }

    const bounds = event.currentTarget.getBoundingClientRect();
    addPlacement(
      signatureId,
      pageIndex,
      event.clientX - bounds.left - (getDragAnchor(event)?.offsetX ?? 0),
      event.clientY - bounds.top - (getDragAnchor(event)?.offsetY ?? 0),
    );
  };

  const handleWordPageDrop = (
    event: React.DragEvent<HTMLDivElement>,
    pageIndex: number,
  ) => {
    event.preventDefault();
    event.stopPropagation();

    if (mode !== "edit") {
      return;
    }

    if (placements.length >= MAX_PLACEMENTS) {
      setStatus(`Placement limit reached: maximum ${MAX_PLACEMENTS}.`);
      return;
    }

    const signatureId = event.dataTransfer.getData("application/pdf-sign-tag");
    const signature = signatures.find((asset) => asset.id === signatureId);

    if (!signature) {
      return;
    }

    const anchor = getDragAnchor(event);
    const bounds = event.currentTarget.getBoundingClientRect();
    const { width: defaultWidth, height: defaultHeight } =
      getDefaultSignatureSize(bounds.width);
    const placement: Placement = {
      id: makeId("placement"),
      signatureId,
      pageIndex,
      x: clamp(
        (event.clientX - bounds.left - (anchor?.offsetX ?? 0)) / bounds.width,
        0,
        1 - defaultWidth,
      ),
      y: clamp(
        (event.clientY - bounds.top - (anchor?.offsetY ?? 0)) / bounds.height,
        0,
        1 - defaultHeight,
      ),
      width: defaultWidth,
      height: defaultHeight,
    };

    setPlacements((current) => [...current, placement]);
    setSelectedPlacementId(placement.id);
  };

  const startInteraction = (
    event: React.PointerEvent<HTMLElement>,
    placement: Placement,
    kind: "move" | "resize",
  ) => {
    if (mode !== "edit") {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.setPointerCapture(event.pointerId);
    setSelectedPlacementId(placement.id);
    setInteraction({
      kind,
      placementId: placement.id,
      pointerId: event.pointerId,
      startClientX: event.clientX,
      startClientY: event.clientY,
      startPlacement: placement,
    });
  };

  useEffect(() => {
    if (!interaction) {
      return;
    }

    const handlePointerMove = (event: PointerEvent) => {
      if (event.pointerId !== interaction.pointerId) {
        return;
      }

      const pageSize = pageSizesRef.current[interaction.startPlacement.pageIndex];

      if (!pageSize) {
        return;
      }

      const deltaX = event.clientX - interaction.startClientX;
      const deltaY = event.clientY - interaction.startClientY;
      const ratioDeltaX = deltaX / (pageSize.width * zoom);
      const ratioDeltaY = deltaY / (pageSize.height * zoom);

      setPlacements((current) =>
        current.map((placement) => {
          if (placement.id !== interaction.placementId) {
            return placement;
          }

          if (interaction.kind === "move") {
            return {
              ...placement,
              x: clamp(
                interaction.startPlacement.x + ratioDeltaX,
                0,
                1 - placement.width,
              ),
              y: clamp(
                interaction.startPlacement.y + ratioDeltaY,
                0,
                1 - placement.height,
              ),
            };
          }

          const width = clamp(
            interaction.startPlacement.width + ratioDeltaX,
            MIN_SIGNATURE_RATIO,
            1 - interaction.startPlacement.x,
          );
          const height = clamp(
            interaction.startPlacement.height + ratioDeltaY,
            MIN_SIGNATURE_RATIO,
            1 - interaction.startPlacement.y,
          );

          return {
            ...placement,
            width,
            height,
          };
        }),
      );
    };

    const handlePointerUp = (event: PointerEvent) => {
      if (event.pointerId === interaction.pointerId) {
        setInteraction(null);
      }
    };

    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", handlePointerUp);
    window.addEventListener("pointercancel", handlePointerUp);

    return () => {
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", handlePointerUp);
      window.removeEventListener("pointercancel", handlePointerUp);
    };
  }, [interaction, zoom]);

  const removeSelectedPlacement = () => {
    if (!selectedPlacementId || mode !== "edit") {
      return;
    }

    setPlacements((current) =>
      current.filter((placement) => placement.id !== selectedPlacementId),
    );
    setSelectedPlacementId(null);
  };

  const exportPdfDocument = async () => {
    if (!documentBytes || placements.length === 0) {
      setStatus("Add at least one signature before exporting.");
      return;
    }

    setIsExporting(true);
    setStatus("Exporting signed PDF...");

    try {
      const outputPdf = await PDFDocument.load(copyBytes(documentBytes));
      const pages = outputPdf.getPages();
      const embeddedImages = new Map<string, Awaited<ReturnType<typeof outputPdf.embedPng>>>();

      for (const placement of placements) {
        const signature = signatures.find(
          (asset) => asset.id === placement.signatureId,
        );
        const page = pages[placement.pageIndex];

        if (!signature || !page) {
          continue;
        }

        let image = embeddedImages.get(signature.id);

        if (!image) {
          image =
            signature.mimeType === "image/png"
              ? await outputPdf.embedPng(signature.bytes)
              : await outputPdf.embedJpg(signature.bytes);
          embeddedImages.set(signature.id, image);
        }

        const pageWidth = page.getWidth();
        const pageHeight = page.getHeight();
        const width = placement.width * pageWidth;
        const height = placement.height * pageHeight;

        page.drawImage(image, {
          x: placement.x * pageWidth,
          y: pageHeight - placement.y * pageHeight - height,
          width,
          height,
        });
      }

      const baseName = sanitizeDownloadName(documentName);

      downloadBytes({
        bytes: await outputPdf.save(),
        fileName: `${baseName}-signed.pdf`,
        mimeType: "application/pdf",
      });
      setStatus("Signed PDF exported.");
    } catch (error) {
      setStatus(
        error instanceof Error
          ? `Export failed: ${error.message}`
          : "Export failed.",
      );
    } finally {
      setIsExporting(false);
    }
  };

  const exportWordPreviewAsPdf = async () => {
    const stage = wordStageRef.current;

    if (!stage || placements.length === 0) {
      setStatus("Add at least one signature before exporting.");
      return;
    }

    setIsExporting(true);
    setStatus("Exporting PDF from DOCX preview...");

    const previousTransform = stage.style.transform;

    try {
      const html2canvas = (await import("html2canvas")).default;
      const pageElements = Array.from(
        stage.querySelectorAll<HTMLElement>(".docx"),
      );

      if (pageElements.length === 0) {
        throw new Error("No rendered DOCX pages found.");
      }

      stage.style.transform = "none";
      await new Promise((resolve) => requestAnimationFrame(resolve));

      const outputPdf = await PDFDocument.create();
      const targetPageCount = docxMetadataPageCount ?? pageElements.length;
      const fallbackSize = {
        height:
          docxPageSize?.height ??
          pageElements[0]?.offsetHeight ??
          DEFAULT_WORD_PAGE_SIZE.height,
        width:
          docxPageSize?.width ??
          pageElements[0]?.offsetWidth ??
          DEFAULT_WORD_PAGE_SIZE.width,
      };

      for (let pageIndex = 0; pageIndex < targetPageCount; pageIndex += 1) {
        const pageElement = pageElements[pageIndex];
        const pageWidth =
          docxPageSize?.width ?? pageElement?.offsetWidth ?? fallbackSize.width;
        const pageHeight =
          docxPageSize?.height ?? pageElement?.offsetHeight ?? fallbackSize.height;
        const canvas = pageElement
          ? resizeCanvasToPageSize(
              await html2canvas(pageElement, {
                backgroundColor: "#ffffff",
                scale: 2,
                useCORS: false,
              }),
              pageWidth,
              pageHeight,
            )
          : createBlankPageCanvas(pageWidth, pageHeight);

        for (const placement of placements.filter(
          (item) => item.pageIndex === pageIndex,
        )) {
          const signature = signatures.find(
            (asset) => asset.id === placement.signatureId,
          );

          if (!signature) {
            continue;
          }

          await drawSignatureOnCanvas({
            canvas,
            placement,
            signature,
          });
        }

        const imageBytes = await canvasToPngBytes(canvas);
        const image = await outputPdf.embedPng(imageBytes);
        const page = outputPdf.addPage([pageWidth, pageHeight]);

        page.drawImage(image, {
          x: 0,
          y: 0,
          width: pageWidth,
          height: pageHeight,
        });
      }

      const baseName = sanitizeDownloadName(documentName);

      downloadBytes({
        bytes: await outputPdf.save(),
        fileName: `${baseName}-signed.pdf`,
        mimeType: "application/pdf",
      });
      setStatus("Signed PDF exported.");
    } catch (error) {
      setStatus(
        error instanceof Error
          ? `Export failed: ${error.message}`
          : "Export failed.",
      );
    } finally {
      stage.style.transform = previousTransform;
      setIsExporting(false);
    }
  };

  const exportWordDocument = async () => {
    if (!documentBytes || placements.length === 0) {
      setStatus("Add at least one signature before exporting.");
      return;
    }

    setIsExporting(true);
    setStatus("Exporting signed Word document...");

    try {
      const { default: JSZip } = await import("jszip");
      const zip = await JSZip.loadAsync(toArrayBuffer(documentBytes));
      const documentXmlSource = await zip.file("word/document.xml")?.async("string");
      const relationshipsSource = await zip
        .file("word/_rels/document.xml.rels")
        ?.async("string");
      const contentTypesSource = await zip.file("[Content_Types].xml")?.async("string");

      if (!documentXmlSource || !relationshipsSource || !contentTypesSource) {
        throw new Error("DOCX package is missing required XML parts.");
      }

      const parser = new DOMParser();
      const serializer = new XMLSerializer();
      const documentXml = parser.parseFromString(documentXmlSource, "application/xml");
      const relationshipsXml = parser.parseFromString(
        relationshipsSource,
        "application/xml",
      );
      const contentTypesXml = parser.parseFromString(
        contentTypesSource,
        "application/xml",
      );
      const body = documentXml.getElementsByTagName("w:body")[0];
      const relationshipsRoot = relationshipsXml.documentElement;

      if (!body || !relationshipsRoot) {
        throw new Error("Unable to parse DOCX structure.");
      }

      const blocks = Array.from(body.children).filter(
        (node) => node.localName !== "sectPr",
      );
      const pageStarts = getPageAnchorMarkers(blocks);
      const pageTargets = pageStarts.map((index) => blocks[index]);
      const relationshipIds = Array.from(
        relationshipsRoot.getElementsByTagName("Relationship"),
      )
        .map((node) => node.getAttribute("Id") ?? "")
        .map((value) =>
          value.startsWith("rId") ? Number.parseInt(value.slice(3), 10) : NaN,
        )
        .filter((value) => Number.isFinite(value));
      let nextRelationshipId =
        (relationshipIds.length > 0 ? Math.max(...relationshipIds) : 0) + 1;
      let nextDrawingId = 1_000;
      const fallbackPageSize =
        docxPageSize ?? pageSizes[0] ?? DEFAULT_WORD_PAGE_SIZE;

      for (const placement of placements) {
        const signature = signatures.find(
          (asset) => asset.id === placement.signatureId,
        );

        if (!signature) {
          continue;
        }

        const targetParagraph = ensureParagraphTarget(
          documentXml,
          body,
          pageTargets[placement.pageIndex],
        );
        const relationshipId = `rId${nextRelationshipId}`;
        const extension = getXmlMimeExtension(signature.mimeType);
        const contentType = getXmlMimeContentType(signature.mimeType);
        const mediaName = `signature-${placement.id}.${extension}`;
        const pageSize = pageSizes[placement.pageIndex] ?? fallbackPageSize;
        const widthPx = placement.width * pageSize.width;
        const heightPx = placement.height * pageSize.height;
        const xPx = placement.x * pageSize.width;
        const yPx = placement.y * pageSize.height;
        const relationshipNode = relationshipsXml.createElementNS(
          "http://schemas.openxmlformats.org/package/2006/relationships",
          "Relationship",
        );

        relationshipNode.setAttribute("Id", relationshipId);
        relationshipNode.setAttribute(
          "Type",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        );
        relationshipNode.setAttribute("Target", `media/${mediaName}`);
        relationshipsRoot.append(relationshipNode);
        ensureContentTypeDefault(contentTypesXml, extension, contentType);
        zip.file(`word/media/${mediaName}`, toArrayBuffer(signature.bytes));

        const wrapperXml = parser.parseFromString(
          `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">${buildAnchoredDrawingXml({
            relationshipId,
            name: signature.name,
            widthPx,
            heightPx,
            xPx,
            yPx,
            drawingId: nextDrawingId,
          })}</root>`,
          "application/xml",
        );
        const importedNode = wrapperXml.documentElement.firstElementChild;

        if (!importedNode) {
          throw new Error("Unable to build signature drawing.");
        }

        targetParagraph.append(documentXml.importNode(importedNode, true));
        nextRelationshipId += 1;
        nextDrawingId += 1;
      }

      zip.file("word/document.xml", serializer.serializeToString(documentXml));
      zip.file(
        "word/_rels/document.xml.rels",
        serializer.serializeToString(relationshipsXml),
      );
      zip.file(
        "[Content_Types].xml",
        serializer.serializeToString(contentTypesXml),
      );

      const baseName = sanitizeDownloadName(documentName);
      const savedBytes = new Uint8Array(
        await zip.generateAsync({
          compression: "DEFLATE",
          type: "arraybuffer",
        }),
      );
      downloadBytes({
        bytes: savedBytes,
        fileName: `${baseName}-signed.docx`,
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      setStatus("Signed Word document exported.");
    } catch (error) {
      setStatus(
        error instanceof Error
          ? `Export failed: ${error.message}`
          : "Export failed.",
      );
    } finally {
      setIsExporting(false);
    }
  };

  const pageCount =
    documentType === "docx"
      ? docxMetadataPageCount ?? pageSizes.length
      : pdfDocument?.numPages ?? pageSizes.length;
  const handleWordRendered = useCallback(
    (count: number) => {
      setStatus(
        docxMetadataPageCount && docxMetadataPageCount > count
          ? `${documentName} rendered ${count} page(s); export will preserve ${docxMetadataPageCount} page(s).`
          : `${documentName} rendered with ${count} page(s).`,
      );
    },
    [docxMetadataPageCount, documentName],
  );

  return (
    <main className="editor-shell">
      <header className="editor-toolbar">
        <div className="brand-block">
          <span className="brand-mark">PS</span>
          <div>
            <h1>PDF Sign Tag</h1>
            <p>{documentName || "No document loaded"}</p>
          </div>
        </div>

        <div className="toolbar-actions">
          <label className="file-control">
            <input
              accept="application/pdf,.pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,.docx"
              type="file"
              onChange={handleDocumentUpload}
            />
            PDF / DOCX
          </label>
          <div className="segmented-control" aria-label="Editor mode">
            <button
              aria-pressed={mode === "edit"}
              onClick={() => setMode("edit")}
              type="button"
            >
              Edit
            </button>
            <button
              aria-pressed={mode === "review"}
              onClick={() => {
                setMode("review");
                setInteraction(null);
                setSelectedPlacementId(null);
              }}
              type="button"
            >
              Review
            </button>
          </div>
          <div className="zoom-control">
            <button
              aria-label="Zoom out"
              onClick={() => setZoom((value) => clamp(value - 0.1, 0.6, 1.8))}
              type="button"
            >
              -
            </button>
            <span>{Math.round(zoom * 100)}%</span>
            <button
              aria-label="Zoom in"
              onClick={() => setZoom((value) => clamp(value + 0.1, 0.6, 1.8))}
              type="button"
            >
              +
            </button>
          </div>
          {documentType === "docx" ? (
            <>
              <button
                className="secondary-action"
                disabled={!documentBytes || placements.length === 0 || isExporting}
                onClick={exportWordPreviewAsPdf}
                type="button"
              >
                {isExporting ? "Exporting" : "Export PDF"}
              </button>
              <button
                className="primary-action"
                disabled={!documentBytes || placements.length === 0 || isExporting}
                onClick={exportWordDocument}
                type="button"
              >
                {isExporting ? "Exporting" : "Export Word"}
              </button>
            </>
          ) : (
            <button
              className="primary-action"
              disabled={!documentBytes || placements.length === 0 || isExporting}
              onClick={exportPdfDocument}
              type="button"
            >
              {isExporting ? "Exporting" : "Export PDF"}
            </button>
          )}
        </div>
      </header>

      <section className="editor-body">
        <aside className="signature-sidebar">
          <div className="panel-heading">
            <h2>Signatures</h2>
            <span>{signatures.length}</span>
          </div>

          <div className="signature-list">
            {signatures.length === 0 ? (
              <p className="empty-copy">Loading signatures.</p>
            ) : (
              signatures.map((signature) => (
                <button
                  className="signature-tile"
                  draggable={mode === "edit"}
                  key={signature.id}
                  onDragStart={(event) => {
                    event.dataTransfer.setData(
                      "application/pdf-sign-tag",
                      signature.id,
                    );
                    event.dataTransfer.effectAllowed = "copy";
                    setSignatureDragImage({ event, signature });
                  }}
                  type="button"
                >
                  {/* eslint-disable-next-line @next/next/no-img-element */}
                  <img alt={signature.name} src={signature.src} />
                  <span>{signature.name}</span>
                </button>
              ))
            )}
          </div>

          <div className="inspector">
            <h2>Selection</h2>
            {selectedPlacement ? (
              <>
                <dl>
                  <div>
                    <dt>Page</dt>
                    <dd>
                      {selectedPlacement.pageIndex + 1}
                      {pageCount ? ` / ${pageCount}` : ""}
                    </dd>
                  </div>
                  <div>
                    <dt>Width</dt>
                    <dd>{Math.round(selectedPlacement.width * 100)}%</dd>
                  </div>
                </dl>
                <button
                  disabled={mode !== "edit"}
                  onClick={removeSelectedPlacement}
                  type="button"
                >
                  Delete selected
                </button>
              </>
            ) : (
              <p className="empty-copy">Select a placed signature.</p>
            )}
          </div>
        </aside>

        <div className="document-workspace">
          <div className="workspace-status" role="status">
            <span>{mode === "edit" ? "Edit mode" : "Review mode"}</span>
            <p>{status}</p>
          </div>

          {!documentType ? (
            <div className="drop-empty-state">
              <h2>Open a PDF or Word file to start placing signatures</h2>
              <p>
                Upload a PDF or DOCX, then drag a signature onto any page.
              </p>
            </div>
          ) : documentType === "pdf" && pdfDocument ? (
            <div className="page-stack" onClick={() => setSelectedPlacementId(null)}>
              {Array.from({ length: pdfDocument.numPages }, (_, index) => (
                <PdfPageSurface
                  key={index}
                  mode={mode}
                  onDrop={handlePageDrop}
                  onPageSize={updatePageSize}
                  pageIndex={index}
                  pdfDocument={pdfDocument}
                  placements={placements}
                  selectedPlacementId={selectedPlacementId}
                  signatures={signatures}
                  startInteraction={startInteraction}
                  zoom={zoom}
                />
              ))}
            </div>
          ) : documentType === "docx" && documentBytes ? (
            <WordDocumentSurface
              bytes={documentBytes}
              expectedPageCount={docxMetadataPageCount}
              expectedPageSize={docxPageSize}
              mode={mode}
              onDrop={handleWordPageDrop}
              onPageSize={updatePageSize}
              onRendered={handleWordRendered}
              placements={placements}
              selectedPlacementId={selectedPlacementId}
              signatures={signatures}
              stageRef={wordStageRef}
              startInteraction={startInteraction}
              zoom={zoom}
            />
          ) : (
            <div className="drop-empty-state">
              <h2>Unable to render this document</h2>
              <p>Try another PDF or DOCX file.</p>
            </div>
          )}
        </div>
      </section>
    </main>
  );
}

function WordDocumentSurface({
  bytes,
  expectedPageCount,
  expectedPageSize,
  mode,
  onDrop,
  onPageSize,
  onRendered,
  placements,
  selectedPlacementId,
  signatures,
  stageRef,
  startInteraction,
  zoom,
}: {
  bytes: Uint8Array;
  expectedPageCount: number | null;
  expectedPageSize: PageSize | null;
  mode: EditorMode;
  onDrop: (event: React.DragEvent<HTMLDivElement>, pageIndex: number) => void;
  onPageSize: (pageIndex: number, size: PageSize) => void;
  onRendered: (pageCount: number) => void;
  placements: Placement[];
  selectedPlacementId: string | null;
  signatures: SignatureAsset[];
  stageRef: React.RefObject<HTMLDivElement | null>;
  startInteraction: (
    event: React.PointerEvent<HTMLElement>,
    placement: Placement,
    kind: "move" | "resize",
  ) => void;
  zoom: number;
}) {
  const [pageRects, setPageRects] = useState<WordPageRect[]>([]);
  const [stageSize, setStageSize] = useState<PageSize>({ width: 0, height: 0 });
  const renderedPageCountRef = useRef(0);

  const measurePages = useCallback(() => {
    const stage = stageRef.current;

    if (!stage) {
      return;
    }

    const pages = Array.from(stage.querySelectorAll<HTMLElement>(".docx"));
    const nextRects = pages.map((page) => ({
      left: page.offsetLeft,
      top: page.offsetTop,
      width: page.offsetWidth,
      height: page.offsetHeight,
    }));
    const expandedRects = expandWordPageRects(
      nextRects,
      expectedPageCount,
      expectedPageSize,
    );
    const virtualWidth = expandedRects.reduce(
      (maximum, rect) => Math.max(maximum, rect.left + rect.width),
      0,
    );
    const virtualHeight = expandedRects.reduce(
      (maximum, rect) => Math.max(maximum, rect.top + rect.height),
      0,
    );

    setPageRects((current) =>
      arePageRectsEqual(current, expandedRects) ? current : expandedRects,
    );
    setStageSize((current) => {
      const nextSize = {
        width: expectedPageCount ? virtualWidth : Math.max(stage.scrollWidth, virtualWidth),
        height: expectedPageCount
          ? virtualHeight
          : Math.max(stage.scrollHeight, virtualHeight),
      };

      if (
        Math.abs(current.width - nextSize.width) < 0.5 &&
        Math.abs(current.height - nextSize.height) < 0.5
      ) {
        return current;
      }

      return nextSize;
    });
    expandedRects.forEach((rect, index) =>
      onPageSize(index, { width: rect.width, height: rect.height }),
    );
    if (renderedPageCountRef.current !== nextRects.length) {
      renderedPageCountRef.current = nextRects.length;
      onRendered(nextRects.length);
    }
  }, [expectedPageCount, expectedPageSize, onPageSize, onRendered, stageRef]);

  useEffect(() => {
    let cancelled = false;
    const stage = stageRef.current;

    if (!stage) {
      return;
    }

    stage.replaceChildren();

    import("docx-preview")
      .then(({ renderAsync }) =>
        renderAsync(new Blob([toArrayBuffer(bytes)]), stage, undefined, {
          breakPages: true,
          className: "docx",
          experimental: false,
          ignoreFonts: true,
          ignoreLastRenderedPageBreak: false,
          inWrapper: true,
          renderComments: false,
          renderEndnotes: true,
          renderFooters: true,
          renderFootnotes: true,
          renderHeaders: true,
          renderChanges: false,
          renderAltChunks: false,
          trimXmlDeclaration: true,
          useBase64URL: true,
        }),
      )
      .then(() => {
        if (!cancelled) {
          removeExternalStyles(stage);
          requestAnimationFrame(measurePages);
        }
      })
      .catch((error) => {
        if (!cancelled) {
          stage.textContent =
            error instanceof Error
              ? `Unable to render DOCX: ${error.message}`
              : "Unable to render DOCX.";
        }
      });

    return () => {
      cancelled = true;
      stage.replaceChildren();
    };
  }, [bytes, measurePages, stageRef]);

  useEffect(() => {
    const stage = stageRef.current;

    if (!stage) {
      return;
    }

    const resizeObserver = new ResizeObserver(() => measurePages());

    resizeObserver.observe(stage);

    return () => resizeObserver.disconnect();
  }, [measurePages, stageRef]);

  return (
    <div className="page-stack" onClick={() => undefined}>
      <div
        className="word-document-outer"
        style={{
          height: stageSize.height * zoom,
          width: stageSize.width * zoom,
        }}
      >
        <div
          className="word-document-stage"
          ref={stageRef}
          style={{
            transform: `scale(${zoom})`,
          }}
        />
        <div
          className="word-signature-overlay"
          style={{
            height: stageSize.height * zoom,
            width: stageSize.width * zoom,
          }}
        >
          {pageRects.map((rect, pageIndex) => (
            <div
              className={`word-page-drop-zone${
                rect.virtual ? " is-virtual" : ""
              }`}
              key={pageIndex}
              onDragOver={(event) => {
                if (mode === "edit") {
                  event.preventDefault();
                  event.dataTransfer.dropEffect = "copy";
                }
              }}
              onDrop={(event) => onDrop(event, pageIndex)}
              style={{
                height: rect.height * zoom,
                left: rect.left * zoom,
                top: rect.top * zoom,
                width: rect.width * zoom,
              }}
            >
              {placements
                .filter((placement) => placement.pageIndex === pageIndex)
                .map((placement) => {
                  const signature = signatures.find(
                    (asset) => asset.id === placement.signatureId,
                  );

                  if (!signature) {
                    return null;
                  }

                  const isSelected = selectedPlacementId === placement.id;

                  return (
                    <button
                      className={`placed-signature${
                        isSelected ? " is-selected" : ""
                      }`}
                      key={placement.id}
                      onClick={(event) => event.stopPropagation()}
                      onPointerDown={(event) =>
                        startInteraction(event, placement, "move")
                      }
                      style={{
                        height: `${placement.height * 100}%`,
                        left: `${placement.x * 100}%`,
                        top: `${placement.y * 100}%`,
                        width: `${placement.width * 100}%`,
                      }}
                      type="button"
                    >
                      {/* eslint-disable-next-line @next/next/no-img-element */}
                      <img alt="" draggable={false} src={signature.src} />
                      {mode === "edit" && isSelected ? (
                        <span
                          aria-hidden="true"
                          className="resize-handle"
                          onPointerDown={(event) =>
                            startInteraction(event, placement, "resize")
                          }
                        />
                      ) : null}
                    </button>
                  );
                })}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function PdfPageSurface({
  mode,
  onDrop,
  onPageSize,
  pageIndex,
  pdfDocument,
  placements,
  selectedPlacementId,
  signatures,
  startInteraction,
  zoom,
}: {
  mode: EditorMode;
  onDrop: (event: React.DragEvent<HTMLDivElement>, pageIndex: number) => void;
  onPageSize: (pageIndex: number, size: PageSize) => void;
  pageIndex: number;
  pdfDocument: PDFDocumentProxy;
  placements: Placement[];
  selectedPlacementId: string | null;
  signatures: SignatureAsset[];
  startInteraction: (
    event: React.PointerEvent<HTMLElement>,
    placement: Placement,
    kind: "move" | "resize",
  ) => void;
  zoom: number;
}) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [page, setPage] = useState<PDFPageProxy | null>(null);
  const [pageSize, setPageSize] = useState<PageSize | null>(null);

  useEffect(() => {
    let cancelled = false;

    pdfDocument.getPage(pageIndex + 1).then((loadedPage) => {
      if (cancelled) {
        return;
      }

      const viewport = loadedPage.getViewport({ scale: 1 });

      setPage(loadedPage);
      setPageSize({ width: viewport.width, height: viewport.height });
      onPageSize(pageIndex, { width: viewport.width, height: viewport.height });
    });

    return () => {
      cancelled = true;
    };
  }, [onPageSize, pageIndex, pdfDocument]);

  useEffect(() => {
    if (!page || !canvasRef.current) {
      return;
    }

    const canvas = canvasRef.current;
    const context = canvas.getContext("2d");

    if (!context) {
      return;
    }

    const viewport = page.getViewport({ scale: zoom });
    const outputScale = window.devicePixelRatio || 1;

    canvas.width = Math.floor(viewport.width * outputScale);
    canvas.height = Math.floor(viewport.height * outputScale);
    canvas.style.width = `${viewport.width}px`;
    canvas.style.height = `${viewport.height}px`;

    const renderTask = page.render({
      canvas,
      canvasContext: context,
      transform:
        outputScale !== 1 ? [outputScale, 0, 0, outputScale, 0, 0] : undefined,
      viewport,
    });

    renderTask.promise.catch((error: unknown) => {
      if (
        error instanceof Error &&
        (error.name === "RenderingCancelledException" ||
          error.message.includes("Rendering cancelled"))
      ) {
        return;
      }

      console.error(error);
    });

    return () => {
      renderTask.cancel();
    };
  }, [page, zoom]);

  const visiblePlacements = placements.filter(
    (placement) => placement.pageIndex === pageIndex,
  );
  const width = pageSize ? pageSize.width * zoom : 640;
  const height = pageSize ? pageSize.height * zoom : 840;

  return (
    <article className="pdf-page-frame">
      <div className="page-label">Page {pageIndex + 1}</div>
      <div
        className="pdf-page-surface"
        onClick={(event) => event.stopPropagation()}
        onDragOver={(event) => {
          if (mode === "edit") {
            event.preventDefault();
            event.dataTransfer.dropEffect = "copy";
          }
        }}
        onDrop={(event) => onDrop(event, pageIndex)}
        style={{ width, height }}
      >
        <canvas ref={canvasRef} />
        <div className="signature-layer">
          {visiblePlacements.map((placement) => {
            const signature = signatures.find(
              (asset) => asset.id === placement.signatureId,
            );

            if (!signature) {
              return null;
            }

            const isSelected = selectedPlacementId === placement.id;

            return (
              <button
                className={`placed-signature${isSelected ? " is-selected" : ""}`}
                key={placement.id}
                onClick={(event) => {
                  event.stopPropagation();
                }}
                onPointerDown={(event) =>
                  startInteraction(event, placement, "move")
                }
                style={{
                  left: `${placement.x * 100}%`,
                  top: `${placement.y * 100}%`,
                  width: `${placement.width * 100}%`,
                  height: `${placement.height * 100}%`,
                }}
                type="button"
              >
                {/* eslint-disable-next-line @next/next/no-img-element */}
                <img alt="" draggable={false} src={signature.src} />
                {mode === "edit" && isSelected ? (
                  <span
                    aria-hidden="true"
                    className="resize-handle"
                    onPointerDown={(event) =>
                      startInteraction(event, placement, "resize")
                    }
                  />
                ) : null}
              </button>
            );
          })}
        </div>
      </div>
    </article>
  );
}
