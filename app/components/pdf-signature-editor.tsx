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

type BaseSignatureOption = {
  id: string;
  name: string;
  description?: string;
  disabled?: boolean;
  attributes?: Record<string, string | number | boolean>;
};

export type ImageSignatureOption = BaseSignatureOption & {
  kind: "image";
  src: string;
  mimeType: string;
};

export type TextSignatureOption = BaseSignatureOption & {
  kind: "text";
  value: string;
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: "normal" | "bold";
  aspectRatio?: number;
};

export type SignatureOption = ImageSignatureOption | TextSignatureOption;

export type PdfSignatureEditorOptions = {
  title?: string;
  initialStatus?: string;
  documentUploadLabel?: string;
  signaturesTitle?: string;
  emptySignaturesText?: string;
  loadingSignaturesText?: string;
  emptyDocumentTitle?: string;
  emptyDocumentDescription?: string;
  maxPdfBytes?: number;
  maxDocxBytes?: number;
  maxPages?: number;
  maxPlacements?: number;
  defaultSignatureWidthPx?: number;
  signatureAspectRatio?: number;
  minSignatureRatio?: number;
};

type PdfSignatureEditorProps = {
  signatures: SignatureOption[];
  options?: PdfSignatureEditorOptions;
};

type ImageSignatureAsset = ImageSignatureOption & {
  bytes: Uint8Array;
};

type TextSignatureAsset = TextSignatureOption;

type SignatureAsset = ImageSignatureAsset | TextSignatureAsset;

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
  textStyle?: TextPlacementStyle;
};

type TextPlacementStyle = {
  align?: "left" | "center" | "right";
  fontFamily?: string;
  fontSize?: number;
  fontStyle?: "normal" | "italic";
  fontWeight?: "normal" | "bold";
  textDecoration?: "none" | "underline";
  value?: string;
};

type ResolvedTextStyle = {
  align: "left" | "center" | "right";
  fontFamily: string;
  fontSize: number;
  fontStyle: "normal" | "italic";
  fontWeight: "normal" | "bold";
  textDecoration: "none" | "underline";
  value: string;
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

type PlacementContextMenu = {
  placementId: string | null;
  x: number;
  y: number;
} | null;

type TextPropertiesModalState = {
  activeTab: "general" | "appearance";
  placementId: string;
} | null;

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
const DEFAULT_TEXT_TAG_FONT_SIZE = 13;
const TEXT_TAG_PADDING_PX = 1;
const TEXT_TAG_WIDTH_FACTOR = 0.62;
const CANVAS_EXPORT_SCALE = 2;
const WORD_PAGE_GAP_PX = 24;

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

function isImageSignature(signature: SignatureAsset): signature is ImageSignatureAsset {
  return signature.kind === "image";
}

function isTextSignature(signature: SignatureAsset): signature is TextSignatureAsset {
  return signature.kind === "text";
}

async function loadSignatureAsset(signature: SignatureOption): Promise<SignatureAsset> {
  if (signature.kind === "text") {
    return signature;
  }

  if (signature.mimeType === "image/png" || signature.mimeType === "image/jpeg") {
    const response = await fetch(signature.src);

    return {
      ...signature,
      bytes: new Uint8Array(await response.arrayBuffer()),
    };
  }

  return {
    ...signature,
    bytes: await svgDataUrlToPngBytes(signature.src),
    mimeType: "image/png",
  };
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

function getResolvedTextStyle(
  signature: TextSignatureAsset,
  placement?: Placement,
): ResolvedTextStyle {
  return {
    align: placement?.textStyle?.align ?? "center",
    fontFamily:
      placement?.textStyle?.fontFamily ?? signature.fontFamily ?? "Arial",
    fontSize:
      placement?.textStyle?.fontSize ??
      signature.fontSize ??
      DEFAULT_TEXT_TAG_FONT_SIZE,
    fontStyle: placement?.textStyle?.fontStyle ?? "normal",
    fontWeight:
      placement?.textStyle?.fontWeight ?? signature.fontWeight ?? "bold",
    textDecoration: placement?.textStyle?.textDecoration ?? "none",
    value: placement?.textStyle?.value ?? signature.value,
  };
}

function getTextTagMetrics({
  fontSize,
  text,
  scale = 1,
}: {
  fontSize?: number;
  scale?: number;
  text: string;
}) {
  const resolvedFontSize = (fontSize ?? DEFAULT_TEXT_TAG_FONT_SIZE) * scale;
  const horizontalPadding = TEXT_TAG_PADDING_PX * scale;
  const estimatedTextWidth =
    text.length * resolvedFontSize * TEXT_TAG_WIDTH_FACTOR;

  return {
    boxHeight: resolvedFontSize + horizontalPadding * 2,
    boxWidth: estimatedTextWidth + horizontalPadding * 2,
    fontSize: resolvedFontSize,
    horizontalPadding,
    lineHeight: resolvedFontSize,
  };
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

function buildAnchoredTextBoxXml({
  align,
  drawingId,
  fontFamily,
  fontSizePx,
  fontStyle,
  fontWeight,
  heightPx,
  textDecoration,
  text,
  widthPx,
  xPx,
  yPx,
}: {
  align: "left" | "center" | "right";
  drawingId: number;
  fontFamily: string;
  fontSizePx: number;
  fontStyle: "normal" | "italic";
  fontWeight: "normal" | "bold";
  heightPx: number;
  textDecoration: "none" | "underline";
  text: string;
  widthPx: number;
  xPx: number;
  yPx: number;
}) {
  const widthEmu = pixelsToEmu(widthPx);
  const heightEmu = pixelsToEmu(heightPx);
  const xEmu = pixelsToEmu(xPx);
  const yEmu = pixelsToEmu(yPx);
  const escapedText = escapeXml(text);
  const escapedFontFamily = escapeXml(fontFamily);
  const fontSizeHalfPoints = Math.max(1, Math.round(fontSizePx * 1.5));
  const justification =
    align === "left" ? "left" : align === "right" ? "right" : "center";
  const boldXml = fontWeight === "bold" ? "<w:b/>" : "";
  const italicXml = fontStyle === "italic" ? "<w:i/>" : "";
  const underlineXml =
    textDecoration === "underline" ? '<w:u w:val="single"/>' : "";

  return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${xEmu}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${yEmu}</wp:posOffset></wp:positionV><wp:extent cx="${widthEmu}" cy="${heightEmu}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${drawingId}" name="Text tag ${drawingId}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="${justification}"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${escapedFontFamily}" w:hAnsi="${escapedFontFamily}"/>${boldXml}${italicXml}${underlineXml}<w:sz w:val="${fontSizeHalfPoints}"/></w:rPr><w:t>${escapedText}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
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

  if (isTextSignature(signature)) {
    const textStyle = getResolvedTextStyle(signature, placement);
    const height = placement.height * canvas.height;
    const width = placement.width * canvas.width;
    const x = placement.x * canvas.width;
    const y = placement.y * canvas.height;
    const metrics = getTextTagMetrics({
      fontSize: textStyle.fontSize,
      scale: CANVAS_EXPORT_SCALE,
      text: textStyle.value,
    });

    context.fillStyle = "#172033";
    context.font = `${textStyle.fontStyle} ${textStyle.fontWeight === "normal" ? "400" : "700"} ${metrics.fontSize}px ${textStyle.fontFamily}, Arial, Helvetica, sans-serif`;
    context.textAlign = textStyle.align;
    context.textBaseline = "middle";
    const textX =
      textStyle.align === "left"
        ? x + metrics.horizontalPadding
        : textStyle.align === "right"
          ? x + width - metrics.horizontalPadding
          : x + width / 2;
    context.fillText(
      textStyle.value,
      textX,
      y + height / 2,
      Math.max(1, width - metrics.horizontalPadding * 2),
    );
    return;
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

function getDefaultPlacementSize({
  pageSize,
  signatureAspectRatio,
  signature,
  signatureWidthPx,
}: {
  pageSize: PageSize;
  signatureAspectRatio: number;
  signature: SignatureAsset;
  signatureWidthPx: number;
}) {
  if (isTextSignature(signature)) {
    const textStyle = getResolvedTextStyle(signature);
    const metrics = getTextTagMetrics({
      fontSize: textStyle.fontSize,
      text: textStyle.value,
    });

    return {
      height: metrics.boxHeight / pageSize.height,
      width: metrics.boxWidth / pageSize.width,
    };
  }

  const width = Math.min(0.34, signatureWidthPx / pageSize.width);

  return {
    height: width * signatureAspectRatio,
    width,
  };
}

function getTextPlacementSize({
  pageSize,
  placement,
  signature,
}: {
  pageSize: PageSize;
  placement?: Placement;
  signature: TextSignatureAsset;
}) {
  const textStyle = getResolvedTextStyle(signature, placement);
  const metrics = getTextTagMetrics({
    fontSize: textStyle.fontSize,
    text: textStyle.value,
  });

  return {
    height: metrics.boxHeight / pageSize.height,
    width: metrics.boxWidth / pageSize.width,
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
  signatureAspectRatio,
  signature,
  signatureWidthPx,
  zoom,
}: {
  event: React.DragEvent<HTMLElement>;
  signatureAspectRatio: number;
  signature: SignatureAsset;
  signatureWidthPx: number;
  zoom: number;
}) {
  const dragImage = document.createElement("div");
  const image = document.createElement("img");
  const label = document.createElement("span");
  const textStyle = isTextSignature(signature)
    ? getResolvedTextStyle(signature)
    : null;
  const textMetrics = isTextSignature(signature)
    ? getTextTagMetrics({
        fontSize: textStyle?.fontSize,
        scale: zoom,
        text: signature.value,
      })
    : null;
  const dragHeight = textMetrics
    ? textMetrics.boxHeight
    : signatureWidthPx * signatureAspectRatio * zoom;
  const dragWidth = textMetrics ? textMetrics.boxWidth : signatureWidthPx * zoom;
  const offsetX = dragWidth / 2;
  const offsetY = dragHeight / 2;

  dragImage.className = "signature-drag-preview";
  dragImage.style.height = `${dragHeight}px`;
  dragImage.style.width = `${dragWidth}px`;

  if (isImageSignature(signature)) {
    image.src = signature.src;
    image.alt = "";
    dragImage.append(image);
  } else {
    label.textContent = signature.value;
    label.style.fontFamily = `${textStyle?.fontFamily ?? "Arial"}, Helvetica, sans-serif`;
    label.style.fontSize = `${textMetrics?.fontSize ?? DEFAULT_TEXT_TAG_FONT_SIZE}px`;
    label.style.fontStyle = textStyle?.fontStyle ?? "normal";
    label.style.fontWeight = textStyle?.fontWeight === "normal" ? "400" : "800";
    label.style.lineHeight = "1";
    label.style.padding = `0 ${textMetrics?.horizontalPadding ?? TEXT_TAG_PADDING_PX}px`;
    label.style.textDecoration = textStyle?.textDecoration ?? "none";
    dragImage.classList.add("is-text");
    dragImage.append(label);
  }
  document.body.append(dragImage);
  event.dataTransfer.setDragImage(dragImage, offsetX, offsetY);
  event.dataTransfer.setData(
    "application/pdf-sign-tag-anchor",
    JSON.stringify({ offsetX, offsetY } satisfies DragAnchor),
  );
  window.requestAnimationFrame(() => dragImage.remove());
}

export default function PdfSignatureEditor({
  options,
  signatures: signatureOptions,
}: PdfSignatureEditorProps) {
  const editorOptions = useMemo(
    () => ({
      defaultSignatureWidthPx:
        options?.defaultSignatureWidthPx ?? DEFAULT_SIGNATURE_WIDTH_PX,
      documentUploadLabel: options?.documentUploadLabel ?? "PDF / DOCX",
      emptyDocumentDescription:
        options?.emptyDocumentDescription ??
        "Upload a PDF or DOCX, then drag a signature onto any page.",
      emptyDocumentTitle:
        options?.emptyDocumentTitle ??
        "Open a PDF or Word file to start placing signatures",
      emptySignaturesText: options?.emptySignaturesText ?? "No signatures available.",
      initialStatus: options?.initialStatus ?? "Upload a PDF or DOCX to begin.",
      loadingSignaturesText:
        options?.loadingSignaturesText ?? "Loading signatures.",
      maxDocxBytes: options?.maxDocxBytes ?? MAX_DOCX_BYTES,
      maxPages: options?.maxPages ?? MAX_PAGES,
      maxPdfBytes: options?.maxPdfBytes ?? MAX_PDF_BYTES,
      maxPlacements: options?.maxPlacements ?? MAX_PLACEMENTS,
      minSignatureRatio: options?.minSignatureRatio ?? MIN_SIGNATURE_RATIO,
      signatureAspectRatio:
        options?.signatureAspectRatio ?? SIGNATURE_ASPECT_RATIO,
      signaturesTitle: options?.signaturesTitle ?? "Signatures",
      title: options?.title ?? "PDF Sign Tag",
    }),
    [options],
  );
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
  const [contextMenu, setContextMenu] = useState<PlacementContextMenu>(null);
  const [textPropertiesModal, setTextPropertiesModal] =
    useState<TextPropertiesModalState>(null);
  const [clipboardPlacement, setClipboardPlacement] = useState<Placement | null>(
    null,
  );
  const [status, setStatus] = useState(editorOptions.initialStatus);
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
      signatureOptions.map((signature) => loadSignatureAsset(signature)),
    )
      .then((loadedSignatures) => {
        if (mounted) {
          setSignatures(loadedSignatures);
        }
      })
      .catch(() => {
        if (mounted) {
          setStatus("Unable to load signatures.");
        }
      });

    return () => {
      mounted = false;
    };
  }, [signatureOptions]);

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
    setClipboardPlacement(null);
    setContextMenu(null);
    setTextPropertiesModal(null);
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

        if (file.size > editorOptions.maxPdfBytes) {
          setStatus(
            `PDF rejected: maximum file size is ${Math.round(
              editorOptions.maxPdfBytes / 1024 / 1024,
            )} MB.`,
          );
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

        if (loadedPdf.numPages > editorOptions.maxPages) {
          await loadedPdf.destroy();
          setStatus(
            `PDF rejected: maximum page count is ${editorOptions.maxPages}.`,
          );
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

      if (file.size > editorOptions.maxDocxBytes) {
        setStatus(
          `DOCX rejected: maximum file size is ${Math.round(
            editorOptions.maxDocxBytes / 1024 / 1024,
          )} MB.`,
        );
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
    if (placements.length >= editorOptions.maxPlacements) {
      setStatus(
        `Placement limit reached: maximum ${editorOptions.maxPlacements}.`,
      );
      return;
    }

    const pageSize = pageSizesRef.current[pageIndex];
    const signature = signatures.find((asset) => asset.id === signatureId);

    if (!pageSize || !signature) {
      return;
    }

    const { width: defaultWidth, height: defaultHeight } =
      getDefaultPlacementSize({
        pageSize,
        signatureAspectRatio: editorOptions.signatureAspectRatio,
        signature,
        signatureWidthPx: editorOptions.defaultSignatureWidthPx,
      });
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
    setContextMenu(null);
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

    if (placements.length >= editorOptions.maxPlacements) {
      setStatus(
        `Placement limit reached: maximum ${editorOptions.maxPlacements}.`,
      );
      return;
    }

    const signatureId = event.dataTransfer.getData("application/pdf-sign-tag");
    const signature = signatures.find((asset) => asset.id === signatureId);

    if (!signature) {
      return;
    }

    const anchor = getDragAnchor(event);
    const bounds = event.currentTarget.getBoundingClientRect();
    const pageSize = {
      height: bounds.height / zoom,
      width: bounds.width / zoom,
    };
    const { width: defaultWidth, height: defaultHeight } =
      getDefaultPlacementSize({
        pageSize,
        signatureAspectRatio: editorOptions.signatureAspectRatio,
        signature,
        signatureWidthPx: editorOptions.defaultSignatureWidthPx,
      });
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
    setContextMenu(null);
  };

  const startInteraction = (
    event: React.PointerEvent<HTMLElement>,
    placement: Placement,
    kind: "move" | "resize",
  ) => {
    if (mode !== "edit" || event.button !== 0) {
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
            editorOptions.minSignatureRatio,
            1 - interaction.startPlacement.x,
          );
          const height = clamp(
            interaction.startPlacement.height + ratioDeltaY,
            editorOptions.minSignatureRatio,
            1 - interaction.startPlacement.y,
          );
          const signature = signaturesRef.current.find(
            (asset) => asset.id === placement.signatureId,
          );

          if (signature && isTextSignature(signature)) {
            const nextFontSize = Math.max(
              8,
              height * pageSize.height - TEXT_TAG_PADDING_PX * 2,
            );
            const nextPlacement = {
              ...placement,
              textStyle: {
                ...placement.textStyle,
                fontSize: nextFontSize,
              },
            };
            const nextSize = getTextPlacementSize({
              pageSize,
              placement: nextPlacement,
              signature,
            });

            return {
              ...nextPlacement,
              height: clamp(
                nextSize.height,
                editorOptions.minSignatureRatio,
                1 - interaction.startPlacement.y,
              ),
              width: clamp(
                nextSize.width,
                editorOptions.minSignatureRatio,
                1 - interaction.startPlacement.x,
              ),
            };
          }

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
  }, [editorOptions.minSignatureRatio, interaction, zoom]);

  const removeSelectedPlacement = () => {
    if (!selectedPlacementId || mode !== "edit") {
      return;
    }

    setPlacements((current) =>
      current.filter((placement) => placement.id !== selectedPlacementId),
    );
    setSelectedPlacementId(null);
    setContextMenu(null);
  };

  const openPlacementContextMenu = (
    event: React.MouseEvent<HTMLElement>,
    placement: Placement,
  ) => {
    event.preventDefault();
    event.stopPropagation();

    const bounds = event.currentTarget.getBoundingClientRect();

    setSelectedPlacementId(placement.id);
    setContextMenu({
      placementId: placement.id,
      x: bounds.left,
      y: bounds.bottom,
    });
  };

  const clearCanvasSelection = () => {
    setContextMenu(null);
  };

  const openCanvasContextMenu = (event: React.MouseEvent<HTMLElement>) => {
    event.preventDefault();
    event.stopPropagation();

    if (mode !== "edit" || (!selectedPlacementId && !clipboardPlacement)) {
      return;
    }

    const placement = selectedPlacementId
      ? placements.find((item) => item.id === selectedPlacementId)
      : null;

    if (selectedPlacementId && !placement && !clipboardPlacement) {
      return;
    }

    setContextMenu({
      placementId: placement?.id ?? null,
      x: event.clientX,
      y: event.clientY,
    });
  };

  const copyPlacement = (placement: Placement) => {
    setClipboardPlacement(placement);
    setStatus("Signature tag copied.");
  };

  const deletePlacement = (placementId: string) => {
    setPlacements((current) =>
      current.filter((placement) => placement.id !== placementId),
    );
    setSelectedPlacementId(null);
    setContextMenu(null);
  };

  const pastePlacement = (sourcePlacement: Placement) => {
    const pageSize = pageSizesRef.current[sourcePlacement.pageIndex];

    if (!pageSize) {
      return;
    }

    const nextPlacement: Placement = {
      ...sourcePlacement,
      id: makeId("placement"),
      x: clamp(sourcePlacement.x + 0.025, 0, 1 - sourcePlacement.width),
      y: clamp(sourcePlacement.y + 0.025, 0, 1 - sourcePlacement.height),
    };

    setPlacements((current) => {
      if (current.length >= editorOptions.maxPlacements) {
        return current;
      }

      return [...current, nextPlacement];
    });
    setSelectedPlacementId(nextPlacement.id);
    setContextMenu(null);
  };

  const handleContextMenuAction = (
    action: "cut" | "copy" | "paste" | "delete" | "properties",
  ) => {
    const placement = placements.find(
      (item) => item.id === contextMenu?.placementId,
    );

    if (action === "paste") {
      const sourcePlacement = clipboardPlacement ?? placement;

      if (sourcePlacement) {
        pastePlacement(sourcePlacement);
      } else {
        setContextMenu(null);
      }
      return;
    }

    if (!placement) {
      setContextMenu(null);
      return;
    }

    if (action === "copy") {
      copyPlacement(placement);
      setContextMenu(null);
      return;
    }

    if (action === "cut") {
      copyPlacement(placement);
      deletePlacement(placement.id);
      return;
    }

    if (action === "delete") {
      deletePlacement(placement.id);
      return;
    }

    const signature = signatures.find((asset) => asset.id === placement.signatureId);

    if (signature && isTextSignature(signature)) {
      setSelectedPlacementId(placement.id);
      setTextPropertiesModal({
        activeTab: "general",
        placementId: placement.id,
      });
    }

    setSelectedPlacementId(placement.id);
    setContextMenu(null);
  };

  useEffect(() => {
    if (!contextMenu) {
      return;
    }

    const closeMenu = () => setContextMenu(null);

    window.addEventListener("click", closeMenu);
    window.addEventListener("keydown", closeMenu);
    window.addEventListener("scroll", closeMenu, true);

    return () => {
      window.removeEventListener("click", closeMenu);
      window.removeEventListener("keydown", closeMenu);
      window.removeEventListener("scroll", closeMenu, true);
    };
  }, [contextMenu]);

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

        const pageWidth = page.getWidth();
        const pageHeight = page.getHeight();
        const width = placement.width * pageWidth;
        const height = placement.height * pageHeight;

        if (isTextSignature(signature)) {
          const textStyle = getResolvedTextStyle(signature, placement);
          const metrics = getTextTagMetrics({
            fontSize: textStyle.fontSize,
            text: textStyle.value,
          });
          const measuredWidth =
            textStyle.value.length * metrics.fontSize * TEXT_TAG_WIDTH_FACTOR;
          const textX =
            textStyle.align === "left"
              ? placement.x * pageWidth + metrics.horizontalPadding
              : textStyle.align === "right"
                ? placement.x * pageWidth +
                  width -
                  metrics.horizontalPadding -
                  measuredWidth
                : placement.x * pageWidth + (width - measuredWidth) / 2;

          page.drawText(textStyle.value, {
            x: textX,
            y:
              pageHeight -
              placement.y * pageHeight -
              height / 2 -
              metrics.fontSize * 0.35,
            maxWidth: Math.max(1, width - metrics.horizontalPadding * 2),
            size: metrics.fontSize,
          });
        } else {
          let image = embeddedImages.get(signature.id);

          if (!image) {
            image =
              signature.mimeType === "image/png"
                ? await outputPdf.embedPng(signature.bytes)
                : await outputPdf.embedJpg(signature.bytes);
            embeddedImages.set(signature.id, image);
          }

          page.drawImage(image, {
            x: placement.x * pageWidth,
            y: pageHeight - placement.y * pageHeight - height,
            width,
            height,
          });
        }
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
        const pageSize = pageSizes[placement.pageIndex] ?? fallbackPageSize;
        const widthPx = placement.width * pageSize.width;
        const heightPx = placement.height * pageSize.height;
        const xPx = placement.x * pageSize.width;
        const yPx = placement.y * pageSize.height;
        let drawingXml: string;

        if (isTextSignature(signature)) {
          const textStyle = getResolvedTextStyle(signature, placement);
          const metrics = getTextTagMetrics({
            fontSize: textStyle.fontSize,
            text: textStyle.value,
          });

          drawingXml = buildAnchoredTextBoxXml({
            align: textStyle.align,
            drawingId: nextDrawingId,
            fontFamily: textStyle.fontFamily,
            fontSizePx: metrics.fontSize,
            fontStyle: textStyle.fontStyle,
            fontWeight: textStyle.fontWeight,
            heightPx,
            text: textStyle.value,
            textDecoration: textStyle.textDecoration,
            widthPx,
            xPx,
            yPx,
          });
        } else {
          const relationshipId = `rId${nextRelationshipId}`;
          const extension = getXmlMimeExtension(signature.mimeType);
          const contentType = getXmlMimeContentType(signature.mimeType);
          const mediaName = `signature-${placement.id}.${extension}`;
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
          drawingXml = buildAnchoredDrawingXml({
            drawingId: nextDrawingId,
            heightPx,
            name: signature.name,
            relationshipId,
            widthPx,
            xPx,
            yPx,
          });
          nextRelationshipId += 1;
        }

        const wrapperXml = parser.parseFromString(
          `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${drawingXml}</root>`,
          "application/xml",
        );
        const importedNode = wrapperXml.documentElement.firstElementChild;

        if (!importedNode) {
          throw new Error("Unable to build signature drawing.");
        }

        targetParagraph.append(documentXml.importNode(importedNode, true));
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
  const textPropertiesPlacement =
    placements.find((placement) => placement.id === textPropertiesModal?.placementId) ??
    null;
  const textPropertiesSignature = textPropertiesPlacement
    ? signatures.find((asset) => asset.id === textPropertiesPlacement.signatureId)
    : null;
  const textPropertiesStyle =
    textPropertiesPlacement &&
    textPropertiesSignature &&
    isTextSignature(textPropertiesSignature)
      ? getResolvedTextStyle(textPropertiesSignature, textPropertiesPlacement)
      : null;
  const updateTextPlacementStyle = (updates: TextPlacementStyle) => {
    if (
      !textPropertiesPlacement ||
      !textPropertiesSignature ||
      !isTextSignature(textPropertiesSignature)
    ) {
      return;
    }

    setPlacements((current) =>
      current.map((placement) => {
        if (placement.id !== textPropertiesPlacement.id) {
          return placement;
        }

        const nextPlacement = {
          ...placement,
          textStyle: {
            ...placement.textStyle,
            ...updates,
          },
        };
        const pageSize = pageSizesRef.current[placement.pageIndex];

        if (!pageSize) {
          return nextPlacement;
        }

        const nextSize = getTextPlacementSize({
          pageSize,
          placement: nextPlacement,
          signature: textPropertiesSignature,
        });

        return {
          ...nextPlacement,
          height: clamp(
            nextSize.height,
            editorOptions.minSignatureRatio,
            1 - placement.y,
          ),
          width: clamp(
            nextSize.width,
            editorOptions.minSignatureRatio,
            1 - placement.x,
          ),
        };
      }),
    );
  };

  return (
    <main className="editor-shell">
      <header className="editor-toolbar">
        <div className="brand-block">
          <span className="brand-mark">PS</span>
          <div>
            <h1>{editorOptions.title}</h1>
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
            {editorOptions.documentUploadLabel}
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
            <h2>{editorOptions.signaturesTitle}</h2>
            <span>{signatures.length}</span>
          </div>

          <div className="signature-list">
            {signatures.length === 0 ? (
              <p className="empty-copy">
                {signatureOptions.length === 0
                  ? editorOptions.emptySignaturesText
                  : editorOptions.loadingSignaturesText}
              </p>
            ) : (
              signatures.map((signature) => (
                <button
                  className="signature-tile"
                  draggable={mode === "edit"}
                  disabled={signature.disabled}
                  key={signature.id}
                  onDragStart={(event) => {
                    if (signature.disabled) {
                      event.preventDefault();
                      return;
                    }

                    event.dataTransfer.setData(
                      "application/pdf-sign-tag",
                      signature.id,
                    );
                    event.dataTransfer.effectAllowed = "copy";
                    setSignatureDragImage({
                      event,
                      signature,
                      signatureAspectRatio: editorOptions.signatureAspectRatio,
                      signatureWidthPx: editorOptions.defaultSignatureWidthPx,
                      zoom,
                    });
                  }}
                  type="button"
                  title={signature.description}
                >
                  {isImageSignature(signature) ? (
                    // eslint-disable-next-line @next/next/no-img-element
                    <img alt={signature.name} src={signature.src} />
                  ) : (
                    <strong className="signature-text-preview">
                      {signature.value}
                    </strong>
                  )}
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
              <h2>{editorOptions.emptyDocumentTitle}</h2>
              <p>{editorOptions.emptyDocumentDescription}</p>
            </div>
          ) : documentType === "pdf" && pdfDocument ? (
            <div className="page-stack" onClick={() => setSelectedPlacementId(null)}>
              {Array.from({ length: pdfDocument.numPages }, (_, index) => (
                <PdfPageSurface
                  key={index}
                  mode={mode}
                  onDrop={handlePageDrop}
                  onPageBackgroundClick={clearCanvasSelection}
                  onPageBackgroundContextMenu={openCanvasContextMenu}
                  onPlacementContextMenu={openPlacementContextMenu}
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
              onPageBackgroundClick={clearCanvasSelection}
              onPageBackgroundContextMenu={openCanvasContextMenu}
              onPlacementContextMenu={openPlacementContextMenu}
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
      {contextMenu ? (
        <div
          className="tag-context-menu"
          onClick={(event) => event.stopPropagation()}
          style={{
            left: contextMenu.x,
            top: contextMenu.y,
          }}
        >
          <button
            disabled={!contextMenu.placementId}
            onClick={() => handleContextMenuAction("cut")}
            type="button"
          >
            <span aria-hidden="true">⌘</span>
            Cut
          </button>
          <button
            disabled={!contextMenu.placementId}
            onClick={() => handleContextMenuAction("copy")}
            type="button"
          >
            <span aria-hidden="true">□</span>
            Copy
          </button>
          <button
            disabled={!clipboardPlacement}
            onClick={() => handleContextMenuAction("paste")}
            type="button"
          >
            <span aria-hidden="true">▣</span>
            Paste
          </button>
          <button
            disabled={!contextMenu.placementId}
            onClick={() => handleContextMenuAction("delete")}
            type="button"
          >
            <span aria-hidden="true">⌫</span>
            Delete
          </button>
          <button
            disabled={!contextMenu.placementId}
            onClick={() => handleContextMenuAction("properties")}
            type="button"
          >
            <span aria-hidden="true">▤</span>
            Properties
          </button>
        </div>
      ) : null}
      {textPropertiesModal && textPropertiesStyle ? (
        <div className="properties-backdrop">
          <section className="properties-dialog" role="dialog" aria-modal="true">
            <header className="properties-header">
              <h2>Textbox Properties</h2>
              <button
                aria-label="Close properties"
                onClick={() => setTextPropertiesModal(null)}
                type="button"
              >
                ×
              </button>
            </header>
            <div className="properties-tabs">
              <button
                aria-pressed={textPropertiesModal.activeTab === "general"}
                onClick={() =>
                  setTextPropertiesModal({
                    ...textPropertiesModal,
                    activeTab: "general",
                  })
                }
                type="button"
              >
                General
              </button>
              <button
                aria-pressed={textPropertiesModal.activeTab === "appearance"}
                onClick={() =>
                  setTextPropertiesModal({
                    ...textPropertiesModal,
                    activeTab: "appearance",
                  })
                }
                type="button"
              >
                Appearance
              </button>
            </div>
            {textPropertiesModal.activeTab === "general" ? (
              <div className="properties-grid">
                <label>
                  <span>Name</span>
                  <input readOnly value={textPropertiesSignature?.name ?? ""} />
                </label>
                <label>
                  <span>Tooltip</span>
                  <input
                    readOnly
                    value={textPropertiesSignature?.description ?? ""}
                  />
                </label>
                <label>
                  <span>Value</span>
                  <input
                    onChange={(event) =>
                      updateTextPlacementStyle({ value: event.target.value })
                    }
                    value={textPropertiesStyle.value}
                  />
                </label>
                <label>
                  <span>Form Field Visibility</span>
                  <select defaultValue="visible">
                    <option value="visible">visible</option>
                  </select>
                </label>
              </div>
            ) : (
              <div className="properties-appearance">
                <label>
                  <span>Font</span>
                  <select
                    onChange={(event) =>
                      updateTextPlacementStyle({
                        fontFamily: event.target.value,
                      })
                    }
                    value={textPropertiesStyle.fontFamily}
                  >
                    <option value="Arial">Arial</option>
                    <option value="Helvetica">Helvetica</option>
                    <option value="Times New Roman">Times New Roman</option>
                    <option value="Courier New">Courier New</option>
                  </select>
                </label>
                <label>
                  <span>Size</span>
                  <input
                    min={8}
                    onChange={(event) =>
                      updateTextPlacementStyle({
                        fontSize:
                          Number.parseFloat(event.target.value) || undefined,
                      })
                    }
                    type="number"
                    value={Math.round(textPropertiesStyle.fontSize ?? 0) || ""}
                  />
                </label>
                <div className="format-buttons">
                  <button
                    aria-pressed={textPropertiesStyle.fontWeight === "bold"}
                    onClick={() =>
                      updateTextPlacementStyle({
                        fontWeight:
                          textPropertiesStyle.fontWeight === "bold"
                            ? "normal"
                            : "bold",
                      })
                    }
                    type="button"
                  >
                    B
                  </button>
                  <button
                    aria-pressed={textPropertiesStyle.fontStyle === "italic"}
                    onClick={() =>
                      updateTextPlacementStyle({
                        fontStyle:
                          textPropertiesStyle.fontStyle === "italic"
                            ? "normal"
                            : "italic",
                      })
                    }
                    type="button"
                  >
                    I
                  </button>
                  <button
                    aria-pressed={
                      textPropertiesStyle.textDecoration === "underline"
                    }
                    onClick={() =>
                      updateTextPlacementStyle({
                        textDecoration:
                          textPropertiesStyle.textDecoration === "underline"
                            ? "none"
                            : "underline",
                      })
                    }
                    type="button"
                  >
                    U
                  </button>
                </div>
                <div className="align-buttons">
                  {(["left", "center", "right"] as const).map((align) => (
                    <button
                      aria-pressed={textPropertiesStyle.align === align}
                      key={align}
                      onClick={() => updateTextPlacementStyle({ align })}
                      type="button"
                    >
                      {align}
                    </button>
                  ))}
                </div>
              </div>
            )}
          </section>
        </div>
      ) : null}
    </main>
  );
}

function WordDocumentSurface({
  bytes,
  expectedPageCount,
  expectedPageSize,
  mode,
  onDrop,
  onPageBackgroundClick,
  onPageBackgroundContextMenu,
  onPlacementContextMenu,
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
  onPageBackgroundClick: () => void;
  onPageBackgroundContextMenu: (event: React.MouseEvent<HTMLElement>) => void;
  onPlacementContextMenu: (
    event: React.MouseEvent<HTMLElement>,
    placement: Placement,
  ) => void;
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
              onClick={(event) => {
                event.stopPropagation();
                onPageBackgroundClick();
              }}
              onContextMenu={onPageBackgroundContextMenu}
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
                  const textStyle = isTextSignature(signature)
                    ? getResolvedTextStyle(signature, placement)
                    : null;
                  const textMetrics = textStyle
                    ? getTextTagMetrics({
                        fontSize: textStyle.fontSize,
                        scale: zoom,
                        text: textStyle.value,
                      })
                    : null;

                  return (
                    <button
                      className={`placed-signature${
                        isSelected ? " is-selected" : ""
                      }`}
                      key={placement.id}
                      onClick={(event) => event.stopPropagation()}
                      onContextMenu={(event) =>
                        onPlacementContextMenu(event, placement)
                      }
                      onPointerDown={(event) =>
                        startInteraction(event, placement, "move")
                      }
                      style={{
                        height: `${placement.height * 100}%`,
                        left: `${placement.x * 100}%`,
                        top: `${placement.y * 100}%`,
                        width: `${placement.width * 100}%`,
                        ...(textMetrics
                          ? {
                              "--text-tag-font-size": `${textMetrics.fontSize}px`,
                              "--text-tag-padding-x": `${textMetrics.horizontalPadding}px`,
                              "--text-tag-font-family": textStyle?.fontFamily,
                              "--text-tag-font-style": textStyle?.fontStyle,
                              "--text-tag-font-weight": textStyle?.fontWeight,
                              "--text-tag-text-align": textStyle?.align,
                              "--text-tag-text-decoration":
                                textStyle?.textDecoration,
                            }
                          : {}),
                      }}
                      type="button"
                    >
                      {isImageSignature(signature) ? (
                        // eslint-disable-next-line @next/next/no-img-element
                        <img alt="" draggable={false} src={signature.src} />
                      ) : (
                        <span className="placed-signature-text">
                          {textStyle?.value ?? signature.value}
                        </span>
                      )}
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
  onPageBackgroundClick,
  onPageBackgroundContextMenu,
  onPlacementContextMenu,
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
  onPageBackgroundClick: () => void;
  onPageBackgroundContextMenu: (event: React.MouseEvent<HTMLElement>) => void;
  onPlacementContextMenu: (
    event: React.MouseEvent<HTMLElement>,
    placement: Placement,
  ) => void;
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
        onClick={(event) => {
          event.stopPropagation();
          onPageBackgroundClick();
        }}
        onContextMenu={onPageBackgroundContextMenu}
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
            const textStyle = isTextSignature(signature)
              ? getResolvedTextStyle(signature, placement)
              : null;
            const textMetrics = textStyle
              ? getTextTagMetrics({
                  fontSize: textStyle.fontSize,
                  scale: zoom,
                  text: textStyle.value,
                })
              : null;

            return (
              <button
                className={`placed-signature${isSelected ? " is-selected" : ""}`}
                key={placement.id}
                onClick={(event) => {
                  event.stopPropagation();
                }}
                onContextMenu={(event) =>
                  onPlacementContextMenu(event, placement)
                }
                onPointerDown={(event) =>
                  startInteraction(event, placement, "move")
                }
                style={{
                  left: `${placement.x * 100}%`,
                  top: `${placement.y * 100}%`,
                  width: `${placement.width * 100}%`,
                  height: `${placement.height * 100}%`,
                  ...(textMetrics
                    ? {
                        "--text-tag-font-size": `${textMetrics.fontSize}px`,
                        "--text-tag-padding-x": `${textMetrics.horizontalPadding}px`,
                        "--text-tag-font-family": textStyle?.fontFamily,
                        "--text-tag-font-style": textStyle?.fontStyle,
                        "--text-tag-font-weight": textStyle?.fontWeight,
                        "--text-tag-text-align": textStyle?.align,
                        "--text-tag-text-decoration": textStyle?.textDecoration,
                      }
                    : {}),
                }}
                type="button"
              >
                {isImageSignature(signature) ? (
                  // eslint-disable-next-line @next/next/no-img-element
                  <img alt="" draggable={false} src={signature.src} />
                ) : (
                  <span className="placed-signature-text">
                    {textStyle?.value ?? signature.value}
                  </span>
                )}
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
