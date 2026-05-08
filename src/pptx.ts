import { Buffer } from "node:buffer";

import { imageSize } from "image-size";
import PptxGenJS from "pptxgenjs";

import { hasExtendedRichTextBlock, renderPlainText, type NotionBlockNode, type NotionPageContent } from "./notion.ts";

type DeckSlide = {
  title?: string;
  nodes: NotionBlockNode[];
};

type LayoutBox = {
  x: number;
  y: number;
  w: number;
  h: number;
};

type ImageData = {
  data: string;
  width: number;
  height: number;
};

type RichTextRun = {
  text: string;
  options?: Record<string, unknown>;
};

type ParagraphSpec = {
  runs: RichTextRun[];
  options?: TextStyleOptions;
};

const SLIDE_WIDTH = 13.333;
const SLIDE_HEIGHT = 7.5;
const SLIDE_MARGIN_X = 0.6;
const SLIDE_MARGIN_TOP = 0.45;
const SLIDE_MARGIN_BOTTOM = 0.35;
const TITLE_HEIGHT = 0.65;
const CONTENT_TOP = SLIDE_MARGIN_TOP + TITLE_HEIGHT + 0.2;
const CONTENT_HEIGHT = SLIDE_HEIGHT - CONTENT_TOP - SLIDE_MARGIN_BOTTOM;
const COLUMN_GAP = 0.28;
const BLOCK_GAP = 0.12;
const QUOTE_GAP = 0.16;
const PX_PER_INCH = 96;
const BODY_FONT_FACE = "Inter";
const BODY_FONT_SIZE = 21;
const HEADER_FONT_SIZE = 25;
const CAPTION_FONT_SIZE = 11;
const CODE_FONT_FACE = "Courier New";
const BODY_LINE_SPACING = 1.1;
const PARAGRAPH_SPACE_AFTER = 4;
const LIST_PARAGRAPH_SPACE_AFTER = 10;
const LIST_BULLET_INDENT = 18;

type TextStyleOptions = Record<string, unknown> & {
  fontSize?: number;
  indentLevel?: number;
  bullet?: boolean | Record<string, unknown>;
};

type PptxSlide = ReturnType<PptxGenJS["addSlide"]>;

export async function renderPptxDocument(page: NotionPageContent, outputPath: string): Promise<void> {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "GitHub Copilot";
  pptx.company = "notion-to-ppt";
  pptx.subject = page.title;
  pptx.title = page.title;
  pptx.theme = {
    headFontFace: BODY_FONT_FACE,
    bodyFontFace: BODY_FONT_FACE,
  };

  const slides = splitSlides(page);

  for (const [index, deckSlide] of slides.entries()) {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    slide.color = "1F2937";
    slide.addNotes(`${page.title}\n${page.url}`);

    const titleHeight = deckSlide.title ? renderSlideTitle(slide, deckSlide.title) : 0;

    renderSlideFrame(slide, index + 1, slides.length);
    await renderNodesInBox(slide, deckSlide.nodes, {
      x: SLIDE_MARGIN_X,
      y: SLIDE_MARGIN_TOP + titleHeight + (titleHeight > 0 ? 0.12 : 0),
      w: SLIDE_WIDTH - SLIDE_MARGIN_X * 2,
      h: SLIDE_HEIGHT - (SLIDE_MARGIN_TOP + titleHeight + (titleHeight > 0 ? 0.12 : 0)) - SLIDE_MARGIN_BOTTOM,
    }, 0);
  }

  await pptx.writeFile({ fileName: outputPath });
}

function splitSlides(page: NotionPageContent): DeckSlide[] {
  const slides: DeckSlide[] = [{ nodes: [] }];

  for (const node of page.blocks) {
    const currentSlide = slides[slides.length - 1];

    if (!currentSlide) {
      throw new Error("Expected an active slide while rendering the deck.");
    }

    if (node.block.type === "heading_1") {
      slides.push({
        title: renderPlainText(node.block.heading_1.rich_text) || page.title,
        nodes: [],
      });
      continue;
    }

    if (node.block.type === "divider") {
      slides.push({ nodes: [] });
      continue;
    }

    currentSlide.nodes.push(node);
  }

  return slides.filter((slide) => slide.title || slide.nodes.length > 0);
}

function renderSlideFrame(slide: PptxSlide, slideNumber: number, slideCount: number): void {
  slide.addText(`${slideNumber}/${slideCount}`, {
    x: SLIDE_WIDTH - 1.0,
    y: SLIDE_HEIGHT - 0.32,
    w: 0.45,
    h: 0.18,
    align: "right",
    margin: 0,
    fontFace: BODY_FONT_FACE,
    fontSize: 9,
    color: "64748B",
    charSpacing: 0,
  });
}

function renderSlideTitle(slide: PptxSlide, title: string): number {
  const height = Math.min(TITLE_HEIGHT, estimateTextHeight(title, HEADER_FONT_SIZE, SLIDE_WIDTH - SLIDE_MARGIN_X * 2, 8.8));

  slide.addText(title, {
    x: SLIDE_MARGIN_X,
    y: SLIDE_MARGIN_TOP,
    w: SLIDE_WIDTH - SLIDE_MARGIN_X * 2,
    h: height,
    margin: 0,
    fontFace: BODY_FONT_FACE,
    fontSize: HEADER_FONT_SIZE,
    bold: true,
    color: "0F172A",
    valign: "top",
    fit: "none",
    charSpacing: 0,
    lineSpacingMultiple: 1.05,
  });

  return height;
}

async function renderNodesInBox(slide: PptxSlide, nodes: NotionBlockNode[], box: LayoutBox, listDepth: number): Promise<number> {
  let cursorY = box.y;

  for (let index = 0; index < nodes.length; index += 1) {
    const node = nodes[index];
    if (!node) {
      continue;
    }

    if (isListTextNode(node)) {
      const groupedNodes: NotionBlockNode[] = [node];
      while (index + 1 < nodes.length) {
        const nextNode = nodes[index + 1];
        if (!nextNode || !isListTextNode(nextNode)) {
          break;
        }
        groupedNodes.push(nextNode);
        index += 1;
      }

      const usedHeight = await renderListGroup(slide, groupedNodes, {
        x: box.x,
        y: cursorY,
        w: box.w,
        h: Math.max(0, box.h - (cursorY - box.y)),
      }, listDepth);
      cursorY += usedHeight;

      if (cursorY > box.y + box.h) {
        return box.h;
      }

      continue;
    }

    const usedHeight = await renderNode(slide, node, {
      x: box.x,
      y: cursorY,
      w: box.w,
      h: Math.max(0, box.h - (cursorY - box.y)),
    }, listDepth);
    cursorY += usedHeight;

    if (cursorY > box.y + box.h) {
      return box.h;
    }
  }

  return Math.max(0, cursorY - box.y);
}

async function renderNode(slide: PptxSlide, node: NotionBlockNode, box: LayoutBox, listDepth: number): Promise<number> {
  const block = node.block;
  const runtimeBlock = block as { type: string } & Record<string, unknown>;

  if (hasExtendedRichTextBlock(runtimeBlock, "heading_4")) {
    return renderTextWithChildren(slide, toRichTextRuns(runtimeBlock.heading_4.rich_text), box, {
      fontSize: HEADER_FONT_SIZE,
      bold: true,
      color: "1E293B",
    }, node.children, 0);
  }

  switch (block.type) {
    case "paragraph": {
      return renderTextWithChildren(slide, toRichTextRuns(block.paragraph.rich_text), box, {
        fontSize: BODY_FONT_SIZE,
        color: "1F2937",
      }, node.children, 0);
    }
    case "heading_2": {
      return renderTextWithChildren(slide, toRichTextRuns(block.heading_2.rich_text), box, {
        fontSize: HEADER_FONT_SIZE,
        bold: true,
        color: "0F172A",
      }, node.children, 0);
    }
    case "heading_3": {
      return renderTextWithChildren(slide, toRichTextRuns(block.heading_3.rich_text), box, {
        fontSize: HEADER_FONT_SIZE,
        bold: true,
        color: "0F172A",
      }, node.children, 0);
    }
    case "bulleted_list_item": {
      return renderListGroup(slide, [node], box, listDepth);
    }
    case "numbered_list_item": {
      return renderListGroup(slide, [node], box, listDepth);
    }
    case "to_do": {
      const marker = block.to_do.checked ? "[x] " : "[ ] ";
      return renderTextWithChildren(slide, withPrefix(toRichTextRuns(block.to_do.rich_text), marker), box, {
        fontSize: BODY_FONT_SIZE,
        color: "1F2937",
        indentLevel: listDepth,
      }, node.children, listDepth + 1);
    }
    case "toggle": {
      return renderTextWithChildren(slide, withPrefix(toRichTextRuns(block.toggle.rich_text), "• "), box, {
        fontSize: BODY_FONT_SIZE,
        color: "1F2937",
        indentLevel: listDepth,
      }, node.children, listDepth + 1);
    }
    case "quote": {
      return renderQuoteBlock(slide, toRichTextRuns(block.quote.rich_text), box, node.children);
    }
    case "callout": {
      const emoji = block.callout.icon?.type === "emoji" ? `${block.callout.icon.emoji} ` : "";
      return renderQuoteBlock(slide, withPrefix(toRichTextRuns(block.callout.rich_text), emoji), box, node.children, true);
    }
    case "code": {
      const language = block.code.language.trim().toLowerCase();
      return renderCodeBlock(slide, renderPlainText(block.code.rich_text), language, box);
    }
    case "image": {
      return renderImageBlock(slide, block, box);
    }
    case "bookmark": {
      return renderLinkBlock(slide, block.bookmark.url, block.bookmark.url, box);
    }
    case "embed": {
      return renderLinkBlock(slide, block.embed.url, "Embedded content", box);
    }
    case "video": {
      const videoUrl = block.video.type === "external" ? block.video.external.url : block.video.file.url;
      return renderLinkBlock(slide, videoUrl, "Video", box);
    }
    case "file": {
      const fileUrl = block.file.type === "external" ? block.file.external.url : block.file.file.url;
      const caption = renderPlainText(block.file.caption) || "File";
      return renderLinkBlock(slide, fileUrl, caption, box);
    }
    case "equation": {
      return renderSimpleText(slide, [{ text: `$${block.equation.expression}$` }], box, {
        fontSize: BODY_FONT_SIZE,
        italic: true,
        color: "1F2937",
      });
    }
    case "table_of_contents": {
      return 0;
    }
    case "child_page": {
      return renderSimpleText(slide, [{ text: block.child_page.title }], box, {
        fontSize: HEADER_FONT_SIZE,
        bold: true,
        color: "0F172A",
      });
    }
    case "column_list": {
      return renderColumnListBlock(slide, node, box);
    }
    case "column": {
      return renderNodesInBox(slide, node.children, box, 0);
    }
    case "table": {
      return renderTableBlock(slide, node, box);
    }
    case "table_row": {
      return 0;
    }
    default: {
      return renderNodesInBox(slide, node.children, box, listDepth);
    }
  }
}

async function renderTextWithChildren(
  slide: PptxSlide,
  runs: RichTextRun[],
  box: LayoutBox,
  textOptions: TextStyleOptions,
  children: NotionBlockNode[],
  childDepth: number,
): Promise<number> {
  let usedHeight = 0;

  if (runs.length > 0) {
    usedHeight += renderSimpleText(slide, runs, box, textOptions);
  }

  if (children.length > 0) {
    const childHeight = await renderNodesInBox(slide, children, {
      x: box.x,
      y: box.y + usedHeight,
      w: box.w,
      h: Math.max(0, box.h - usedHeight),
    }, childDepth);
    usedHeight += childHeight;
  }

  return usedHeight;
}

function renderSimpleText(
  slide: PptxSlide,
  runs: RichTextRun[],
  box: LayoutBox,
  textOptions: TextStyleOptions,
): number {
  if (runs.length === 0 || box.h <= 0 || box.w <= 0) {
    return 0;
  }

  const fontSize = typeof textOptions.fontSize === "number" ? textOptions.fontSize : BODY_FONT_SIZE;
  const indentLevel = typeof textOptions.indentLevel === "number" ? textOptions.indentLevel : 0;
  const indentOffset = indentLevel * 0.35;
  const plainText = runs.map((run) => run.text).join("");
  const height = Math.min(box.h, estimateTextHeight(plainText, fontSize, Math.max(0.4, box.w - indentOffset), 8.8));
  const options = { ...textOptions };

  delete options.indentLevel;

  slide.addText(runs, {
    x: box.x + indentOffset,
    y: box.y,
    w: Math.max(0.4, box.w - indentOffset),
    h: height,
    margin: 0,
    breakLine: false,
    fit: "none",
    valign: "top",
    fontFace: BODY_FONT_FACE,
    charSpacing: 0,
    lineSpacingMultiple: BODY_LINE_SPACING,
    ...options,
  });

  return height + BLOCK_GAP;
}

function renderParagraphText(
  slide: PptxSlide,
  paragraphs: ParagraphSpec[],
  box: LayoutBox,
  defaultOptions: TextStyleOptions,
  trailingGap = BLOCK_GAP,
): number {
  if (paragraphs.length === 0 || box.h <= 0 || box.w <= 0) {
    return 0;
  }

  const normalizedParagraphs = paragraphs
    .map((paragraph) => normalizeParagraph(paragraph, defaultOptions))
    .filter((paragraph) => paragraph.runs.length > 0);

  if (normalizedParagraphs.length === 0) {
    return 0;
  }

  const maxFontSize = normalizedParagraphs.reduce((largest, paragraph) => {
    const fontSize = typeof paragraph.options.fontSize === "number" ? paragraph.options.fontSize : BODY_FONT_SIZE;
    return Math.max(largest, fontSize);
  }, BODY_FONT_SIZE);
  const height = Math.min(box.h, estimateParagraphHeight(normalizedParagraphs, box.w, maxFontSize));
  const textRuns = flattenParagraphs(normalizedParagraphs);
  const options = { ...defaultOptions };

  delete options.indentLevel;
  delete options.bullet;
  delete options.paraSpaceAfter;
  delete options.paraSpaceBefore;

  slide.addText(textRuns, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: height,
    margin: 0,
    breakLine: false,
    fit: "none",
    valign: "top",
    fontFace: BODY_FONT_FACE,
    charSpacing: 0,
    lineSpacingMultiple: BODY_LINE_SPACING,
    ...options,
  });

  return height + trailingGap;
}

function normalizeParagraph(paragraph: ParagraphSpec, defaultOptions: TextStyleOptions): Required<ParagraphSpec> {
  const options = {
    fontSize: BODY_FONT_SIZE,
    color: "1F2937",
    lineSpacingMultiple: BODY_LINE_SPACING,
    paraSpaceAfter: PARAGRAPH_SPACE_AFTER,
    ...defaultOptions,
    ...paragraph.options,
  } as TextStyleOptions;
  const runs = paragraph.runs.length > 0 ? paragraph.runs : [{ text: " " }];

  return { runs, options };
}

function flattenParagraphs(paragraphs: Required<ParagraphSpec>[]): RichTextRun[] {
  const flattened: RichTextRun[] = [];

  paragraphs.forEach((paragraph, paragraphIndex) => {
    paragraph.runs.forEach((run, runIndex) => {
      const mergedOptions = {
        ...(runIndex === 0 ? paragraph.options : {}),
        ...(run.options ?? {}),
      } as Record<string, unknown>;

      if (paragraphIndex < paragraphs.length - 1 && runIndex === paragraph.runs.length - 1) {
        mergedOptions.breakLine = true;
      }

      flattened.push({
        text: run.text,
        options: mergedOptions,
      });
    });
  });

  return flattened;
}

async function renderListGroup(
  slide: PptxSlide,
  nodes: NotionBlockNode[],
  box: LayoutBox,
  listDepth: number,
): Promise<number> {
  const paragraphs = buildListParagraphs(nodes, listDepth);
  let usedHeight = 0;

  if (paragraphs.length > 0) {
    usedHeight += renderParagraphText(slide, paragraphs, box, {
      fontSize: BODY_FONT_SIZE,
      color: "1F2937",
    });
  }

  for (const node of nodes) {
    const nonListChildren = node.children.filter((child) => !isListTextNode(child));
    if (nonListChildren.length === 0) {
      continue;
    }

    const childIndent = 0.32 * (listDepth + 1);
    const childHeight = await renderNodesInBox(slide, nonListChildren, {
      x: box.x + childIndent,
      y: box.y + usedHeight,
      w: Math.max(0.5, box.w - childIndent),
      h: Math.max(0, box.h - usedHeight),
    }, listDepth + 1);
    usedHeight += childHeight;
  }

  return usedHeight;
}

function buildListParagraphs(nodes: NotionBlockNode[], listDepth: number): ParagraphSpec[] {
  const paragraphs: ParagraphSpec[] = [];
  let numberIndex = 0;

  for (const node of nodes) {
    const { paragraph, nestedNodes, shouldResetNumbering } = toListParagraph(node, listDepth, numberIndex);
    if (paragraph) {
      paragraphs.push(paragraph);
    }

    if (shouldResetNumbering) {
      numberIndex = 0;
    } else if (node.block.type === "numbered_list_item") {
      numberIndex += 1;
    }

    if (nestedNodes.length > 0) {
      paragraphs.push(...buildListParagraphs(nestedNodes, listDepth + 1));
    }
  }

  return paragraphs;
}

function toListParagraph(
  node: NotionBlockNode,
  listDepth: number,
  numberIndex: number,
): { paragraph?: ParagraphSpec; nestedNodes: NotionBlockNode[]; shouldResetNumbering: boolean } {
  switch (node.block.type) {
    case "bulleted_list_item": {
      const prefix = listDepth > 0 ? `${"    ".repeat(listDepth)}• ` : "• ";
      return {
        paragraph: {
          runs: toParagraphRuns(withPrefix(toRichTextRuns(node.block.bulleted_list_item.rich_text), prefix), {
            fontSize: BODY_FONT_SIZE,
            color: "1F2937",
            indentLevel: listDepth,
            paraSpaceAfter: LIST_PARAGRAPH_SPACE_AFTER,
          }),
          options: {
            fontSize: BODY_FONT_SIZE,
            color: "1F2937",
            indentLevel: listDepth,
            paraSpaceAfter: LIST_PARAGRAPH_SPACE_AFTER,
          },
        },
        nestedNodes: node.children.filter(isListTextNode),
        shouldResetNumbering: true,
      };
    }
    case "numbered_list_item": {
      const prefix = `${"    ".repeat(listDepth)}${numberIndex + 1}. `;
      return {
        paragraph: {
          runs: toParagraphRuns(withPrefix(toRichTextRuns(node.block.numbered_list_item.rich_text), prefix), {
            fontSize: BODY_FONT_SIZE,
            color: "1F2937",
            indentLevel: listDepth,
            paraSpaceAfter: LIST_PARAGRAPH_SPACE_AFTER,
          }),
          options: {
            fontSize: BODY_FONT_SIZE,
            color: "1F2937",
            indentLevel: listDepth,
            paraSpaceAfter: LIST_PARAGRAPH_SPACE_AFTER,
          },
        },
        nestedNodes: node.children.filter(isListTextNode),
        shouldResetNumbering: false,
      };
    }
    default:
      return { nestedNodes: [], shouldResetNumbering: true };
  }
}

function isListTextNode(node: NotionBlockNode): boolean {
  return node.block.type === "bulleted_list_item" || node.block.type === "numbered_list_item";
}

function toParagraphRuns(runs: RichTextRun[], paragraphOptions: TextStyleOptions): RichTextRun[] {
  if (runs.length === 0) {
    return [{ text: " ", options: paragraphOptions }];
  }

  return runs.map((run) => ({
    text: run.text,
    options: {
      ...paragraphOptions,
      ...(run.options ?? {}),
    },
  }));
}

async function renderQuoteBlock(
  slide: PptxSlide,
  runs: RichTextRun[],
  box: LayoutBox,
  children: NotionBlockNode[],
  shaded = false,
): Promise<number> {
  if (box.h <= 0 || box.w <= 0) {
    return 0;
  }

  const textHeight = runs.length > 0 ? Math.min(box.h, estimateTextHeight(runs.map((run) => run.text).join(""), 15, box.w - 0.35)) : 0;

  if (shaded) {
    slide.addShape("roundRect", {
      x: box.x,
      y: box.y,
      w: box.w,
      h: Math.max(textHeight + 0.12, 0.35),
      rectRadius: 0.05,
      line: { color: "FDE68A", width: 1 },
      fill: { color: "FFFBEB" },
    });
  } else {
    slide.addShape("line", {
      x: box.x + 0.05,
      y: box.y,
      w: 0,
      h: Math.max(textHeight, 0.28),
      line: { color: "94A3B8", width: 2.25 },
    });
  }

  let usedHeight = 0;

  if (runs.length > 0) {
    usedHeight += renderSimpleText(slide, runs, {
      x: box.x + 0.18,
      y: box.y + 0.04,
      w: box.w - 0.22,
      h: Math.max(0, box.h - 0.04),
    }, {
      fontSize: BODY_FONT_SIZE,
      color: shaded ? "92400E" : "475569",
      italic: !shaded,
    });
  }

  if (children.length > 0) {
    usedHeight += await renderNodesInBox(slide, children, {
      x: box.x + 0.18,
      y: box.y + usedHeight + 0.02,
      w: box.w - 0.18,
      h: Math.max(0, box.h - usedHeight),
    }, 0);
  }

  return usedHeight + QUOTE_GAP;
}

function renderCodeBlock(slide: PptxSlide, source: string, language: string, box: LayoutBox): number {
  if (!source.trim() || box.h <= 0 || box.w <= 0) {
    return 0;
  }

  const codeText = language && language !== "plain text" ? `${language}\n${source}` : source;
  const height = Math.min(box.h, estimateTextHeight(codeText, 11, box.w, 0.6));

  slide.addShape("roundRect", {
    x: box.x,
    y: box.y,
    w: box.w,
    h: Math.max(height, 0.4),
    rectRadius: 0.04,
    line: { color: "CBD5E1", width: 1 },
    fill: { color: "F8FAFC" },
  });

  slide.addText(codeText, {
    x: box.x + 0.14,
    y: box.y + 0.08,
    w: box.w - 0.28,
    h: Math.max(0.2, height - 0.12),
    margin: 0,
    fontFace: CODE_FONT_FACE,
    fontSize: 11,
    color: "0F172A",
    breakLine: false,
    valign: "top",
    fit: "none",
    charSpacing: 0,
    lineSpacingMultiple: BODY_LINE_SPACING,
  });

  return height + BLOCK_GAP;
}

async function renderImageBlock(
  slide: PptxSlide,
  block: Extract<NotionBlockNode["block"], { type: "image" }>,
  box: LayoutBox,
): Promise<number> {
  const imageUrl = block.image.type === "external" ? block.image.external.url : block.image.file.url;
  const caption = renderPlainText(block.image.caption) || "Image";
  const image = await fetchImageData(imageUrl);

  if (!image) {
    return renderLinkBlock(slide, imageUrl, caption, box);
  }

  const requestedWidth = parseImageWidth(getImageWidthStyle(block), box.w);
  const maxWidth = Math.min(box.w, requestedWidth ?? box.w);
  const captionHeight = caption && caption !== "Image" ? 0.32 : 0;
  const maxHeight = Math.max(0.4, box.h - captionHeight);
  const scaled = scaleDimensions(image.width, image.height, maxWidth * PX_PER_INCH, maxHeight * PX_PER_INCH);
  const imageWidth = scaled.width / PX_PER_INCH;
  const imageHeight = scaled.height / PX_PER_INCH;

  slide.addImage({
    data: image.data,
    x: box.x,
    y: box.y,
    w: imageWidth,
    h: imageHeight,
    altText: caption,
  });

  let usedHeight = imageHeight;

  if (caption && caption !== "Image" && box.h - imageHeight > 0.2) {
    slide.addText(caption, {
      x: box.x,
      y: box.y + imageHeight + 0.04,
      w: Math.max(imageWidth, box.w * 0.5),
      h: 0.24,
      margin: 0,
      fontFace: BODY_FONT_FACE,
      fontSize: CAPTION_FONT_SIZE,
      color: "64748B",
      italic: true,
      charSpacing: 0,
    });
    usedHeight += 0.28;
  }

  return usedHeight + BLOCK_GAP;
}

function renderLinkBlock(slide: PptxSlide, url: string, label: string, box: LayoutBox): number {
  return renderSimpleText(slide, [{
    text: label,
    options: {
      hyperlink: { url, tooltip: url },
      color: "1D4ED8",
      underline: { color: "1D4ED8" },
    },
  }], box, {
    fontSize: BODY_FONT_SIZE,
  });
}

async function renderColumnListBlock(slide: PptxSlide, node: NotionBlockNode, box: LayoutBox): Promise<number> {
  const columns = node.children.filter((child) => child.block.type === "column");
  const otherChildren = node.children.filter((child) => child.block.type !== "column");

  if (columns.length === 0) {
    return renderNodesInBox(slide, otherChildren, box, 0);
  }

  if (columns.length === 1) {
    const onlyColumn = columns[0];
    const primaryHeight = onlyColumn ? await renderNodesInBox(slide, onlyColumn.children, box, 0) : 0;
    const trailingHeight = await renderNodesInBox(slide, otherChildren, {
      x: box.x,
      y: box.y + primaryHeight,
      w: box.w,
      h: Math.max(0, box.h - primaryHeight),
    }, 0);
    return primaryHeight + trailingHeight;
  }

  const widths = resolveColumnWidths(columns);
  let columnX = box.x;
  let maxHeight = 0;

  for (const [index, column] of columns.entries()) {
    const widthToken = widths[index];
    if (widthToken === undefined) {
      continue;
    }

    const width = Math.max(0.4, (Number.parseFloat(widthToken) / 100) * (box.w - COLUMN_GAP * (columns.length - 1)));
    const height = await renderNodesInBox(slide, column.children, {
      x: columnX,
      y: box.y,
      w: width,
      h: box.h,
    }, 0);
    maxHeight = Math.max(maxHeight, height);
    columnX += width + COLUMN_GAP;
  }

  const trailingHeight = await renderNodesInBox(slide, otherChildren, {
    x: box.x,
    y: box.y + maxHeight,
    w: box.w,
    h: Math.max(0, box.h - maxHeight),
  }, 0);

  return maxHeight + trailingHeight;
}

function renderTableBlock(slide: PptxSlide, node: NotionBlockNode, box: LayoutBox): number {
  const rows = node.children
    .filter((child) => child.block.type === "table_row")
    .map((child) => renderTableRowText(child));

  if (rows.length === 0) {
    return 0;
  }

  return renderSimpleText(slide, [{ text: rows.join("\n") }], box, {
    fontFace: CODE_FONT_FACE,
    fontSize: 12,
    color: "1F2937",
  });
}

function renderTableRowText(node: NotionBlockNode): string {
  if (node.block.type !== "table_row") {
    return "";
  }

  return node.block.table_row.cells.map((cell) => renderPlainText(cell) || " ").join(" | ");
}

function toRichTextRuns(richText: Parameters<typeof renderPlainText>[0]): RichTextRun[] {
  const runs: RichTextRun[] = [];

  for (const item of richText) {
    const text = item.type === "equation" ? `$${item.equation.expression}$` : item.plain_text;
    if (!text) {
      continue;
    }

    const options: Record<string, unknown> = {};

    if (item.annotations.bold) {
      options.bold = true;
    }
    if (item.annotations.italic) {
      options.italic = true;
    }
    if (item.annotations.strikethrough) {
      options.strike = true;
    }
    if (item.annotations.code) {
      options.fontFace = CODE_FONT_FACE;
      options.color = "7C2D12";
    }
    if (item.href) {
      options.hyperlink = { url: item.href, tooltip: item.href };
      options.color = options.color ?? "1D4ED8";
      options.underline = { color: options.color };
    }

    runs.push({
      text,
      options: Object.keys(options).length > 0 ? options : undefined,
    });
  }

  return runs;
}

function withPrefix(runs: RichTextRun[], prefix: string): RichTextRun[] {
  if (!prefix) {
    return runs;
  }

  return [{ text: prefix }, ...runs];
}

function estimateTextHeight(text: string, fontSize: number, width: number, density = 11): number {
  const estimatedLines = estimateWrappedLineCount(text, fontSize, width, density);
  const lineHeight = Math.max(0.28, fontSize * 0.0225);
  return estimatedLines * lineHeight + 0.03;
}

function estimateWrappedLineCount(text: string, fontSize: number, width: number, density = 11): number {
  const paragraphs = text.split("\n");
  const fallbackCharactersPerLine = Math.max(18, Math.floor(width * density));
  const availableWidthPoints = Math.max(width * 72, fontSize * 4);
  const averageCharacterWidth = Math.max(fontSize * 0.52, 1);
  const spaceWidth = Math.max(fontSize * 0.28, 1);
  let totalLines = 0;

  for (const paragraph of paragraphs) {
    const normalizedParagraph = paragraph.replace(/\s+/g, " ").trim();
    if (!normalizedParagraph) {
      totalLines += 1;
      continue;
    }

    const words = normalizedParagraph.split(" ");
    let currentLineWidth = 0;
    let paragraphLines = 1;

    for (const word of words) {
      const estimatedWordWidth = Math.min(word.length, fallbackCharactersPerLine) * averageCharacterWidth;
      const nextWidth = currentLineWidth === 0
        ? estimatedWordWidth
        : currentLineWidth + spaceWidth + estimatedWordWidth;

      if (currentLineWidth > 0 && nextWidth > availableWidthPoints) {
        paragraphLines += 1;
        currentLineWidth = estimatedWordWidth;
        continue;
      }

      currentLineWidth = nextWidth;
    }

    totalLines += paragraphLines;
  }

  return Math.max(1, totalLines);
}

function estimateParagraphHeight(paragraphs: Required<ParagraphSpec>[], width: number, fallbackFontSize: number): number {
  let totalHeight = 0;

  paragraphs.forEach((paragraph, index) => {
    const fontSize = typeof paragraph.options.fontSize === "number" ? paragraph.options.fontSize : fallbackFontSize;
    const indentLevel = typeof paragraph.options.indentLevel === "number" ? paragraph.options.indentLevel : 0;
    const text = paragraph.runs.map((run) => run.text).join("");
    const effectiveWidth = Math.max(0.6, width - indentLevel * 0.32);
    const paragraphHeight = estimateTextHeight(text, fontSize, effectiveWidth, 8.8);
    totalHeight += paragraphHeight;

    if (index < paragraphs.length - 1) {
      const paraSpaceAfter = typeof paragraph.options.paraSpaceAfter === "number" ? paragraph.options.paraSpaceAfter : PARAGRAPH_SPACE_AFTER;
      totalHeight += paraSpaceAfter / 72;
    }
  });

  return totalHeight;
}

async function fetchImageData(url: string): Promise<ImageData | null> {
  const response = await fetch(url);

  if (!response.ok) {
    return null;
  }

  const buffer = Buffer.from(await response.arrayBuffer());
  const size = imageSize(buffer);
  const contentType = resolveImageContentType(size.type, response.headers.get("content-type"));

  if (!contentType || !size.width || !size.height) {
    return null;
  }

  return {
    data: `data:${contentType};base64,${buffer.toString("base64")}`,
    width: size.width,
    height: size.height,
  };
}

function resolveImageContentType(type: string | undefined, headerValue: string | null): string | undefined {
  const normalizedHeader = headerValue?.split(";")[0]?.trim().toLowerCase();
  if (normalizedHeader?.startsWith("image/")) {
    return normalizedHeader;
  }

  switch (type) {
    case "png":
      return "image/png";
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "svg":
      return "image/svg+xml";
    default:
      return undefined;
  }
}

function scaleDimensions(width: number, height: number, maxWidth: number, maxHeight: number): { width: number; height: number } {
  const widthRatio = maxWidth / width;
  const heightRatio = maxHeight / height;
  const ratio = Math.min(widthRatio, heightRatio, 1);

  return {
    width: Math.round(width * ratio),
    height: Math.round(height * ratio),
  };
}

function resolveColumnWidths(columns: NotionBlockNode[]): string[] {
  const ratios = columns.map((column) => getColumnWidthRatio(column));
  const explicitTotal = ratios.reduce<number>((sum, value) => sum + (value ?? 0), 0);
  const missingCount = ratios.filter((value) => value === undefined).length;

  if (explicitTotal <= 0 && missingCount === columns.length) {
    return distributeEqualWidths(columns.length);
  }

  const remaining = Math.max(0, 1 - explicitTotal);
  const fallbackRatio = missingCount > 0 ? remaining / missingCount : 0;
  const filledRatios = ratios.map((value) => value ?? fallbackRatio);
  const filledTotal = filledRatios.reduce((sum, value) => sum + value, 0);

  if (filledTotal <= 0) {
    return distributeEqualWidths(columns.length);
  }

  return filledRatios.map((value) => formatWidthToken((value / filledTotal) * 100));
}

function distributeEqualWidths(columnCount: number): string[] {
  const baseWidth = 100 / columnCount;
  const widths = Array.from({ length: columnCount }, () => baseWidth);
  const rounded = widths.map((value) => formatWidthToken(value));
  const roundedTotal = rounded.reduce((sum, value) => sum + Number.parseFloat(value), 0);
  const remainder = 100 - roundedTotal;

  if (rounded.length > 0 && remainder !== 0) {
    const lastIndex = rounded.length - 1;
    const lastValue = rounded[lastIndex];

    if (lastValue !== undefined) {
      rounded[lastIndex] = formatWidthToken(Number.parseFloat(lastValue) + remainder);
    }
  }

  return rounded;
}

function getColumnWidthRatio(column: NotionBlockNode): number | undefined {
  if (column.block.type !== "column") {
    return undefined;
  }

  const widthRatio = (column.block.column as { width_ratio?: number }).width_ratio;
  return typeof widthRatio === "number" && Number.isFinite(widthRatio) && widthRatio > 0 ? widthRatio : undefined;
}

function formatWidthToken(value: number): string {
  return value.toFixed(3).replace(/\.0+$/, "").replace(/(\.\d*?)0+$/, "$1");
}

function getImageWidthStyle(block: Extract<NotionBlockNode["block"], { type: "image" }>): string | undefined {
  const imageData = block.image as {
    width?: number | string;
    display_width?: number | string;
    file?: { width?: number | string };
    external?: { width?: number | string };
  };
  const blockData = block as {
    width?: number | string;
    format?: { block_width?: number | string };
  };

  return normalizeWidthStyle(
    imageData.width ??
      imageData.display_width ??
      imageData.file?.width ??
      imageData.external?.width ??
      blockData.width ??
      blockData.format?.block_width,
  );
}

function normalizeWidthStyle(value: unknown): string | undefined {
  if (typeof value === "number") {
    return normalizeNumericWidth(value);
  }

  if (typeof value !== "string") {
    return undefined;
  }

  const trimmed = value.trim();
  if (!trimmed) {
    return undefined;
  }

  if (/^\d+(?:\.\d+)?px$/i.test(trimmed)) {
    return `${formatWidthToken(Number.parseFloat(trimmed))}px`;
  }

  if (/^\d+(?:\.\d+)?%$/.test(trimmed)) {
    return `${formatWidthToken(Number.parseFloat(trimmed))}%`;
  }

  if (/^\d+(?:\.\d+)?$/.test(trimmed)) {
    return normalizeNumericWidth(Number.parseFloat(trimmed));
  }

  return undefined;
}

function normalizeNumericWidth(value: number): string | undefined {
  if (!Number.isFinite(value) || value <= 0) {
    return undefined;
  }

  if (value <= 1) {
    return `${formatWidthToken(value * 100)}%`;
  }

  return `${formatWidthToken(value)}px`;
}

function parseImageWidth(widthStyle: string | undefined, availableWidth: number): number | undefined {
  if (!widthStyle) {
    return undefined;
  }

  if (widthStyle.endsWith("%")) {
    return (Number.parseFloat(widthStyle) / 100) * availableWidth;
  }

  if (widthStyle.endsWith("px")) {
    return Number.parseFloat(widthStyle) / PX_PER_INCH;
  }

  return undefined;
}