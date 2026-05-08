import { Buffer } from "node:buffer";
import { randomUUID } from "node:crypto";
import { existsSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { imageSize } from "image-size";
import {
  AlignmentType,
  BorderStyle,
  CharacterSet,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import type { ParagraphChild } from "docx";
import { Command } from "commander";
import puppeteer, { type Browser } from "puppeteer-core";

import {
  createNotionClient,
  fetchPageContent,
  renderPlainText,
  type NotionBlockNode,
  type NotionPageContent,
} from "./notion.ts";

type CliOptions = {
  page: string;
  output?: string;
};

type DocxImageType = "jpg" | "png" | "gif" | "bmp";

const MAX_IMAGE_WIDTH = 520;
const MAX_IMAGE_HEIGHT = 680;
const ORDERED_LIST_REFERENCE = "notion-ordered-list";
const DEFAULT_FONT_FAMILY = "Inter";
const ITALIC_FONT_FAMILY = "Inter Italic";
const DEFAULT_FONT_SIZE = 24;
const REGULAR_FONT_PATH = resolve("fonts/Inter-VariableFont_opsz,wght.ttf");
const ITALIC_FONT_PATH = resolve("fonts/Inter-Italic-VariableFont_opsz,wght.ttf");
const REGULAR_FONT_URL = pathToFileURL(REGULAR_FONT_PATH).href;
const ITALIC_FONT_URL = pathToFileURL(ITALIC_FONT_PATH).href;
const MERMAID_SCRIPT_PATH = resolve("node_modules/mermaid/dist/mermaid.min.js");
const MERMAID_BROWSER_PATHS = [
  process.env.GOOGLE_CHROME_BIN,
  process.env.PUPPETEER_EXECUTABLE_PATH,
  "/usr/bin/google-chrome",
  "/usr/bin/google-chrome-stable",
].filter((value): value is string => Boolean(value));
const MERMAID_VIEWPORT_WIDTH = 1600;
const MERMAID_VIEWPORT_HEIGHT = 1200;
const MERMAID_WRAPPING_WIDTH = 700;
const MERMAID_FONT_STACK = '"Trebuchet MS", Verdana, Arial, "Noto Color Emoji", sans-serif';
type DocxBlock = Paragraph | Table;
let mermaidBrowserPromise: Promise<Browser> | null = null;

async function main(): Promise<void> {
  try {
    const options = await parseArgs(process.argv);
    const notion = createNotionClient();
    const page = await fetchPageContent(notion, options.page);
    const outputPath = resolve(options.output ?? `${slugify(page.title)}.docx`);
    const document = await renderDocxDocument(page);
    const buffer = await Packer.toBuffer(document);

    await Bun.write(outputPath, buffer);

    console.log(`Created ${outputPath}`);
  } finally {
    await closeMermaidBrowser();
  }
}

async function parseArgs(argv: string[]): Promise<CliOptions> {
  const program = new Command();

  program
    .name("notion-to-docx")
    .description("Convert a Notion page into a DOCX document.")
    .argument("[page]", "Notion page ID or full page URL")
    .option("-p, --page <page>", "Notion page ID or full page URL")
    .option("-o, --out <path>", "Output DOCX path")
    .showHelpAfterError();

  await program.parseAsync(argv);

  const positionalPage = program.args[0];
  const options = program.opts<{ page?: string; out?: string }>();
  const page = options.page ?? positionalPage;

  if (!page || typeof page !== "string") {
    throw new Error("A Notion page ID or URL is required. Pass it as an argument or with --page.");
  }

  return {
    page,
    output: options.out,
  };
}

async function renderDocxDocument(page: NotionPageContent): Promise<Document> {
  const children = await renderNodes(page.blocks, 0);
  const fonts = await loadEmbeddedFonts();

  return new Document({
    fonts,
    styles: {
      default: {
        document: {
          run: {
            font: DEFAULT_FONT_FAMILY,
            size: DEFAULT_FONT_SIZE,
          },
        },
      },
      paragraphStyles: [
        {
          id: "Title",
          name: "Title",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: {
            font: DEFAULT_FONT_FAMILY,
            size: 34,
            bold: true,
            color: "111827",
          },
          paragraph: {
            spacing: {
              before: 0,
              after: 240,
            },
          },
        },
        {
          id: "Heading1",
          name: "Heading 1",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: {
            font: DEFAULT_FONT_FAMILY,
            size: 30,
            bold: true,
            color: "111827",
          },
          paragraph: {
            spacing: {
              before: 320,
              after: 140,
            },
          },
        },
        {
          id: "Heading2",
          name: "Heading 2",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: {
            font: DEFAULT_FONT_FAMILY,
            size: 28,
            bold: true,
            color: "111827",
          },
          paragraph: {
            spacing: {
              before: 280,
              after: 120,
            },
          },
        },
        {
          id: "Heading3",
          name: "Heading 3",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: {
            font: DEFAULT_FONT_FAMILY,
            size: 26,
            bold: true,
            color: "111827",
          },
          paragraph: {
            spacing: {
              before: 220,
              after: 100,
            },
          },
        },
      ],
    },
    numbering: {
      config: [
        {
          reference: ORDERED_LIST_REFERENCE,
          levels: Array.from({ length: 8 }, (_, level) => ({
            level,
            format: LevelFormat.DECIMAL,
            text: `${Array.from({ length: level + 1 }, (_unused, index) => `%${index + 1}`).join(".")}.`,
            alignment: AlignmentType.START,
            style: {
              paragraph: {
                indent: {
                  left: 720 * (level + 1),
                  hanging: 360,
                },
              },
            },
          })),
        },
      ],
    },
    sections: [
      {
        children: [
          new Paragraph({
            text: sanitizeText(page.title),
            heading: HeadingLevel.TITLE,
            spacing: {
              after: 240,
            },
          }),
          new Paragraph({
            spacing: {
              after: 180,
            },
            children: [
              new ExternalHyperlink({
                link: page.url,
                children: [
                  new TextRun({
                    text: page.url,
                    style: "Hyperlink",
                  }),
                ],
              }),
            ],
          }),
          new Paragraph({ text: "" }),
          ...children,
        ],
      },
    ],
  });
}

async function renderNodes(nodes: NotionBlockNode[], listDepth: number): Promise<DocxBlock[]> {
  const blocks: DocxBlock[] = [];

  for (const node of nodes) {
    blocks.push(...await renderNode(node, listDepth));
  }

  return blocks;
}

async function renderNode(node: NotionBlockNode, listDepth: number): Promise<DocxBlock[]> {
  const block = node.block;

  switch (block.type) {
    case "paragraph": {
      return appendChildren(
        renderParagraphBlock(block.paragraph.rich_text),
        await renderNodes(node.children, 0),
      );
    }
    case "heading_1": {
      return appendChildren(
        renderHeadingBlock(block.heading_1.rich_text, HeadingLevel.HEADING_1),
        await renderNodes(node.children, 0),
      );
    }
    case "heading_2": {
      return appendChildren(
        renderHeadingBlock(block.heading_2.rich_text, HeadingLevel.HEADING_2),
        await renderNodes(node.children, 0),
      );
    }
    case "heading_3": {
      return appendChildren(
        renderHeadingBlock(block.heading_3.rich_text, HeadingLevel.HEADING_3),
        await renderNodes(node.children, 0),
      );
    }
    case "bulleted_list_item": {
      return appendChildren(
        renderBulletBlock(block.bulleted_list_item.rich_text, listDepth),
        await renderNodes(node.children, listDepth + 1),
      );
    }
    case "numbered_list_item": {
      return appendChildren(
        renderNumberedBlock(block.numbered_list_item.rich_text, listDepth),
        await renderNodes(node.children, listDepth + 1),
      );
    }
    case "to_do": {
      const marker = block.to_do.checked ? "☑ " : "☐ ";
      return appendChildren(
        renderParagraphBlock(block.to_do.rich_text, { prefix: marker }),
        await renderNodes(node.children, listDepth + 1),
      );
    }
    case "toggle": {
      return appendChildren(
        renderBulletBlock(block.toggle.rich_text, listDepth),
        await renderNodes(node.children, listDepth + 1),
      );
    }
    case "quote": {
      return appendChildren(
        renderQuoteBlock(block.quote.rich_text),
        await renderNodes(node.children, 0),
      );
    }
    case "callout": {
      const emoji = block.callout.icon?.type === "emoji" ? `${block.callout.icon.emoji} ` : "";
      return appendChildren(
        renderQuoteBlock(block.callout.rich_text, emoji),
        await renderNodes(node.children, 0),
      );
    }
    case "code": {
      const text = renderPlainText(block.code.rich_text);

      if (isMermaidLanguage(block.code.language)) {
        return renderMermaidBlock(text);
      }

      return renderCodeBlock(text);
    }
    case "image": {
      return renderImageBlock(node);
    }
    case "bookmark": {
      return renderLinkParagraphs(block.bookmark.url, block.bookmark.url);
    }
    case "embed": {
      return renderLinkParagraphs(block.embed.url, block.embed.url);
    }
    case "video": {
      const videoUrl = block.video.type === "external" ? block.video.external.url : block.video.file.url;
      return renderLinkParagraphs(videoUrl, "Video");
    }
    case "file": {
      const fileUrl = block.file.type === "external" ? block.file.external.url : block.file.file.url;
      const caption = renderPlainText(block.file.caption) || "File";
      return renderLinkParagraphs(fileUrl, caption);
    }
    case "equation": {
      return [new Paragraph({ text: sanitizeText(`$${block.equation.expression}$`) })];
    }
    case "divider": {
      return [new Paragraph({ thematicBreak: true })];
    }
    case "child_page": {
      return [new Paragraph({ text: sanitizeText(block.child_page.title), heading: HeadingLevel.HEADING_1 })];
    }
    case "column_list": {
      const columns = node.children.filter((child) => child.block.type === "column");
      const otherChildren = node.children.filter((child) => child.block.type !== "column");
      const columnParagraphs = await renderNodes(columns.flatMap((column) => column.children), 0);
      const trailingParagraphs = await renderNodes(otherChildren, 0);
      return [...columnParagraphs, ...trailingParagraphs];
    }
    case "column": {
      return renderNodes(node.children, 0);
    }
    case "table": {
      return renderTableBlock(node);
    }
    case "table_row": {
      return [];
    }
    case "table_of_contents": {
      return [];
    }
    default: {
      return renderNodes(node.children, 0);
    }
  }
}

function renderTableBlock(node: NotionBlockNode): Table[] {
  const block = node.block;

  if (block.type !== "table") {
    return [];
  }

  const rows = node.children
    .filter((child) => child.block.type === "table_row")
    .map((child, rowIndex) => renderTableRow(child, {
      isHeaderRow: block.table.has_column_header && rowIndex === 0,
      hasRowHeader: block.table.has_row_header,
    }));

  if (rows.length === 0) {
    return [];
  }

  return [
    new Table({
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      layout: TableLayoutType.AUTOFIT,
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: "D1D5DB" },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: "D1D5DB" },
        left: { style: BorderStyle.SINGLE, size: 4, color: "D1D5DB" },
        right: { style: BorderStyle.SINGLE, size: 4, color: "D1D5DB" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "D1D5DB" },
        insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "D1D5DB" },
      },
      rows,
    }),
  ];
}

function renderTableRow(
  node: NotionBlockNode,
  options: { isHeaderRow: boolean; hasRowHeader: boolean },
): TableRow {
  const block = node.block;

  if (block.type !== "table_row") {
    return new TableRow({ children: [] });
  }

  return new TableRow({
    tableHeader: options.isHeaderRow,
    children: block.table_row.cells.map((cell, cellIndex) =>
      renderTableCell(cell, {
        isHeader: options.isHeaderRow || (options.hasRowHeader && cellIndex === 0),
      })),
  });
}

function renderTableCell(
  richText: Parameters<typeof renderPlainText>[0],
  options: { isHeader: boolean },
): TableCell {
  const children = renderRichTextChildren(richText, "", options.isHeader);

  return new TableCell({
    width: {
      size: 100,
      type: WidthType.AUTO,
    },
    margins: {
      top: 90,
      bottom: 90,
      left: 120,
      right: 120,
    },
    shading: options.isHeader ? { fill: "F3F4F6" } : undefined,
    children: [
      new Paragraph({
        spacing: {
          before: 0,
          after: 0,
        },
        children: children.length > 0 ? children : [new TextRun("")],
      }),
    ],
  });
}

function renderParagraphBlock(
  richText: Parameters<typeof renderPlainText>[0],
  options?: { prefix?: string },
): Paragraph[] {
  const children = renderRichTextChildren(richText, options?.prefix);
  return children.length > 0
    ? [
        new Paragraph({
          spacing: {
            after: 120,
          },
          children,
        }),
      ]
    : [];
}

function renderHeadingBlock(
  richText: Parameters<typeof renderPlainText>[0],
  heading: (typeof HeadingLevel)[keyof typeof HeadingLevel],
): Paragraph[] {
  const children = renderRichTextChildren(richText);
  return children.length > 0
    ? [
        new Paragraph({
          heading,
          style: resolveHeadingStyle(heading),
          spacing: resolveHeadingSpacing(heading),
          children,
        }),
      ]
    : [];
}

function renderBulletBlock(richText: Parameters<typeof renderPlainText>[0], level: number): Paragraph[] {
  const children = renderRichTextChildren(richText);
  return children.length > 0
    ? [
        new Paragraph({
          bullet: { level: clampListLevel(level) },
          children,
        }),
      ]
    : [];
}

function renderNumberedBlock(richText: Parameters<typeof renderPlainText>[0], level: number): Paragraph[] {
  const children = renderRichTextChildren(richText);
  return children.length > 0
    ? [
        new Paragraph({
          numbering: {
            reference: ORDERED_LIST_REFERENCE,
            level: clampListLevel(level),
          },
          children,
        }),
      ]
    : [];
}

function renderQuoteBlock(
  richText: Parameters<typeof renderPlainText>[0],
  prefix = "",
): Paragraph[] {
  const children = renderRichTextChildren(richText, prefix);
  return children.length > 0
    ? [
        new Paragraph({
          indent: { left: 720 },
          border: {
            left: {
              color: "CFCFCF",
              size: 12,
              space: 8,
              style: BorderStyle.SINGLE,
            },
          },
          spacing: {
            before: 120,
            after: 120,
          },
          children,
        }),
      ]
    : [];
}

function renderCodeBlock(text: string): Paragraph[] {
  const lines = sanitizeText(text).split("\n");

  if (lines.length === 0) {
    return [];
  }

  return lines.map((line) =>
    new Paragraph({
      indent: { left: 720 },
      spacing: { before: 0, after: 0 },
      children: [
        new TextRun({
          text: line,
          font: "Courier New",
        }),
      ],
    }),
  );
}

async function renderMermaidBlock(source: string): Promise<Paragraph[]> {
  const definition = source.trim();

  if (!definition) {
    return [];
  }

  try {
    const image = await renderMermaidImageRun(definition);

    if (!image) {
      return renderCodeBlock(source);
    }

    return [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: {
          before: 120,
          after: 120,
        },
        children: [image],
      }),
    ];
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : String(error);
    console.warn(`Unable to render Mermaid diagram in DOCX output: ${message}`);
    return renderCodeBlock(source);
  }
}

async function renderMermaidImageRun(source: string): Promise<ImageRun | null> {
  const diagram = await renderMermaidPng(source);
  const { width, height, png } = diagram;

  if (!width || !height) {
    return null;
  }

  const dimensions = scaleDimensions(width, height, MAX_IMAGE_WIDTH, MAX_IMAGE_HEIGHT);
  const altText = "Mermaid flowchart";

  return new ImageRun({
    type: "png",
    data: png,
    transformation: dimensions,
    altText: {
      title: altText,
      description: altText,
      name: altText,
    },
  });
}

async function renderMermaidPng(source: string): Promise<{ png: Buffer; width: number; height: number }> {
  if (!existsSync(MERMAID_SCRIPT_PATH)) {
    throw new Error(`Mermaid bundle not found at ${MERMAID_SCRIPT_PATH}. Run bun install first.`);
  }

  const normalizedSource = normalizeMermaidHtmlLabels(source);

  const browser = await getMermaidBrowser();
  const page = await browser.newPage();

  try {
    await page.setViewport({
      width: MERMAID_VIEWPORT_WIDTH,
      height: MERMAID_VIEWPORT_HEIGHT,
      deviceScaleFactor: 2,
    });
    await page.setContent(`<!doctype html><html><head><style>${buildMermaidFontCss()}</style></head><body><div id="container"></div></body></html>`);
    await page.addScriptTag({ path: MERMAID_SCRIPT_PATH });
    await page.evaluate(async () => {
      await document.fonts.ready;
    });

    await page.evaluate(async ({ definition, id, fontFamily, wrappingWidth }) => {
      const mermaid = (globalThis as { mermaid?: { initialize: (config: unknown) => void; render: (id: string, text: string, container?: Element | null) => Promise<{ svg: string }> } }).mermaid;

      if (!mermaid) {
        throw new Error("Mermaid did not load in the headless browser page.");
      }

      mermaid.initialize({
        startOnLoad: false,
        securityLevel: "loose",
        theme: "default",
        fontFamily,
        flowchart: {
          htmlLabels: true,
          useMaxWidth: false,
          wrappingWidth,
        },
      });

      const container = document.getElementById("container");
      const { svg } = await mermaid.render(id, definition, container);
      if (container) {
        container.innerHTML = svg;
      }
    }, {
      definition: normalizedSource,
      id: `mermaid-${randomUUID()}`,
      fontFamily: MERMAID_FONT_STACK,
      wrappingWidth: MERMAID_WRAPPING_WIDTH,
    });

    const diagram = await page.$("#container svg");

    if (!diagram) {
      throw new Error("Mermaid did not render an SVG element.");
    }

    const box = await diagram.boundingBox();

    if (!box || !box.width || !box.height) {
      throw new Error("Mermaid rendered an empty diagram.");
    }

    const png = await diagram.screenshot({
      type: "png",
      omitBackground: false,
    });

    return {
      png: Buffer.from(png),
      width: Math.round(box.width),
      height: Math.round(box.height),
    };
  } finally {
    await page.close();
  }
}

async function renderImageBlock(node: NotionBlockNode): Promise<Paragraph[]> {
  const block = node.block;

  if (block.type !== "image") {
    return [];
  }

  const imageUrl = block.image.type === "external" ? block.image.external.url : block.image.file.url;
  const caption = renderPlainText(block.image.caption) || "Image";
  const image = await fetchImageRun(imageUrl, caption);

  if (!image) {
    return renderLinkParagraphs(imageUrl, caption);
  }

  const paragraphs = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [image],
    }),
  ];

  if (caption !== "Image") {
    paragraphs.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: sanitizeText(caption),
            italics: true,
          }),
        ],
      }),
    );
  }

  return paragraphs;
}

async function fetchImageRun(url: string, altText: string): Promise<ImageRun | null> {
  const response = await fetch(url);

  if (!response.ok) {
    return null;
  }

  const data = Buffer.from(await response.arrayBuffer());
  const size = imageSize(data);
  const imageType = resolveDocxImageType(size.type, response.headers.get("content-type"));

  if (!imageType || !size.width || !size.height) {
    return null;
  }

  const dimensions = scaleDimensions(size.width, size.height, MAX_IMAGE_WIDTH, MAX_IMAGE_HEIGHT);

  return new ImageRun({
    type: imageType,
    data,
    transformation: dimensions,
    altText: {
      title: sanitizeText(altText),
      description: sanitizeText(altText),
      name: sanitizeText(altText),
    },
  });
}

async function getMermaidBrowser(): Promise<Browser> {
  if (!mermaidBrowserPromise) {
    const executablePath = resolveMermaidBrowserPath();

    mermaidBrowserPromise = puppeteer.launch({
      executablePath,
      headless: true,
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    });
  }

  return mermaidBrowserPromise;
}

async function closeMermaidBrowser(): Promise<void> {
  if (!mermaidBrowserPromise) {
    return;
  }

  const browserPromise = mermaidBrowserPromise;
  mermaidBrowserPromise = null;
  const browser = await browserPromise;
  await browser.close();
}

function resolveMermaidBrowserPath(): string {
  const executablePath = MERMAID_BROWSER_PATHS.find((value) => existsSync(value));

  if (!executablePath) {
    throw new Error(
      "Could not find a Chrome executable for Mermaid rendering. Set GOOGLE_CHROME_BIN or PUPPETEER_EXECUTABLE_PATH.",
    );
  }

  return executablePath;
}

function isMermaidLanguage(language: string): boolean {
  return language.trim().toLowerCase() === "mermaid";
}

function buildMermaidFontCss(): string {
  return `
    @font-face {
      font-family: "${DEFAULT_FONT_FAMILY}";
      src: url("${REGULAR_FONT_URL}") format("truetype");
      font-style: normal;
      font-weight: 100 900;
      font-display: block;
    }

    @font-face {
      font-family: "${DEFAULT_FONT_FAMILY}";
      src: url("${ITALIC_FONT_URL}") format("truetype");
      font-style: italic;
      font-weight: 100 900;
      font-display: block;
    }

    html, body, #container {
      margin: 0;
      padding: 0;
      font-family: ${MERMAID_FONT_STACK};
    }

    #container .nodeLabel,
    #container .edgeLabel {
      max-width: none !important;
    }

    #container .nodeLabel p,
    #container .edgeLabel p {
      margin: 0;
      white-space: nowrap !important;
      word-spacing: -0.08em;
      letter-spacing: normal;
    }
  `;
}

function normalizeMermaidHtmlLabels(source: string): string {
  return source.replace(/"([^"\\]*(?:\\.[^"\\]*)*)"/g, (match, value: string) => {
    if (!looksLikeMarkdownLabel(value)) {
      return match;
    }

    const unescaped = value.replace(/\\"/g, '"');
    return `"${convertMarkdownLabelToHtml(unescaped)}"`;
  });
}

function looksLikeMarkdownLabel(value: string): boolean {
  return /\*\*|__|~~|\*[^*]+\*|_[^_]+_/.test(value);
}

function convertMarkdownLabelToHtml(value: string): string {
  return escapeMermaidHtml(value)
    .replace(/\*\*([^*]+)\*\*/g, "<b>$1</b>")
    .replace(/__([^_]+)__/g, "<b>$1</b>")
    .replace(/~~([^~]+)~~/g, "<s>$1</s>")
    .replace(/\*([^*]+)\*/g, "<i>$1</i>")
    .replace(/_([^_]+)_/g, "<i>$1</i>")
    .replace(/\n/g, "<br/>")
    .replace(/"/g, "&quot;");
}

function escapeMermaidHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function renderLinkParagraphs(url: string, label: string): Paragraph[] {
  return [
    new Paragraph({
      spacing: {
        after: 120,
      },
      children: [
        new ExternalHyperlink({
          link: url,
          children: [
            new TextRun({
              text: sanitizeText(label),
              style: "Hyperlink",
            }),
          ],
        }),
      ],
    }),
  ];
}

function renderRichTextChildren(
  richText: Parameters<typeof renderPlainText>[0],
  prefix = "",
  forceBold = false,
): ParagraphChild[] {
  const children: ParagraphChild[] = [];

  if (prefix) {
    children.push(
      new TextRun({
        text: sanitizeText(prefix),
        font: DEFAULT_FONT_FAMILY,
      }),
    );
  }

  for (const item of richText) {
    const value = sanitizeText(item.type === "equation" ? `$${item.equation.expression}$` : item.plain_text);

    if (!value) {
      continue;
    }

    const textRun = new TextRun({
      text: value,
      bold: forceBold || item.annotations.bold,
      italics: item.annotations.italic,
      strike: item.annotations.strikethrough,
      font: resolveTextRunFont(item.annotations.code, item.annotations.italic),
    });

    if (item.href) {
      children.push(
        new ExternalHyperlink({
          link: item.href,
          children: [
            new TextRun({
              text: value,
              bold: forceBold || item.annotations.bold,
              italics: item.annotations.italic,
              strike: item.annotations.strikethrough,
              font: resolveTextRunFont(item.annotations.code, item.annotations.italic),
              style: "Hyperlink",
            }),
          ],
        }),
      );
      continue;
    }

    children.push(textRun);
  }

  return children;
}

function appendChildren(blocks: DocxBlock[], children: DocxBlock[]): DocxBlock[] {
  return [...blocks, ...children];
}

async function loadEmbeddedFonts(): Promise<Array<{
  name: string;
  data: Buffer;
  characterSet: (typeof CharacterSet)[keyof typeof CharacterSet];
}>> {
  const [regularFont, italicFont] = await Promise.all([
    readFile(REGULAR_FONT_PATH),
    readFile(ITALIC_FONT_PATH),
  ]);

  return [
    {
      name: DEFAULT_FONT_FAMILY,
      data: regularFont,
      characterSet: CharacterSet.ANSI,
    },
    {
      name: ITALIC_FONT_FAMILY,
      data: italicFont,
      characterSet: CharacterSet.ANSI,
    },
  ];
}

function resolveTextRunFont(isCode: boolean, isItalic: boolean): string | undefined {
  if (isCode) {
    return "Courier New";
  }

  return isItalic ? ITALIC_FONT_FAMILY : DEFAULT_FONT_FAMILY;
}

function resolveHeadingStyle(heading: (typeof HeadingLevel)[keyof typeof HeadingLevel]): string | undefined {
  switch (heading) {
    case HeadingLevel.TITLE:
      return "Title";
    case HeadingLevel.HEADING_1:
      return "Heading1";
    case HeadingLevel.HEADING_2:
      return "Heading2";
    case HeadingLevel.HEADING_3:
      return "Heading3";
    default:
      return undefined;
  }
}

function resolveHeadingSpacing(
  heading: (typeof HeadingLevel)[keyof typeof HeadingLevel],
): { before?: number; after?: number } | undefined {
  switch (heading) {
    case HeadingLevel.TITLE:
      return { before: 0, after: 240 };
    case HeadingLevel.HEADING_1:
      return { before: 320, after: 140 };
    case HeadingLevel.HEADING_2:
      return { before: 280, after: 120 };
    case HeadingLevel.HEADING_3:
      return { before: 220, after: 100 };
    default:
      return undefined;
  }
}

function clampListLevel(level: number): number {
  return Math.max(0, Math.min(7, level));
}

function resolveDocxImageType(type: string | undefined, contentType: string | null): DocxImageType | null {
  const normalizedType = type?.toLowerCase();

  switch (normalizedType) {
    case "jpg":
    case "jpeg":
      return "jpg";
    case "png":
    case "gif":
    case "bmp":
      return normalizedType;
    default:
      return resolveImageTypeFromContentType(contentType);
  }
}

function resolveImageTypeFromContentType(contentType: string | null): DocxImageType | null {
  const normalized = contentType?.split(";")[0]?.trim().toLowerCase();

  switch (normalized) {
    case "image/jpeg":
    case "image/jpg":
      return "jpg";
    case "image/png":
      return "png";
    case "image/gif":
      return "gif";
    case "image/bmp":
      return "bmp";
    default:
      return null;
  }
}

function scaleDimensions(
  width: number,
  height: number,
  maxWidth: number,
  maxHeight: number,
): { width: number; height: number } {
  const ratio = Math.min(maxWidth / width, maxHeight / height, 1);

  return {
    width: Math.max(1, Math.round(width * ratio)),
    height: Math.max(1, Math.round(height * ratio)),
  };
}

function sanitizeText(value: string): string {
  return value.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");
}

function slugify(value: string): string {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "") || "notion-document";
}

main().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : String(error);
  console.error(message);
  process.exit(1);
});