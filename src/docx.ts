import { Buffer } from "node:buffer";
import { resolve } from "node:path";
import { imageSize } from "image-size";
import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  Packer,
  Paragraph,
  TextRun,
} from "docx";
import type { ParagraphChild } from "docx";
import { Command } from "commander";

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

async function main(): Promise<void> {
  const options = await parseArgs(process.argv);
  const notion = createNotionClient();
  const page = await fetchPageContent(notion, options.page);
  const outputPath = resolve(options.output ?? `${slugify(page.title)}.docx`);
  const document = await renderDocxDocument(page);
  const buffer = await Packer.toBuffer(document);

  await Bun.write(outputPath, buffer);

  console.log(`Created ${outputPath}`);
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

  return new Document({
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
          }),
          new Paragraph({
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

async function renderNodes(nodes: NotionBlockNode[], listDepth: number): Promise<Paragraph[]> {
  const paragraphs: Paragraph[] = [];

  for (const node of nodes) {
    paragraphs.push(...await renderNode(node, listDepth));
  }

  return paragraphs;
}

async function renderNode(node: NotionBlockNode, listDepth: number): Promise<Paragraph[]> {
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
      return renderCodeBlock(renderPlainText(block.code.rich_text));
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
    case "table_of_contents": {
      return [];
    }
    default: {
      return renderNodes(node.children, 0);
    }
  }
}

function renderParagraphBlock(
  richText: Parameters<typeof renderPlainText>[0],
  options?: { prefix?: string },
): Paragraph[] {
  const children = renderRichTextChildren(richText, options?.prefix);
  return children.length > 0 ? [new Paragraph({ children })] : [];
}

function renderHeadingBlock(
  richText: Parameters<typeof renderPlainText>[0],
  heading: (typeof HeadingLevel)[keyof typeof HeadingLevel],
): Paragraph[] {
  const children = renderRichTextChildren(richText);
  return children.length > 0 ? [new Paragraph({ heading, children })] : [];
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

function renderLinkParagraphs(url: string, label: string): Paragraph[] {
  return [
    new Paragraph({
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
): ParagraphChild[] {
  const children: ParagraphChild[] = [];

  if (prefix) {
    children.push(new TextRun({ text: sanitizeText(prefix) }));
  }

  for (const item of richText) {
    const value = sanitizeText(item.type === "equation" ? `$${item.equation.expression}$` : item.plain_text);

    if (!value) {
      continue;
    }

    const textRun = new TextRun({
      text: value,
      bold: item.annotations.bold,
      italics: item.annotations.italic,
      strike: item.annotations.strikethrough,
      font: item.annotations.code ? "Courier New" : undefined,
    });

    if (item.href) {
      children.push(
        new ExternalHyperlink({
          link: item.href,
          children: [
            new TextRun({
              text: value,
              bold: item.annotations.bold,
              italics: item.annotations.italic,
              strike: item.annotations.strikethrough,
              font: item.annotations.code ? "Courier New" : undefined,
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

function appendChildren(paragraphs: Paragraph[], children: Paragraph[]): Paragraph[] {
  return [...paragraphs, ...children];
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