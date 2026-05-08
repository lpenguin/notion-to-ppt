import type { NotionBlockNode, NotionPageContent } from "./notion.ts";
import { hasExtendedRichTextBlock, renderPlainText, renderRichTextMarkdown } from "./notion.ts";

type Slide = {
  title?: string;
  lines: string[];
};

const COLUMN_GAP_REM = 1.2;

export function renderDeckMarkdown(page: NotionPageContent): string {
  const slides: Slide[] = [{ title: page.title, lines: [] }];
  const columnWidthTokens = collectColumnWidthTokens(page.blocks);

  for (const node of page.blocks) {
    const currentSlide = getCurrentSlide(slides);

    if (node.block.type === "heading_1") {
      finalizeSlide(currentSlide);
      slides.push({
        title: renderRichTextMarkdown(node.block.heading_1.rich_text) || page.title,
        lines: [],
      });
      continue;
    }

    if (node.block.type === "divider") {
      finalizeSlide(currentSlide);
      slides.push({ lines: [] });
      continue;
    }

    appendBlockLines(currentSlide.lines, renderBlock(node, 0), node.block.type);
  }

  for (const slide of slides) {
    finalizeSlide(slide);
  }

  const nonEmptySlides = slides.filter((slide) => slide.title || slide.lines.length > 0);

  return [
    "---",
    "marp: true",
    "theme: default",
    "size: 16:9",
    "paginate: true",
    `title: ${escapeYaml(page.title)}`,
    "style: |",
    "  section .notion-columns {",
    "    display: flex;",
    `    gap: ${COLUMN_GAP_REM}rem;`,
    "    align-items: flex-start;",
    "  }",
    "  section .notion-column {",
    "    flex: 1 1 0;",
    "    min-width: 0;",
    "  }",
    ...renderColumnWidthStyles(columnWidthTokens),
    "  section .notion-column > :first-child {",
    "    margin-top: 0;",
    "  }",
    "  section:has(> p > img),",
    "  section:has(> figure) {",
    "    display: flex;",
    "    flex-direction: column;",
    "    align-items: stretch;",
    "  }",
    "  section > p:has(> img),",
    "  section > figure {",
    "    flex: 1 1 auto;",
    "    min-height: 0;",
    "    display: flex;",
    "    align-items: center;",
    "    justify-content: center;",
    "    margin-bottom: 0;",
    "  }",
    "  section > figure {",
    "    margin-left: 0;",
    "    margin-right: 0;",
    "  }",
    "  section li:has(img) {",
    "    display: flex;",
    "    flex-direction: column;",
    "    gap: 0.5rem;",
    "  }",
    "  section li img {",
    "    align-self: center;",
    "    max-height: 24vh;",
    "  }",
    "  section img {",
    "    max-width: 100%;",
    "    max-height: 100%;",
    "    width: auto;",
    "    height: auto;",
    "    object-fit: contain;",
    "  }",
    "  section .notion-column img {",
    "    max-height: 100%;",
    "  }",
    "---",
    "",
    nonEmptySlides
      .map((slide) => {
        const parts: string[] = [];
        if (slide.title) {
          parts.push(`# ${slide.title}`);
        }
        if (slide.lines.length > 0) {
          parts.push(slide.lines.join("\n"));
        }
        return parts.join("\n\n").trim();
      })
      .join("\n\n---\n\n"),
    "",
  ].join("\n");
}

function renderBlock(node: NotionBlockNode, depth: number): string[] {
  const block = node.block;
  const runtimeBlock = block as { type: string } & Record<string, unknown>;
  const indent = "  ".repeat(depth);
  const lines: string[] = [];

  if (hasExtendedRichTextBlock(runtimeBlock, "heading_4")) {
    const text = renderRichTextMarkdown(runtimeBlock.heading_4.rich_text);
    if (text) {
      lines.push(`#### ${text}`);
    }

    for (const child of node.children) {
      appendBlockLines(lines, renderBlock(child, depth), child.block.type);
    }

    return compactLines(lines);
  }

  switch (block.type) {
    case "paragraph": {
      const text = renderRichTextMarkdown(block.paragraph.rich_text);
      if (text) {
        lines.push(text);
      }
      break;
    }
    case "heading_2": {
      const text = renderRichTextMarkdown(block.heading_2.rich_text);
      if (text) {
        lines.push(`## ${text}`);
      }
      break;
    }
    case "heading_3": {
      const text = renderRichTextMarkdown(block.heading_3.rich_text);
      if (text) {
        lines.push(`### ${text}`);
      }
      break;
    }
    case "bulleted_list_item": {
      const text = renderRichTextMarkdown(block.bulleted_list_item.rich_text);
      if (text) {
        lines.push(`${indent}- ${text}`);
      }
      break;
    }
    case "numbered_list_item": {
      const text = renderRichTextMarkdown(block.numbered_list_item.rich_text);
      if (text) {
        lines.push(`${indent}1. ${text}`);
      }
      break;
    }
    case "to_do": {
      const text = renderRichTextMarkdown(block.to_do.rich_text);
      if (text) {
        lines.push(`${indent}- [${block.to_do.checked ? "x" : " "}] ${text}`);
      }
      break;
    }
    case "quote": {
      const text = renderRichTextMarkdown(block.quote.rich_text);
      if (text) {
        lines.push(`${indent}> ${text}`);
      }
      break;
    }
    case "callout": {
      const emoji = block.callout.icon?.type === "emoji" ? `${block.callout.icon.emoji} ` : "";
      const text = renderRichTextMarkdown(block.callout.rich_text);
      if (text) {
        lines.push(`${indent}> ${emoji}${text}`);
      }
      break;
    }
    case "toggle": {
      const text = renderRichTextMarkdown(block.toggle.rich_text);
      if (text) {
        lines.push(`${indent}- ${text}`);
      }
      break;
    }
    case "code": {
      const text = renderPlainText(block.code.rich_text);
      const language = block.code.language === "plain text" ? "text" : block.code.language;
      lines.push(`\`\`\`${language}`);
      lines.push(text);
      lines.push("\`\`\`");
      break;
    }
    case "image": {
      const imageUrl = block.image.type === "external" ? block.image.external.url : block.image.file.url;
      const caption = renderPlainText(block.image.caption) || "Image";
      const widthStyle = getImageWidthStyle(block);
      lines.push(widthStyle ? renderImageHtml(imageUrl, caption, widthStyle) : `![${caption}](${imageUrl})`);
      break;
    }
    case "bookmark": {
      lines.push(`[${block.bookmark.url}](${block.bookmark.url})`);
      break;
    }
    case "embed": {
      lines.push(`[Embedded content](${block.embed.url})`);
      break;
    }
    case "video": {
      const videoUrl = block.video.type === "external" ? block.video.external.url : block.video.file.url;
      lines.push(`[Video](${videoUrl})`);
      break;
    }
    case "file": {
      const fileUrl = block.file.type === "external" ? block.file.external.url : block.file.file.url;
      const caption = renderPlainText(block.file.caption) || "File";
      lines.push(`[${caption}](${fileUrl})`);
      break;
    }
    case "equation": {
      lines.push(`$$${block.equation.expression}$$`);
      break;
    }
    case "table_of_contents": {
      lines.push("<!-- table of contents omitted -->");
      break;
    }
    case "child_page": {
      lines.push(`## ${block.child_page.title}`);
      break;
    }
    case "column_list": {
      lines.push(...renderColumnList(node, depth));
      break;
    }
    case "column": {
      for (const child of node.children) {
        lines.push(...renderBlock(child, depth));
      }
      break;
    }
    default:
      break;
  }

  if (block.type === "column_list" || block.type === "column") {
    return compactLines(lines);
  }

  for (const child of node.children) {
    const childDepth = isListLike(block.type) ? depth + 1 : depth;
    appendBlockLines(lines, renderBlock(child, childDepth), child.block.type);
  }

  return compactLines(lines);
}

function isListLike(type: NotionBlockNode["block"]["type"]): boolean {
  return type === "bulleted_list_item" || type === "numbered_list_item" || type === "to_do" || type === "toggle";
}

function getCurrentSlide(slides: Slide[]): Slide {
  const currentSlide = slides[slides.length - 1];
  if (!currentSlide) {
    throw new Error("Expected an active slide while rendering the deck.");
  }
  return currentSlide;
}

function finalizeSlide(slide: Slide): void {
  slide.lines = compactLines(slide.lines);
}

function appendBlockLines(target: string[], blockLines: string[], blockType: NotionBlockNode["block"]["type"]): void {
  if (blockLines.length === 0) {
    return;
  }

  const previousLine = findLastNonEmptyLine(target);
  const needsBlockSeparation =
    blockType === "image" ||
    (!!previousLine && isListMarkdownLine(previousLine) && !isListLike(blockType));

  if (needsBlockSeparation && target.length > 0 && target[target.length - 1] !== "") {
    target.push("");
  }

  target.push(...blockLines);
}

function findLastNonEmptyLine(lines: string[]): string | undefined {
  for (let index = lines.length - 1; index >= 0; index -= 1) {
    const line = lines[index];
    if (line && line.trim().length > 0) {
      return line;
    }
  }

  return undefined;
}

function isListMarkdownLine(line: string): boolean {
  return /^(?:\s*)(?:[-*+] |\d+\. |- \[[ x]\] )/.test(line);
}

function compactLines(lines: string[]): string[] {
  const compacted: string[] = [];

  for (const line of lines) {
    const value = line.trimEnd();
    const previous = compacted[compacted.length - 1];

    if (!value) {
      if (previous) {
        compacted.push("");
      }
      continue;
    }

    compacted.push(value);
  }

  while (compacted[0] === "") {
    compacted.shift();
  }

  while (compacted[compacted.length - 1] === "") {
    compacted.pop();
  }

  return compacted;
}

function escapeYaml(value: string): string {
  return JSON.stringify(value);
}

function renderColumnList(node: NotionBlockNode, depth: number): string[] {
  const columns = node.children.filter((child) => child.block.type === "column");
  const nonColumnChildren = node.children.filter((child) => child.block.type !== "column");

  if (columns.length === 0) {
    return nonColumnChildren.flatMap((child) => renderBlock(child, depth));
  }

  if (columns.length === 1) {
    const onlyColumn = columns[0];

    if (!onlyColumn) {
      return nonColumnChildren.flatMap((child) => renderBlock(child, depth));
    }

    return [
      ...onlyColumn.children.flatMap((child) => renderBlock(child, depth)),
      ...nonColumnChildren.flatMap((child) => renderBlock(child, depth)),
    ];
  }

  const widths = resolveColumnWidths(columns);
  const lines = [
    '<div class="notion-columns">',
  ];

  columns.forEach((column, index) => {
    const width = widths[index];
    if (width === undefined) {
      return;
    }

    const columnHtml = renderColumnHtml(column.children, width);
    if (columnHtml) {
      lines.push(columnHtml);
    }
  });

  lines.push("</div>");

  for (const child of nonColumnChildren) {
    lines.push(...renderBlock(child, depth));
  }

  return compactLines(lines);
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

function renderColumnHtml(nodes: NotionBlockNode[], width: string): string {
  const html = nodes
    .map((child) => renderBlockHtml(child))
    .filter((value) => value.length > 0)
    .join("\n");

  if (!html) {
    return "";
  }

  return [
    `<div class="notion-column" data-width="${width}">`,
    html,
    "</div>",
  ].join("\n");
}

function renderBlockHtml(node: NotionBlockNode): string {
  const block = node.block;
  const runtimeBlock = block as { type: string } & Record<string, unknown>;

  if (hasExtendedRichTextBlock(runtimeBlock, "heading_4")) {
    const text = renderRichTextHtml(runtimeBlock.heading_4.rich_text);
    return text ? `<h4>${text}</h4>${renderChildHtml(node.children)}` : renderChildHtml(node.children);
  }

  switch (block.type) {
    case "paragraph": {
      const text = renderRichTextHtml(block.paragraph.rich_text);
      return text ? `<p>${text}</p>${renderChildHtml(node.children)}` : renderChildHtml(node.children);
    }
    case "heading_2": {
      const text = renderRichTextHtml(block.heading_2.rich_text);
      return text ? `<h2>${text}</h2>${renderChildHtml(node.children)}` : renderChildHtml(node.children);
    }
    case "heading_3": {
      const text = renderRichTextHtml(block.heading_3.rich_text);
      return text ? `<h3>${text}</h3>${renderChildHtml(node.children)}` : renderChildHtml(node.children);
    }
    case "bulleted_list_item": {
      return renderListItemHtml("ul", renderRichTextHtml(block.bulleted_list_item.rich_text), node.children);
    }
    case "numbered_list_item": {
      return renderListItemHtml("ol", renderRichTextHtml(block.numbered_list_item.rich_text), node.children);
    }
    case "to_do": {
      const text = renderRichTextHtml(block.to_do.rich_text);
      const marker = block.to_do.checked ? "&#x2611;" : "&#x2610;";
      return renderListItemHtml("ul", `${marker} ${text}`.trim(), node.children);
    }
    case "quote": {
      const text = renderRichTextHtml(block.quote.rich_text);
      return text ? `<blockquote><p>${text}</p>${renderChildHtml(node.children)}</blockquote>` : renderChildHtml(node.children);
    }
    case "callout": {
      const emoji = block.callout.icon?.type === "emoji" ? `${escapeHtml(block.callout.icon.emoji)} ` : "";
      const text = renderRichTextHtml(block.callout.rich_text);
      return text ? `<blockquote><p>${emoji}${text}</p>${renderChildHtml(node.children)}</blockquote>` : renderChildHtml(node.children);
    }
    case "toggle": {
      const text = renderRichTextHtml(block.toggle.rich_text);
      return renderListItemHtml("ul", text, node.children);
    }
    case "code": {
      const text = escapeHtml(renderPlainText(block.code.rich_text));
      return `<pre><code>${text}</code></pre>`;
    }
    case "image": {
      const imageUrl = block.image.type === "external" ? block.image.external.url : block.image.file.url;
      const caption = renderPlainText(block.image.caption) || "Image";
      return renderImageHtml(imageUrl, caption, getImageWidthStyle(block));
    }
    case "bookmark": {
      const url = escapeAttribute(block.bookmark.url);
      return `<p><a href="${url}">${escapeHtml(block.bookmark.url)}</a></p>`;
    }
    case "embed": {
      const url = escapeAttribute(block.embed.url);
      return `<p><a href="${url}">Embedded content</a></p>`;
    }
    case "video": {
      const videoUrl = block.video.type === "external" ? block.video.external.url : block.video.file.url;
      return `<p><a href="${escapeAttribute(videoUrl)}">Video</a></p>`;
    }
    case "file": {
      const fileUrl = block.file.type === "external" ? block.file.external.url : block.file.file.url;
      const caption = renderPlainText(block.file.caption) || "File";
      return `<p><a href="${escapeAttribute(fileUrl)}">${escapeHtml(caption)}</a></p>`;
    }
    case "equation": {
      return `<p>$$${escapeHtml(block.equation.expression)}$$</p>`;
    }
    case "table_of_contents": {
      return "<!-- table of contents omitted -->";
    }
    case "child_page": {
      return `<h2>${escapeHtml(block.child_page.title)}</h2>`;
    }
    case "column_list": {
      return renderColumnList(node, 0).join("\n");
    }
    case "column": {
      return renderChildHtml(node.children);
    }
    default:
      return renderChildHtml(node.children);
  }
}

function renderChildHtml(children: NotionBlockNode[]): string {
  return children
    .map((child) => renderBlockHtml(child))
    .filter((value) => value.length > 0)
    .join("\n");
}

function renderListItemHtml(tagName: "ul" | "ol", text: string, children: NotionBlockNode[]): string {
  const childHtml = renderChildHtml(children);
  const body = [text, childHtml].filter((value) => value.length > 0).join("\n");

  if (!body) {
    return "";
  }

  return `<${tagName}><li>${body}</li></${tagName}>`;
}

function renderImageHtml(imageUrl: string, caption: string, widthStyle?: string): string {
  const style = [widthStyle, "max-width:100%", "height:auto"].filter((value) => value && value.length > 0).join("; ");
  return `<figure><img src="${escapeAttribute(imageUrl)}" alt="${escapeAttribute(caption)}" style="${style}" />${renderFigureCaption(caption)}</figure>`;
}

function renderFigureCaption(caption: string): string {
  return caption && caption !== "Image" ? `<figcaption>${escapeHtml(caption)}</figcaption>` : "";
}

function renderRichTextHtml(richText: Parameters<typeof renderRichTextMarkdown>[0]): string {
  return richText
    .map((item) => {
      let text = item.type === "equation" ? `$${escapeHtml(item.equation.expression)}$` : escapeHtml(item.plain_text);

      if (item.annotations.code) {
        text = `<code>${text}</code>`;
      }
      if (item.annotations.bold) {
        text = `<strong>${text}</strong>`;
      }
      if (item.annotations.italic) {
        text = `<em>${text}</em>`;
      }
      if (item.annotations.strikethrough) {
        text = `<del>${text}</del>`;
      }
      if (item.href) {
        text = `<a href="${escapeAttribute(item.href)}">${text}</a>`;
      }

      return text;
    })
    .join("")
    .trim();
}

function escapeHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeAttribute(value: string): string {
  return escapeHtml(value);
}

function collectColumnWidthTokens(nodes: NotionBlockNode[]): string[] {
  const tokens = new Set<string>();

  const visit = (items: NotionBlockNode[]): void => {
    for (const node of items) {
      if (node.block.type === "column_list") {
        for (const token of resolveColumnWidths(node.children.filter((child) => child.block.type === "column"))) {
          tokens.add(token);
        }
      }

      visit(node.children);
    }
  };

  visit(nodes);
  return Array.from(tokens).sort((left, right) => Number.parseFloat(left) - Number.parseFloat(right));
}

function renderColumnWidthStyles(widthTokens: string[]): string[] {
  return widthTokens.flatMap((token) => [
    `  section .notion-column[data-width="${token}"] {`,
    `    flex: 0 0 ${token}%;`,
    `    width: ${token}%;`,
    `    max-width: ${token}%;`,
    "  }",
  ]);
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