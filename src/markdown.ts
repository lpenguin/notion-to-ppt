import type { NotionBlockNode, NotionPageContent } from "./notion.ts";
import { renderPlainText, renderRichTextMarkdown } from "./notion.ts";

type Slide = {
  title?: string;
  lines: string[];
};

export function renderDeckMarkdown(page: NotionPageContent): string {
  const slides: Slide[] = [{ title: page.title, lines: [] }];

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

    currentSlide.lines.push(...renderBlock(node, 0));
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
  const indent = "  ".repeat(depth);
  const lines: string[] = [];

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
      lines.push(`![${caption}](${imageUrl})`);
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
    default:
      break;
  }

  for (const child of node.children) {
    const childDepth = isListLike(block.type) ? depth + 1 : depth;
    lines.push(...renderBlock(child, childDepth));
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