import {
  Client,
  collectPaginatedAPI,
  isFullBlock,
  isFullPage,
} from "@notionhq/client";
import type {
  BlockObjectResponse,
  PageObjectResponse,
  RichTextItemResponse,
} from "@notionhq/client/build/src/api-endpoints";

export type NotionBlockNode = {
  block: BlockObjectResponse;
  children: NotionBlockNode[];
};

export type NotionPageContent = {
  id: string;
  title: string;
  url: string;
  blocks: NotionBlockNode[];
};

export type ExtendedRichTextBlock<TType extends string> = {
  type: TType;
} & Record<TType, { rich_text: RichTextItemResponse[] }>;

export function createNotionClient(token = process.env.NOTION_TOKEN): Client {
  if (!token) {
    throw new Error("Missing NOTION_TOKEN. Set it in your environment or .env file.");
  }

  return new Client({ auth: token });
}

export async function fetchPageContent(client: Client, pageInput: string): Promise<NotionPageContent> {
  const pageId = normalizePageId(pageInput);
  const page = await client.pages.retrieve({ page_id: pageId });

  if (!isFullPage(page)) {
    throw new Error(`Notion returned a partial page response for ${pageInput}.`);
  }

  const blocks = await fetchBlockTree(client, page.id);

  return {
    id: page.id,
    title: extractPageTitle(page),
    url: page.url,
    blocks,
  };
}

async function fetchBlockTree(client: Client, blockId: string): Promise<NotionBlockNode[]> {
  const blocks = await collectPaginatedAPI(client.blocks.children.list, {
    block_id: blockId,
    page_size: 100,
  });

  const fullBlocks = blocks.filter(isFullBlock);

  return Promise.all(
    fullBlocks.map(async (block) => ({
      block,
      children: block.has_children ? await fetchBlockTree(client, block.id) : [],
    })),
  );
}

function extractPageTitle(page: PageObjectResponse): string {
  for (const property of Object.values(page.properties) as Array<PageObjectResponse["properties"][string]>) {
    if (property.type === "title") {
      return renderPlainText(property.title) || "Untitled";
    }
  }

  return "Untitled";
}

export function renderPlainText(richText: RichTextItemResponse[]): string {
  return richText.map((item) => item.plain_text).join("").trim();
}

export function renderRichTextMarkdown(richText: RichTextItemResponse[]): string {
  return richText
    .map((item) => {
      let text = item.plain_text;

      if (item.type === "equation") {
        text = `$${item.equation.expression}$`;
      }

      if (item.href) {
        text = `[${text}](${item.href})`;
      }

      if (item.annotations.code) {
        text = `\`${text}\``;
      }
      if (item.annotations.bold) {
        text = `**${text}**`;
      }
      if (item.annotations.italic) {
        text = `*${text}*`;
      }
      if (item.annotations.strikethrough) {
        text = `~~${text}~~`;
      }

      return text;
    })
    .join("")
    .trim();
}

export function hasExtendedRichTextBlock<TType extends string, TBlock extends { type: string } & Record<string, unknown>>(
  block: TBlock,
  type: TType,
): block is TBlock & ExtendedRichTextBlock<TType> {
  if (block.type !== type) {
    return false;
  }

  const value = (block as Record<string, unknown>)[type];
  return typeof value === "object" && value !== null && Array.isArray((value as { rich_text?: unknown }).rich_text);
}

function normalizePageId(input: string): string {
  const trimmed = input.trim();
  const matched = trimmed.match(/[0-9a-f]{32}|[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i);

  if (!matched) {
    throw new Error(`Could not extract a Notion page ID from: ${input}`);
  }

  const raw = matched[0].replace(/-/g, "");
  return `${raw.slice(0, 8)}-${raw.slice(8, 12)}-${raw.slice(12, 16)}-${raw.slice(16, 20)}-${raw.slice(20)}`;
}