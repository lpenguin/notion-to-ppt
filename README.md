# notion-to-ppt

Convert a Notion page into a PowerPoint deck with Bun, TypeScript, and Marp.

## What it does

- Pulls a Notion page and its nested child blocks through the official Notion SDK.
- Converts the page content into Marp-flavored Markdown.
- Renders the Markdown into a `.pptx` presentation through Marp.

## Setup

1. Install dependencies:

```bash
bun install
```

2. Create a Notion integration and share the source page with it.

3. Set your token:

```bash
cp .env.example .env
```

Then add your Notion integration token to `.env`.

## Usage

```bash
bun run convert --page "https://www.notion.so/workspace/My-Page-12345678123412341234123456789abc" --out deck.pptx
```

You can also keep the intermediate Marp markdown:

```bash
bun run convert --page 12345678-1234-1234-1234-123456789abc --out deck.pptx --markdown deck.md
```

## Mapping rules

- `Heading 1` starts a new slide.
- `Divider` starts a new slide.
- Paragraphs, lists, quotes, toggles, callouts, code blocks, links, files, embeds, and images are converted into slide content.
- Notion column layouts are converted into Marp-specific HTML so sibling columns stay side-by-side on the same slide.
- Nested list-like Notion blocks are preserved as nested Markdown lists where possible.

## Notes

- Marp generates slide-image based PPTX output by default.
- Notion file URLs can be temporary, so exported decks should be generated close to when they are used.
- Column widths currently fall back to equal splits because the installed Notion SDK types expose column structure but not width ratios.
- Complex Notion layouts such as databases and advanced tables are not fully converted in this first version.