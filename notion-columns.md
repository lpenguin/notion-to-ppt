# Notion Columns

## Changes Made

- Added first-pass support for Notion `column_list` and `column` blocks in `src/markdown.ts`.
- Kept the existing fetch pipeline unchanged. Column handling happens in the markdown renderer, which is the right control point for slide layout.
- Rendered multi-column content as Marp-specific HTML so sibling columns stay side by side on the same slide.
- Added a width fallback helper that splits available width evenly across sibling columns.
- Added HTML rendering helpers for content inside column layouts so nested paragraphs, headings, lists, quotes, callouts, code, images, files, embeds, equations, and child pages still render correctly inside the column containers.
- Preserved existing markdown rendering for non-column content.
- Updated `README.md` to document column support and the current equal-width limitation.

## Current Behavior

- If a Notion `column_list` contains multiple columns, it is emitted as a flex layout in the generated Marp markdown.
- If a `column_list` has only one column, it falls back to normal flow instead of forcing layout markup.
- If there are no direct `column` children, the renderer falls back to rendering child blocks normally.
- Column widths currently default to equal percentages because the installed `@notionhq/client` types expose column structure but not width ratios.

## Validation Performed

- Ran `bun run check` successfully after the renderer changes.
- Did not yet run a live Notion page conversion or visual PPTX verification.

## Further Steps

1. Run the converter against a real Notion page containing 2-column and 3-column layouts and inspect the generated markdown and PPTX output.
2. Verify how Marp handles the emitted HTML in PPT export for mixed content such as lists, images, and callouts inside columns.
3. Inspect a real Notion API response to determine whether column width metadata exists at runtime even though it is not present in the installed SDK typings.
4. If width metadata is available, replace the equal-split helper with actual per-column width ratios.
5. Decide how to handle nested column layouts if they appear in real pages, and flatten them only if Marp rendering becomes unstable.
6. Add fixture-based tests for markdown output so column rendering behavior is covered by regression checks.