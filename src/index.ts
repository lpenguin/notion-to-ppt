import { basename, dirname, extname, join, resolve } from "node:path";
import { Command } from "commander";

import { renderDeckMarkdown } from "./markdown.ts";
import { renderPptxWithMarp } from "./marp.ts";
import { createNotionClient, fetchPageContent } from "./notion.ts";

type CliOptions = {
  page: string;
  output?: string;
  markdown?: string;
};

async function main(): Promise<void> {
  const options = await parseArgs(process.argv);
  const notion = createNotionClient();
  const page = await fetchPageContent(notion, options.page);
  const markdown = renderDeckMarkdown(page);

  const outputPath = resolve(options.output ?? `${slugify(page.title)}.pptx`);
  const markdownPath = resolve(
    options.markdown ?? join(dirname(outputPath), `${basename(outputPath, extname(outputPath))}.md.tmp`),
  );

  await Bun.write(markdownPath, markdown);
  await renderPptxWithMarp({ markdownPath, outputPath, title: page.title });

  if (!options.markdown) {
    await Bun.file(markdownPath).delete();
  }

  console.log(`Created ${outputPath}`);
}

async function parseArgs(argv: string[]): Promise<CliOptions> {
  const program = new Command();

  program
    .name("notion-to-ppt")
    .description("Convert a Notion page into a PPTX deck with Marp.")
    .argument("[page]", "Notion page ID or full page URL")
    .option("-p, --page <page>", "Notion page ID or full page URL")
    .option("-o, --out <path>", "Output PPTX path")
    .option("-m, --markdown <path>", "Keep the generated Marp markdown at the given path")
    .showHelpAfterError();

  await program.parseAsync(argv);

  const positionalPage = program.args[0];
  const options = program.opts<{ page?: string; out?: string; markdown?: string }>();
  const page = options.page ?? positionalPage;

  if (!page || typeof page !== "string") {
    throw new Error("A Notion page ID or URL is required. Pass it as an argument or with --page.");
  }

  return {
    page,
    output: options.out,
    markdown: options.markdown,
  };
}

function slugify(value: string): string {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "") || "notion-deck";
}

main().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : String(error);
  console.error(message);
  process.exit(1);
});