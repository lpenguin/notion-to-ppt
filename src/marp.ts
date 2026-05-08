import { access } from "node:fs/promises";

import { marpCli } from "@marp-team/marp-cli";

process.env.PLAYWRIGHT_BROWSERS_PATH ??= "0";

type MarpOptions = {
  markdownPath: string;
  outputPath: string;
  title: string;
};

export async function renderPptxWithMarp(options: MarpOptions): Promise<void> {
  const browserPath = await resolveChromiumPath();

  await assertBrowserPath(browserPath);

  const exitCode = await marpCli([
    options.markdownPath,
    "--pptx",
    "--browser-path",
    browserPath,
    "-o",
    options.outputPath,
    "--title",
    options.title,
  ]);

  if (exitCode !== 0) {
    throw new Error(`Marp failed with exit code ${exitCode}.`);
  }
}

async function resolveChromiumPath(): Promise<string> {
  const { chromium } = await import("playwright");
  return chromium.executablePath();
}

async function assertBrowserPath(browserPath: string): Promise<void> {
  try {
    await access(browserPath);
  } catch {
    throw new Error(
      `Playwright Chromium was not found at ${browserPath}. Run \"bun run install:browsers\" before exporting PPTX.`,
    );
  }
}