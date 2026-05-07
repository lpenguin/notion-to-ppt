import { marpCli } from "@marp-team/marp-cli";

type MarpOptions = {
  markdownPath: string;
  outputPath: string;
  title: string;
};

export async function renderPptxWithMarp(options: MarpOptions): Promise<void> {
  const exitCode = await marpCli([
    options.markdownPath,
    "--pptx",
    "-o",
    options.outputPath,
    "--title",
    options.title,
  ]);

  if (exitCode !== 0) {
    throw new Error(`Marp failed with exit code ${exitCode}.`);
  }
}