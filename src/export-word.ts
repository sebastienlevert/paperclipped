import * as path from "path";
import * as vscode from "vscode";
import { execFile } from "child_process";

/**
 * Check if the MarkMyWord CLI is available.
 */
async function isMarkMyWordInstalled(): Promise<boolean> {
  return new Promise((resolve) => {
    execFile("markmyword", ["--version"], { timeout: 10000 }, (err) => {
      resolve(!err);
    });
  });
}

/**
 * Convert a markdown file to Word (.docx) using the MarkMyWord CLI.
 */
async function convertToWord(
  inputPath: string,
  outputPath: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    execFile(
      "markmyword",
      ["convert", "-i", inputPath, "-o", outputPath, "--force"],
      { timeout: 30000 },
      (err, _stdout, stderr) => {
        if (err) {
          reject(new Error(stderr || err.message));
        } else {
          resolve();
        }
      }
    );
  });
}

/**
 * Export the current markdown file to a Word document.
 * The .docx is saved alongside the .md file.
 */
export async function exportToWord(filePath: string): Promise<void> {
  const ext = path.extname(filePath).toLowerCase();
  if (ext !== ".md" && ext !== ".markdown") {
    vscode.window.showWarningMessage(
      "Export to Word is only available for Markdown files."
    );
    return;
  }

  // Check if MarkMyWord is installed
  const installed = await isMarkMyWordInstalled();
  if (!installed) {
    const install = await vscode.window.showWarningMessage(
      "MarkMyWord CLI is required for Markdown → Word conversion. Install it?",
      "Install",
      "Cancel"
    );
    if (install === "Install") {
      const terminal = vscode.window.createTerminal("Install MarkMyWord");
      terminal.show();
      terminal.sendText(
        "dotnet tool install -g specworks.markmyword.cli"
      );
      vscode.window.showInformationMessage(
        "Installing MarkMyWord. Run the export again after installation completes."
      );
    }
    return;
  }

  // Build output path
  const dir = path.dirname(filePath);
  const baseName = path.basename(filePath, ext);
  const outputPath = path.join(dir, `${baseName}.docx`);

  try {
    await vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: "Exporting to Word...",
        cancellable: false,
      },
      async () => {
        await convertToWord(filePath, outputPath);
      }
    );

    const action = await vscode.window.showInformationMessage(
      `Exported to ${path.basename(outputPath)}`,
      "Open in Word",
      "Reveal in Explorer"
    );

    if (action === "Open in Word") {
      const uri = vscode.Uri.file(outputPath);
      vscode.commands.executeCommand("paperclipped.openInWord", uri);
    } else if (action === "Reveal in Explorer") {
      vscode.commands.executeCommand(
        "revealFileInOS",
        vscode.Uri.file(outputPath)
      );
    }
  } catch (err: any) {
    vscode.window.showErrorMessage(
      `Export failed: ${err.message}`
    );
  }
}
