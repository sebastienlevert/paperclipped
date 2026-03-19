import * as vscode from "vscode";
import {
  isWorkspaceInOneDrive,
  isInOneDrive,
  discoverOneDriveRoots,
} from "./onedrive";
import { shareFile, openInOfficeApp, openOnWeb } from "./sharing";

export function activate(context: vscode.ExtensionContext): void {
  if (process.platform !== "win32") {
    return; // Windows only
  }

  // Set initial context for menu visibility
  refreshOneDriveContext();

  context.subscriptions.push(
    vscode.workspace.onDidChangeWorkspaceFolders(() =>
      refreshOneDriveContext()
    )
  );

  // ── Commands ──────────────────────────────────────────────

  context.subscriptions.push(
    vscode.commands.registerCommand(
      "onedrive-sharing.share",
      async (uri?: vscode.Uri) => {
        const filePath = resolveFilePath(uri);
        if (!filePath) {
          return;
        }
        if (!isInOneDrive(filePath)) {
          vscode.window.showWarningMessage(
            "This file is not in a OneDrive folder."
          );
          return;
        }
        await shareFile(filePath);
      }
    ),

    vscode.commands.registerCommand(
      "onedrive-sharing.openInWord",
      async (uri?: vscode.Uri) => {
        const filePath = resolveFilePath(uri);
        if (filePath) {
          await openInOfficeApp(filePath, "word");
        }
      }
    ),

    vscode.commands.registerCommand(
      "onedrive-sharing.openInExcel",
      async (uri?: vscode.Uri) => {
        const filePath = resolveFilePath(uri);
        if (filePath) {
          await openInOfficeApp(filePath, "excel");
        }
      }
    ),

    vscode.commands.registerCommand(
      "onedrive-sharing.openInPowerPoint",
      async (uri?: vscode.Uri) => {
        const filePath = resolveFilePath(uri);
        if (filePath) {
          await openInOfficeApp(filePath, "powerpoint");
        }
      }
    ),

    vscode.commands.registerCommand(
      "onedrive-sharing.openOnWeb",
      async (uri?: vscode.Uri) => {
        const filePath = resolveFilePath(uri);
        if (!filePath) {
          return;
        }
        if (!isInOneDrive(filePath)) {
          vscode.window.showWarningMessage(
            "This file is not in a OneDrive folder."
          );
          return;
        }
        await openOnWeb(filePath);
      }
    )
  );

  // ── Startup log ──────────────────────────────────────────
  const roots = discoverOneDriveRoots();
  if (roots.length > 0) {
    console.log(
      `[OneDrive Sharing] Detected ${roots.length} root(s): ${roots.map((r) => r.localPath).join(", ")}`
    );
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function refreshOneDriveContext(): void {
  const active = isWorkspaceInOneDrive(vscode.workspace.workspaceFolders);
  vscode.commands.executeCommand(
    "setContext",
    "onedrive:isOneDriveWorkspace",
    active
  );
}

function resolveFilePath(uri?: vscode.Uri): string | undefined {
  if (uri) {
    return uri.fsPath;
  }
  const editor = vscode.window.activeTextEditor;
  if (editor) {
    return editor.document.uri.fsPath;
  }
  vscode.window.showWarningMessage("No file selected.");
  return undefined;
}

export function deactivate(): void {
  // nothing to clean up
}
