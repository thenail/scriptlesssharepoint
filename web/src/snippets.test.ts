import { describe, expect, it } from "vitest";

import { connectSnippets, sitePath } from "./snippets";

describe("connectSnippets", () => {
  it("gives each AI app the command", () => {
    const [desktop, code, vscode] = connectSnippets("scriptless-sharepoint-mcp");
    expect(JSON.parse(desktop.code)).toEqual({ mcpServers: { "scriptless-sharepoint": { command: "scriptless-sharepoint-mcp" } } });
    expect(code.code).toBe("claude mcp add scriptless-sharepoint -- scriptless-sharepoint-mcp");
    expect(JSON.parse(vscode.code).servers["scriptless-sharepoint"]).toEqual({ type: "stdio", command: "scriptless-sharepoint-mcp" });
  });

  it("quotes a path with spaces for the shell", () => {
    const code = connectSnippets("C:\\Program Files\\SSP\\scriptless-sharepoint-mcp.exe")[1].code;
    expect(code).toBe('claude mcp add scriptless-sharepoint -- "C:\\Program Files\\SSP\\scriptless-sharepoint-mcp.exe"');
  });
});

describe("sitePath", () => {
  it("shortens site URLs", () => {
    expect(sitePath("https://contoso.sharepoint.com/sites/hr/")).toBe("/sites/hr");
    expect(sitePath("https://contoso.sharepoint.com")).toBe("/");
    expect(sitePath("not a url")).toBe("not a url");
  });
});
