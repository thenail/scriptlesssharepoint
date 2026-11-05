/** The settings each AI app needs, for the MCP server started as `command`. */
export function connectSnippets(command: string) {
  const name = "scriptless-sharepoint";
  return [
    {
      key: "claudeDesktop",
      title: "Claude Desktop",
      how: "Settings › Developer › Edit Config, add this to claude_desktop_config.json, then restart Claude.",
      code: JSON.stringify({ mcpServers: { [name]: { command } } }, null, 2),
    },
    {
      key: "claudeCode",
      title: "Claude Code",
      how: "Run this in a terminal.",
      code: `claude mcp add ${name} -- ${command.includes(" ") ? `"${command}"` : command}`,
    },
    {
      key: "vscode",
      title: "VS Code",
      how: "Add this to .vscode/mcp.json, or run MCP: Add Server.",
      code: JSON.stringify({ servers: { [name]: { type: "stdio", command } } }, null, 2),
    },
  ];
}

/** "/sites/hr" for a site URL, "/" for the tenant root. */
export function sitePath(url: string): string {
  try {
    const path = new URL(url).pathname.replace(/\/$/, "");
    return path || "/";
  } catch {
    return url;
  }
}
