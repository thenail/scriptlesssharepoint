// scriptless-sharepoint-mcp: the local MCP server on stdio. AI apps start it
// themselves (Claude Desktop, Claude Code, VS Code, Windows' agent registry);
// in the Store package it's `scriptless-sharepoint-mcp.exe`, an app execution alias.
//   scriptless-sharepoint-mcp              serve MCP on stdin/stdout
//   scriptless-sharepoint-mcp --manifest   print the MCP bundle manifest (package/build.ps1)
//   scriptless-sharepoint-mcp --version
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

using ScriptlessSharePoint;
using ScriptlessSharePoint.Connections;
using ScriptlessSharePoint.Mcp;

var utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

if (args.Contains("--manifest"))
{
    using var stdout = new StreamWriter(Console.OpenStandardOutput(), utf8);
    stdout.WriteLine(Manifest.Bundle(AppInfo.Version).ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
    return 0;
}
if (args.Contains("--version"))
{
    Console.WriteLine(AppInfo.Version);
    return 0;
}

// stdout carries the protocol only
var output = new StreamWriter(Console.OpenStandardOutput(), utf8) { AutoFlush = true, NewLine = "\n" };
var errors = new StreamWriter(Console.OpenStandardError(), utf8) { AutoFlush = true };
Console.SetOut(errors);
var writeLock = new object();
ScriptlessSharePoint.Diagnostics.Log.Init("mcp");
void Log(string line)
{
    lock (writeLock)
    {
        errors.WriteLine($"[{AppInfo.McpAlias}] {line}");
    }
    ScriptlessSharePoint.Diagnostics.Log.Write("mcp", line);
}

var server = new McpServer(
    new SharePointBackend(new ConnectionStore()),
    message =>
    {
        var json = message.ToJsonString();
        lock (writeLock)
        {
            output.WriteLine(json);
        }
    },
    AppInfo.Version,
    Log);

Log($"{AppInfo.Name} MCP server {AppInfo.Version} ready");
using var input = new StreamReader(Console.OpenStandardInput(), utf8);
var running = new List<Task>();
while (await input.ReadLineAsync() is { } line)
{
    // not awaited one by one: an elicitation's answer arrives while its call waits
    running.Add(Task.Run(async () =>
    {
        try
        {
            await server.ReceiveLineAsync(line);
        }
        catch (Exception e)
        {
            Log($"error: {e.Message}");
        }
    }));
    running.RemoveAll(t => t.IsCompleted);
}
await Task.WhenAll(running);
return 0;
