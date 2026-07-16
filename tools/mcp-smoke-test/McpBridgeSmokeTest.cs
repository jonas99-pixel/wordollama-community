using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using WordOllama.lib;

internal static class McpBridgeSmokeTest
{
    private static int Main(string[] args)
    {
        try
        {
            return RunAsync(args).GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }

    private static async Task<int> RunAsync(string[] args)
    {
        if (args.Length < 1 || args.Length > 2)
        {
            Console.Error.WriteLine("Usage: McpBridgeSmokeTest.exe <strict-server.js> [command]");
            return 2;
        }

        string scriptPath = Path.GetFullPath(args[0]);
        string command = args.Length == 2 ? Path.GetFullPath(args[1]) : "node";
        string commandArgs = command.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase)
            ? ""
            : Quote(scriptPath);
        string workingDirectory = Path.GetDirectoryName(scriptPath);

        using (var bridge = new McpBridge(
            "smoke",
            command,
            commandArgs,
            workingDirectory,
            new Dictionary<string, string> { ["WORDOLLAMA_MCP_SMOKE"] = "ok" }))
        {
            if (!await bridge.ConnectAsync())
            {
                Console.Error.WriteLine("FAIL: initialize handshake failed.");
                return 3;
            }

            var tools = await bridge.ListToolsAsync();
            if (tools.Count != 3 || tools[0].Name != "echo" || tools[1].Name != "second_tool" || tools[2].Name != "env_probe")
            {
                Console.Error.WriteLine($"FAIL: expected three paginated tools, found {tools.Count} tools.");
                return 4;
            }

            if (tools[1].Annotations?["destructiveHint"]?.ToObject<bool>() != true ||
                tools[2].Annotations?["readOnlyHint"]?.ToObject<bool>() != true)
            {
                Console.Error.WriteLine("FAIL: tool annotations were not preserved.");
                return 6;
            }

            string result = await bridge.CallToolAsync("echo", new JObject { ["text"] = "hello" });
            if (result != "echo:hello")
            {
                Console.Error.WriteLine($"FAIL: unexpected tool result: {result}");
                return 5;
            }

            string envResult = await bridge.CallToolAsync("env_probe", new JObject());
            if (!envResult.Contains("env:ok") || !envResult.Contains($"cwd:{workingDirectory}"))
            {
                Console.Error.WriteLine($"FAIL: cwd/env were not passed to server: {envResult}");
                return 7;
            }
        }

        Console.WriteLine("PASS: MCP lifecycle, server ping, paginated tools/list, cwd/env, annotations, and tools/call succeeded.");
        return 0;
    }

    private static string Quote(string value)
    {
        return "\"" + value.Replace("\"", "\\\"") + "\"";
    }
}
