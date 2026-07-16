using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WordOllama.lib;

internal static class McpManagerSmokeTest
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
        if (args.Length != 1)
        {
            Console.Error.WriteLine("Usage: McpManagerSmokeTest.exe <strict-server.cmd>");
            return 2;
        }

        using (var manager = new McpToolManager())
        {
            string commandPath = Path.GetFullPath(args[0]);
            string workingDirectory = Path.GetDirectoryName(commandPath);
            await manager.AddServerAsync(
                "测试 server",
                commandPath,
                "",
                workingDirectory,
                new Dictionary<string, string> { ["WORDOLLAMA_MCP_SMOKE"] = "ok" });
            if (manager.ServerCount != 1 || manager.ToolCount != 3)
            {
                Console.Error.WriteLine($"FAIL: expected 1 server and 3 tools, found {manager.ServerCount}/{manager.ToolCount}.");
                return 3;
            }

            var wrapper = new McpToolWrapper(manager.GetDiscoveredTools()[0], manager);
            if (wrapper.Name.Length > 64 || !Regex.IsMatch(wrapper.Name, "^[A-Za-z0-9_-]+$"))
            {
                Console.Error.WriteLine($"FAIL: unsafe model-facing tool name: {wrapper.Name}");
                return 4;
            }

            string result = wrapper.Execute(
                new Newtonsoft.Json.Linq.JObject { ["text"] = "manager" },
                null,
                null);
            if (result != "echo:manager")
            {
                Console.Error.WriteLine($"FAIL: unexpected wrapper result: {result}");
                return 5;
            }

            var destructiveWrapper = new McpToolWrapper(manager.GetDiscoveredTools()[1], manager);
            if (ToolExecutionPolicy.GetEffects(destructiveWrapper) != ToolEffect.ExternalMcp)
            {
                Console.Error.WriteLine("FAIL: MCP wrapper was not classified as an external MCP effect.");
                return 6;
            }
        }

        Console.WriteLine("PASS: MCP manager discovery, safe tool naming, annotations, wrapper call, and cleanup succeeded.");
        return 0;
    }
}
