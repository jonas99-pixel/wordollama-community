using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using WordOllama.lib;

internal static class McpHttpSmokeTest
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
            Console.Error.WriteLine("Usage: McpHttpSmokeTest.exe <http-server.js>");
            return 2;
        }

        int port = 31997;
        string serverScript = Path.GetFullPath(args[0]);
        using (var process = Process.Start(new ProcessStartInfo
        {
            FileName = "node",
            Arguments = Quote(serverScript) + " " + port,
            UseShellExecute = false,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            CreateNoWindow = true
        }))
        {
            try
            {
                await WaitForHealthAsync(port);

                int streamable = await ExerciseAsync(
                    McpTransportKind.StreamableHttp,
                    $"http://127.0.0.1:{port}/mcp");
                if (streamable != 0) return streamable;

                int sse = await ExerciseAsync(
                    McpTransportKind.Sse,
                    $"http://127.0.0.1:{port}/sse");
                if (sse != 0) return sse;
            }
            finally
            {
                try
                {
                    if (process != null && !process.HasExited)
                        process.Kill();
                }
                catch { }
            }
        }

        Console.WriteLine("PASS: Streamable HTTP and legacy SSE MCP transports succeeded.");
        return 0;
    }

    private static async Task<int> ExerciseAsync(McpTransportKind kind, string endpoint)
    {
        using (var bridge = new McpBridge("remote", kind, endpoint))
        {
            if (!await bridge.ConnectAsync())
            {
                Console.Error.WriteLine($"FAIL: {kind} initialize failed.");
                return 3;
            }

            var tools = await bridge.ListToolsAsync();
            if (tools.Count != 1 || tools[0].Name != "remote_echo")
            {
                Console.Error.WriteLine($"FAIL: {kind} expected remote_echo, found {tools.Count} tools.");
                return 4;
            }

            string result = await bridge.CallToolAsync("remote_echo", new JObject { ["text"] = kind.ToString() });
            if (result != $"remote:{kind}")
            {
                Console.Error.WriteLine($"FAIL: {kind} unexpected result: {result}");
                return 5;
            }
        }
        return 0;
    }

    private static async Task WaitForHealthAsync(int port)
    {
        using (var client = new HttpClient())
        {
            for (int i = 0; i < 30; i++)
            {
                try
                {
                    string text = await client.GetStringAsync($"http://127.0.0.1:{port}/health");
                    if (text == "ok") return;
                }
                catch
                {
                    await Task.Delay(100);
                }
            }
        }
        throw new TimeoutException("HTTP smoke server did not start.");
    }

    private static string Quote(string value)
    {
        return "\"" + value.Replace("\"", "\\\"") + "\"";
    }
}
