const readline = require("readline");

let initialized = false;
let pingAcknowledged = false;

function send(message) {
  process.stdout.write(`${JSON.stringify(message)}\n`);
}

const rl = readline.createInterface({
  input: process.stdin,
  crlfDelay: Infinity,
});

rl.on("line", (line) => {
  let message;
  try {
    message = JSON.parse(line);
  } catch {
    return;
  }

  if (message.method === "initialize" && message.id !== undefined) {
    send({
      jsonrpc: "2.0",
      id: message.id,
      result: {
        protocolVersion: "2024-11-05",
        capabilities: { tools: {} },
        serverInfo: { name: "wordollama-mcp-smoke", version: "1.0.0" },
      },
    });
    return;
  }

  if (message.method === "notifications/initialized") {
    initialized = true;
    send({
      jsonrpc: "2.0",
      id: "server-ping",
      method: "ping",
      params: {},
    });
    return;
  }

  if (message.id === "server-ping" && message.result) {
    pingAcknowledged = true;
    return;
  }

  if (message.method === "tools/list" && message.id !== undefined) {
    if (!initialized) {
      send({
        jsonrpc: "2.0",
        id: message.id,
        error: { code: -32002, message: "Client lifecycle handshake is incomplete" },
      });
      return;
    }

    if (message.params?.cursor === "page-2") {
      send({
        jsonrpc: "2.0",
        id: message.id,
        result: {
          tools: [
            {
              name: "second_tool",
              description: "Tool returned on the second page.",
              inputSchema: { type: "object", properties: {} },
              annotations: { readOnlyHint: false, destructiveHint: true },
            },
            {
              name: "env_probe",
              description: "Returns process cwd and a configured environment variable.",
              inputSchema: { type: "object", properties: {} },
              annotations: { readOnlyHint: true },
            },
          ],
        },
      });
      return;
    }

    send({
      jsonrpc: "2.0",
      id: message.id,
      result: {
        tools: [
          {
            name: "echo",
            description: "Returns the provided text.",
            inputSchema: {
              type: "object",
              properties: { text: { type: "string" } },
              required: ["text"],
            },
          },
        ],
        nextCursor: "page-2",
      },
    });
    return;
  }

  if (message.method === "tools/call" && message.id !== undefined) {
    if (!pingAcknowledged) {
      send({
        jsonrpc: "2.0",
        id: message.id,
        error: { code: -32003, message: "Client did not answer the server ping" },
      });
      return;
    }

    const toolName = message.params?.name;
    const text =
      toolName === "env_probe"
        ? `cwd:${process.cwd()};env:${process.env.WORDOLLAMA_MCP_SMOKE ?? ""}`
        : `echo:${message.params?.arguments?.text ?? ""}`;

    send({
      jsonrpc: "2.0",
      id: message.id,
      result: {
        content: [
          {
            type: "text",
            text,
          },
        ],
      },
    });
  }
});
