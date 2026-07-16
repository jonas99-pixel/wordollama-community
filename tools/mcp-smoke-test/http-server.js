const http = require("http");

const port = Number(process.argv[2] || 31997);
let legacySseResponse = null;

function sendJson(res, message, extraHeaders = {}) {
  res.writeHead(200, {
    "content-type": "application/json",
    ...extraHeaders,
  });
  res.end(JSON.stringify(message));
}

function responseFor(message) {
  if (message.method === "initialize") {
    return {
      jsonrpc: "2.0",
      id: message.id,
      result: {
        protocolVersion: "2025-11-25",
        capabilities: { tools: {} },
        serverInfo: { name: "wordollama-http-smoke", version: "1.0.0" },
      },
    };
  }

  if (message.method === "tools/list") {
    return {
      jsonrpc: "2.0",
      id: message.id,
      result: {
        tools: [
          {
            name: "remote_echo",
            description: "Remote echo tool.",
            inputSchema: {
              type: "object",
              properties: { text: { type: "string" } },
              required: ["text"],
            },
            annotations: { readOnlyHint: true },
          },
        ],
      },
    };
  }

  if (message.method === "tools/call") {
    return {
      jsonrpc: "2.0",
      id: message.id,
      result: {
        content: [
          {
            type: "text",
            text: `remote:${message.params?.arguments?.text ?? ""}`,
          },
        ],
      },
    };
  }

  return { jsonrpc: "2.0", id: message.id, result: {} };
}

function writeSse(res, eventName, data, id) {
  if (id) res.write(`id: ${id}\n`);
  if (eventName) res.write(`event: ${eventName}\n`);
  res.write(`data: ${typeof data === "string" ? data : JSON.stringify(data)}\n\n`);
}

function readBody(req) {
  return new Promise((resolve) => {
    let body = "";
    req.setEncoding("utf8");
    req.on("data", (chunk) => (body += chunk));
    req.on("end", () => resolve(body ? JSON.parse(body) : null));
  });
}

const server = http.createServer(async (req, res) => {
  if (req.method === "GET" && req.url === "/health") {
    res.writeHead(200, { "content-type": "text/plain" });
    res.end("ok");
    return;
  }

  if (req.method === "POST" && req.url === "/mcp") {
    const message = await readBody(req);
    if (message.method === "notifications/initialized" || message.result || message.error) {
      res.writeHead(202);
      res.end();
      return;
    }

    if (message.method === "initialize") {
      sendJson(res, responseFor(message), { "MCP-Session-Id": "smoke-session" });
      return;
    }

    if (message.method === "tools/call") {
      res.writeHead(200, { "content-type": "text/event-stream" });
      writeSse(res, "message", {
        jsonrpc: "2.0",
        id: "server-ping",
        method: "ping",
        params: {},
      }, "1");
      writeSse(res, "message", responseFor(message), "2");
      res.end();
      return;
    }

    sendJson(res, responseFor(message));
    return;
  }

  if (req.method === "GET" && req.url === "/sse") {
    legacySseResponse = res;
    res.writeHead(200, {
      "content-type": "text/event-stream",
      "cache-control": "no-cache",
      connection: "keep-alive",
    });
    writeSse(res, "endpoint", `http://127.0.0.1:${port}/message`, "endpoint-1");
    req.on("close", () => {
      if (legacySseResponse === res) legacySseResponse = null;
    });
    return;
  }

  if (req.method === "POST" && req.url === "/message") {
    const message = await readBody(req);
    res.writeHead(202);
    res.end();
    if (legacySseResponse && message && !message.method?.startsWith("notifications/")) {
      writeSse(legacySseResponse, "message", responseFor(message));
    }
    return;
  }

  res.writeHead(404);
  res.end();
});

server.listen(port, "127.0.0.1", () => {
  console.error(`http-smoke:${port}`);
});
