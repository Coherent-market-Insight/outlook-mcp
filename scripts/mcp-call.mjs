// Dev probe: spawn the server over stdio and call tools through the real MCP protocol.
//   node scripts/mcp-call.mjs <tool> '<json args>'            one call
//   node scripts/mcp-call.mjs --list                          tools/list with token estimate
// Set OUTLOOK_MCP_EXE to point at a specific build.
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const here = path.dirname(fileURLToPath(import.meta.url));
const exe = process.env.OUTLOOK_MCP_EXE ??
  path.join(here, '..', 'src', 'OutlookMcp', 'bin', 'Release', 'net10.0', 'win-x64', 'outlook-mcp.exe');

export function connect(extraArgs = []) {
  const child = spawn(exe, extraArgs, { stdio: ['pipe', 'pipe', 'inherit'] });
  let buf = '';
  const pending = new Map();
  let nextId = 1;
  child.stdout.on('data', (d) => {
    buf += d.toString('utf8');
    let i;
    while ((i = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, i).trim();
      buf = buf.slice(i + 1);
      if (!line) continue;
      const msg = JSON.parse(line);
      if (msg.id && pending.has(msg.id)) { pending.get(msg.id)(msg); pending.delete(msg.id); }
    }
  });
  const rpc = (method, params) => new Promise((resolve) => {
    const id = nextId++;
    pending.set(id, resolve);
    child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n');
  });
  return {
    child,
    async init() {
      await rpc('initialize', { protocolVersion: '2025-03-26', capabilities: {}, clientInfo: { name: 'probe', version: '1' } });
      child.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized' }) + '\n');
    },
    list: async () => (await rpc('tools/list', {})).result.tools,
    async call(name, args = {}) {
      const r = await rpc('tools/call', { name, arguments: args });
      if (r.error) return { error: true, text: r.error.message };
      return { error: !!r.result.isError, text: (r.result.content ?? []).map((c) => c.text).join('\n') };
    },
    close: () => child.kill(),
  };
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const c = connect();
  await c.init();
  if (process.argv[2] === '--list') {
    const tools = await c.list();
    const chars = JSON.stringify(tools).length;
    console.log(`${tools.length} tools, ${chars} chars (~${Math.round(chars / 4)} tokens)`);
    for (const t of tools) console.log(`${t.name.padEnd(22)} ${JSON.stringify(t.inputSchema).length} chars`);
  } else {
    const [name, json] = process.argv.slice(2);
    const r = await c.call(name, json ? JSON.parse(json) : {});
    console.log(r.error ? `ERROR: ${r.text}` : r.text);
  }
  c.close();
}
