/**
 * Type into the running dev build and report what the editor logged.
 *
 * The editor's slow paths are only reachable through a real document with a
 * real caret, and measuring them by asking a person to type is a round trip
 * per question. This drives the renderer over the DevTools protocol instead —
 * dev builds only, where the port is open.
 *
 *   bun scripts/drive-editor.ts <characters to type>
 */

const PORT = 9333;

interface Target {
  type: string;
  title: string;
  url: string;
  webSocketDebuggerUrl?: string;
}

async function rendererTarget(): Promise<Target> {
  const response = await fetch(`http://127.0.0.1:${PORT}/json/list`);
  const targets = (await response.json()) as Target[];
  const page = targets.find(
    (target) => target.type === 'page' && !target.url.startsWith('devtools://')
  );
  if (!page?.webSocketDebuggerUrl) {
    throw new Error(`no renderer page among ${targets.length} target(s)`);
  }
  return page;
}

class Session {
  private nextId = 1;
  private pending = new Map<number, (result: unknown) => void>();
  readonly events: Array<{ method: string; params: Record<string, unknown> }> = [];

  private constructor(private socket: WebSocket) {}

  static async open(url: string): Promise<Session> {
    const socket = new WebSocket(url);
    const session = new Session(socket);
    socket.onmessage = (event) => {
      const message = JSON.parse(String(event.data)) as {
        id?: number;
        method?: string;
        params?: Record<string, unknown>;
        result?: unknown;
        error?: { message: string };
      };
      if (message.id !== undefined) {
        const resolve = session.pending.get(message.id);
        session.pending.delete(message.id);
        resolve?.(message.error ? { error: message.error } : message.result);
      } else if (message.method) {
        session.events.push({ method: message.method, params: message.params ?? {} });
      }
    };
    await new Promise<void>((resolve, reject) => {
      socket.onopen = () => resolve();
      socket.onerror = () => reject(new Error('could not connect to the renderer'));
    });
    return session;
  }

  send(method: string, params: Record<string, unknown> = {}): Promise<unknown> {
    const id = this.nextId++;
    return new Promise((resolve) => {
      this.pending.set(id, resolve);
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }

  close(): void {
    this.socket.close();
  }
}

const text = process.argv[2] ?? 'abcde';
const page = await rendererTarget();
const session = await Session.open(page.webSocketDebuggerUrl!);
await session.send('Runtime.enable');
await session.send('Console.enable');

// A caret has to be somewhere before there is anything to type into, and the
// editor places it from a click on a page rather than from focus alone.
const target = (await session.send('Runtime.evaluate', {
  expression: `(() => {
    const pages = document.querySelectorAll('canvas');
    const page = pages[Math.min(${Number(process.argv[3] ?? 0)}, pages.length - 1)];
    if (!page) return null;
    page.scrollIntoView({ block: 'center' });
    const r = page.getBoundingClientRect();
    return { x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2), pages: pages.length };
  })()`,
  returnByValue: true,
})) as { result?: { value?: { x: number; y: number; pages: number } | null } };
const point = target.result?.value;
if (!point) throw new Error('no page to click');
console.log(`clicking page ${process.argv[3] ?? 0} of ${point.pages} at ${point.x},${point.y}`);

for (const type of ['mousePressed', 'mouseReleased'] as const) {
  await session.send('Input.dispatchMouseEvent', {
    type,
    x: point.x,
    y: point.y,
    button: 'left',
    clickCount: 1,
  });
}
await new Promise((resolve) => setTimeout(resolve, 800));

for (const character of text) {
  await session.send('Input.dispatchKeyEvent', {
    type: 'keyDown',
    text: character,
    unmodifiedText: character,
    key: character,
  });
  await session.send('Input.dispatchKeyEvent', { type: 'keyUp', key: character });
  await new Promise((resolve) => setTimeout(resolve, Number(process.argv[4] ?? 1500)));
}

await new Promise((resolve) => setTimeout(resolve, 1500));
session.close();
console.log(`typed ${text.length} character(s)`);
