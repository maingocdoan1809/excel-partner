import http from "node:http";

/** A tiny stand-in for an Odoo server: form sign-in, a session probe, and a page that tries to phone home. */
export interface FakeOdoo {
  origin: string;
  /** Requests that reached the "other" server, which is not an allowed origin. */
  leaks: string[];
  otherOrigin: string;
  /** Forget every session so the next probe says "not signed in". */
  expireSessions(): void;
  logins: number;
  close(): Promise<void>;
}

const PAGE = (user: string, other: string) => `<!doctype html><html><body>
<div class="o_web_client"><div class="o_action_manager"></div>
<div class="modal show" role="dialog" style="position:fixed;bottom:0;right:0;background:#fff;padding:8px"><h4 class="modal-title">Confirmation</h4><p>Are you sure?</p><button class="btn">Ok</button><button class="btn">Cancel</button></div><div id="hello">Hello ${user}</div><div id="egress">pending</div></div>
<script>
const targets = [${JSON.stringify(other + "/by-fetch")}, "http://169.254.169.254/latest/meta-data"];
Promise.all(targets.map(t => fetch(t, { mode: "no-cors" }).then(() => "leaked", () => "blocked")))
  .then(r => { document.getElementById("egress").textContent = r.includes("leaked") ? "leaked" : "blocked"; });
fetch("/longpolling/poll", { method: "POST", body: "{}" }).catch(() => {});
const img = new Image(); img.src = ${JSON.stringify(other + "/by-image")}; document.body.appendChild(img);
</script></body></html>`;

export async function startFakeOdoo(users: Record<string, string> = { admin: "s3cret" }): Promise<FakeOdoo> {
  const sessions = new Map<string, string>();
  const leaks: string[] = [];
  let logins = 0;
  const other = http.createServer((req, res) => { leaks.push(req.url ?? ""); res.setHeader("Access-Control-Allow-Origin", "*"); res.end("x"); });
  await new Promise<void>(resolve => other.listen(0, "127.0.0.1", resolve));
  const otherOrigin = `http://127.0.0.1:${(other.address() as { port: number }).port}`;

  const server = http.createServer(async (req, res) => {
    const url = new URL(req.url ?? "/", "http://x");
    const sid = /(?:^|; )session_id=([^;]+)/.exec(req.headers.cookie ?? "")?.[1];
    const user = sid ? sessions.get(sid) : undefined;
    const body = await new Promise<string>(resolve => { let data = ""; req.on("data", chunk => { data += chunk; }); req.on("end", () => resolve(data)); });
    if (url.pathname === "/web/session/get_session_info") {
      res.setHeader("Content-Type", "application/json");
      res.end(JSON.stringify({ jsonrpc: "2.0", result: { uid: user ? 2 : false } }));
    } else if (url.pathname === "/longpolling/poll") {
      // never answered: the connection stays open like Odoo's bus
      req.on("close", () => undefined);
    } else if (url.pathname === "/web/login" && req.method === "POST") {
      const form = new URLSearchParams(body);
      const login = form.get("login") ?? "";
      if (users[login] !== undefined && users[login] === form.get("password")) {
        logins++;
        const id = `sess-${Math.random().toString(36).slice(2)}`;
        sessions.set(id, login);
        res.writeHead(303, { "Set-Cookie": `session_id=${id}; Path=/; HttpOnly`, Location: "/web" }).end();
      } else res.writeHead(200, { "Content-Type": "text/html" }).end("<p class='alert'>Wrong login/password</p>");
    } else if (url.pathname === "/web/login") {
      res.writeHead(200, { "Content-Type": "text/html" }).end(`<form method="post"><input name="login"><input name="password" type="password"><button type="submit">Log in</button></form>`);
    } else if (url.pathname === "/web") {
      if (!user) { res.writeHead(303, { Location: "/web/login" }).end(); return; }
      res.writeHead(200, { "Content-Type": "text/html" }).end(PAGE(user, otherOrigin));
    } else res.writeHead(404).end("not found");
  });
  await new Promise<void>(resolve => server.listen(0, "127.0.0.1", resolve));
  return {
    origin: `http://127.0.0.1:${(server.address() as { port: number }).port}`, otherOrigin, leaks,
    expireSessions: () => sessions.clear(),
    get logins() { return logins; },
    close: () => new Promise<void>(resolve => server.close(() => other.close(() => resolve()))),
  };
}
