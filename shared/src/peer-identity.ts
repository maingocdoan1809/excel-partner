import { createRequire } from "node:module";
import type net from "node:net";

/**
 * Who is at the other end of a loopback connection. 127.0.0.1 is shared by every user signed in to the machine (Remote
 * Desktop servers, fast user switching), so a local server that hands out session tokens must also check that the
 * connecting process runs as the same Windows user as the server. Anything that cannot be identified is foreign.
 */
export type PeerCheck = (socket: net.Socket) => boolean;

const verdicts = new WeakMap<net.Socket, boolean>();

/** Cached per socket, so the requests of a keep-alive connection cost nothing after the first. Off with OFFICE_PARTNER_PEER_CHECK=0. */
export function isSameUserPeer(socket: net.Socket, check: PeerCheck = ownerIsCurrentUser): boolean {
  if (process.env.OFFICE_PARTNER_PEER_CHECK === "0") return true;
  let verdict = verdicts.get(socket);
  if (verdict === undefined) {
    try { verdict = check(socket); } catch { verdict = false; }
    verdicts.set(socket, verdict);
  }
  return verdict;
}

/** Drop every connection from another user's process before it is read, for HTTP, HTTPS and WebSocket upgrades alike. */
export function refuseForeignPeers(server: net.Server, onRefused: (peer: { address?: string; port?: number }) => void, check?: PeerCheck): void {
  server.prependListener("connection", (socket: net.Socket) => {
    if (isSameUserPeer(socket, check)) return;
    onRefused({ address: socket.remoteAddress, port: socket.remotePort });
    socket.destroy();
  });
}

/** Windows: the owner of the client end of the connection is the current user. Other platforms are development only. */
export function ownerIsCurrentUser(socket: net.Socket): boolean {
  if (process.platform !== "win32") return true;
  const pid = owningPid(socket);
  if (!pid) return false;
  if (pid === process.pid) return true;
  const own = currentUserSid ??= userSid(null);
  const peer = userSid(pid);
  return own !== null && peer === own;
}

const AF_INET = 2, AF_INET6 = 23, TCP_TABLE_OWNER_PID_ALL = 5, ERROR_INSUFFICIENT_BUFFER = 122;
const PROCESS_QUERY_LIMITED_INFORMATION = 0x1000, TOKEN_QUERY = 0x8, TOKEN_USER = 1;
let currentUserSid: string | null | undefined;

/** The process that owns the client end: its TCP row has the client's address and port as local and ours as remote. */
function owningPid(socket: net.Socket): number | null {
  const address = socket.remoteAddress?.replace(/^::ffff:/i, "");
  const { remotePort, localPort } = socket;
  const v4 = address !== undefined && /^127\.\d{1,3}\.\d{1,3}\.\d{1,3}$/.test(address);
  if (!remotePort || !localPort || !(v4 || address === "::1")) return null; // not loopback: another machine, never trusted
  // MIB_TCPROW_OWNER_PID (24 bytes) and MIB_TCP6ROW_OWNER_PID (56 bytes): [row size, local address, local port, remote port, pid].
  const [size, addressAt, localPortAt, remotePortAt, pidAt] = v4 ? [24, 4, 8, 16, 20] : [56, 0, 20, 44, 52];
  const wanted = v4 ? Buffer.from(address.split(".").map(Number)) : Buffer.from([...new Array<number>(15).fill(0), 1]);
  const table = tcpTable(v4 ? AF_INET : AF_INET6);
  const port = (at: number) => (table[at]! << 8) | table[at + 1]!; // network byte order in the low word
  for (let i = 0, rows = table.readUInt32LE(0); i < rows; i++) {
    const row = 4 + i * size;
    if (port(row + localPortAt) === remotePort && port(row + remotePortAt) === localPort
      && table.subarray(row + addressAt, row + addressAt + wanted.length).equals(wanted)) return table.readUInt32LE(row + pidAt);
  }
  return null;
}

function tcpTable(family: number): Buffer {
  const size = [64 * 1024];
  for (let attempt = 0; attempt < 5; attempt++) {
    const buffer = Buffer.alloc(size[0]!);
    const status = win32().GetExtendedTcpTable(buffer, size, 0, family, TCP_TABLE_OWNER_PID_ALL, 0) as number;
    if (status === 0) return buffer;
    if (status !== ERROR_INSUFFICIENT_BUFFER) throw new Error(`GetExtendedTcpTable failed (${status})`);
    size[0] = size[0]! + 16 * 1024; // the table can grow between two calls
  }
  throw new Error("GetExtendedTcpTable kept growing");
}

/** The string SID of a process's user (null: this process), or null when the process cannot be opened, as other users' cannot. */
function userSid(pid: number | null): string | null {
  const api = win32();
  const process = pid === null ? api.GetCurrentProcess() : api.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, pid);
  if (!process) return null;
  try {
    const token: unknown[] = [null];
    if (!api.OpenProcessToken(process, TOKEN_QUERY, token)) return null;
    try {
      const info = Buffer.alloc(256), length = [0];
      if (!api.GetTokenInformation(token[0], TOKEN_USER, info, info.length, length)) return null;
      const text: unknown[] = [null];
      if (!api.ConvertSidToStringSidW(api.koffi.decode(info, "void *"), text)) return null; // TOKEN_USER starts with the SID pointer
      try { return api.koffi.decode(text[0], "char16_t", -1) as string; } finally { api.LocalFree(text[0]); }
    } finally { api.CloseHandle(token[0]); }
  } finally { if (pid !== null) api.CloseHandle(process); }
}

type Koffi = typeof import("koffi");
type Fn = (...args: any[]) => any;
let bindings: Record<"GetExtendedTcpTable" | "OpenProcess" | "GetCurrentProcess" | "CloseHandle" | "LocalFree" | "OpenProcessToken" | "GetTokenInformation" | "ConvertSidToStringSidW", Fn> & { koffi: Koffi } | undefined;

function win32() {
  if (bindings) return bindings;
  const koffi = createRequire(import.meta.url)("koffi") as Koffi;
  const iphlpapi = koffi.load("iphlpapi.dll"), kernel32 = koffi.load("kernel32.dll"), advapi32 = koffi.load("advapi32.dll");
  bindings = {
    koffi,
    GetExtendedTcpTable: iphlpapi.func("uint32_t __stdcall GetExtendedTcpTable(void *table, _Inout_ uint32_t *size, int order, uint32_t family, int tableClass, uint32_t reserved)"),
    OpenProcess: kernel32.func("void * __stdcall OpenProcess(uint32_t access, int inherit, uint32_t pid)"),
    GetCurrentProcess: kernel32.func("void * __stdcall GetCurrentProcess()"),
    CloseHandle: kernel32.func("int __stdcall CloseHandle(void *handle)"),
    LocalFree: kernel32.func("void * __stdcall LocalFree(void *memory)"),
    OpenProcessToken: advapi32.func("int __stdcall OpenProcessToken(void *process, uint32_t access, _Out_ void **token)"),
    GetTokenInformation: advapi32.func("int __stdcall GetTokenInformation(void *token, int infoClass, void *info, uint32_t length, _Out_ uint32_t *returned)"),
    ConvertSidToStringSidW: advapi32.func("int __stdcall ConvertSidToStringSidW(void *sid, _Out_ void **text)"),
  };
  return bindings;
}
