"""Passwords live in the operating system's credential store, never in a file oqa or the assistant can read back.

Windows Credential Manager through ctypes (no dependency); the `keyring` package elsewhere when installed. Callers take a SecretStore,
so tests pass a MemoryStore and never touch the real store."""
import sys


class SecretStoreUnavailable(Exception):
    pass


class SecretStore:
    def get(self, ref):
        raise NotImplementedError

    def set(self, ref, value):
        raise NotImplementedError

    def delete(self, ref):
        raise NotImplementedError

    def has(self, ref):
        return self.get(ref) is not None


class MemoryStore(SecretStore):
    def __init__(self):
        self.items = {}

    def get(self, ref):
        return self.items.get(ref)

    def set(self, ref, value):
        self.items[ref] = value

    def delete(self, ref):
        self.items.pop(ref, None)


class WindowsCredentialStore(SecretStore):
    CRED_TYPE_GENERIC, CRED_PERSIST_LOCAL_MACHINE = 1, 2

    def __init__(self):
        import ctypes
        from ctypes import wintypes
        self.ct = ctypes

        class FILETIME(ctypes.Structure):
            _fields_ = [("lo", wintypes.DWORD), ("hi", wintypes.DWORD)]

        class CREDENTIAL(ctypes.Structure):
            _fields_ = [("Flags", wintypes.DWORD), ("Type", wintypes.DWORD), ("TargetName", wintypes.LPWSTR), ("Comment", wintypes.LPWSTR),
                        ("LastWritten", FILETIME), ("CredentialBlobSize", wintypes.DWORD), ("CredentialBlob", ctypes.POINTER(ctypes.c_ubyte)),
                        ("Persist", wintypes.DWORD), ("AttributeCount", wintypes.DWORD), ("Attributes", ctypes.c_void_p),
                        ("TargetAlias", wintypes.LPWSTR), ("UserName", wintypes.LPWSTR)]

        self.CREDENTIAL = CREDENTIAL
        api = ctypes.WinDLL("advapi32", use_last_error=True)
        api.CredWriteW.argtypes, api.CredWriteW.restype = [ctypes.POINTER(CREDENTIAL), wintypes.DWORD], wintypes.BOOL
        api.CredReadW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.POINTER(ctypes.POINTER(CREDENTIAL))]
        api.CredReadW.restype = wintypes.BOOL
        api.CredDeleteW.argtypes, api.CredDeleteW.restype = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD], wintypes.BOOL
        api.CredFree.argtypes, api.CredFree.restype = [ctypes.c_void_p], None
        self.api = api

    def set(self, ref, value):
        blob = value.encode("utf-16-le")
        buffer = (self.ct.c_ubyte * len(blob)).from_buffer_copy(blob)
        cred = self.CREDENTIAL(Type=self.CRED_TYPE_GENERIC, TargetName=ref, CredentialBlobSize=len(blob),
                               CredentialBlob=self.ct.cast(buffer, self.ct.POINTER(self.ct.c_ubyte)),
                               Persist=self.CRED_PERSIST_LOCAL_MACHINE, UserName="oqa")
        if not self.api.CredWriteW(self.ct.byref(cred), 0):
            raise SecretStoreUnavailable(f"Windows Credential Manager refused the write (error {self.ct.get_last_error()}).")

    def get(self, ref):
        out = self.ct.POINTER(self.CREDENTIAL)()
        if not self.api.CredReadW(ref, self.CRED_TYPE_GENERIC, 0, self.ct.byref(out)):
            return None
        try:
            size = out.contents.CredentialBlobSize
            return self.ct.string_at(out.contents.CredentialBlob, size).decode("utf-16-le")
        finally:
            self.api.CredFree(out)

    def delete(self, ref):
        self.api.CredDeleteW(ref, self.CRED_TYPE_GENERIC, 0)


class KeyringStore(SecretStore):
    SERVICE = "odoo-test"

    def __init__(self):
        import keyring
        self.kr = keyring

    def get(self, ref):
        return self.kr.get_password(self.SERVICE, ref)

    def set(self, ref, value):
        self.kr.set_password(self.SERVICE, ref, value)

    def delete(self, ref):
        try:
            self.kr.delete_password(self.SERVICE, ref)
        except Exception:
            pass


def default_store() -> SecretStore:
    if sys.platform == "win32":
        return WindowsCredentialStore()
    try:
        return KeyringStore()
    except ImportError:
        raise SecretStoreUnavailable("No credential store found. Install the 'keyring' package, or sign in with a saved session instead of a password.")
