"""Windows user-bound DPAPI storage; no plaintext fallback or machine-wide key."""
import base64
import ctypes
import os
from ctypes import wintypes


class DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_ubyte))]


def _transform(data: bytes, decrypt: bool) -> bytes:
    if os.name != "nt":
        raise RuntimeError("에이전트 인증정보는 Windows 사용자 보안 저장소가 필요합니다.")
    source_buffer = (ctypes.c_ubyte * len(data)).from_buffer_copy(data)
    source = DataBlob(len(data), source_buffer)
    result = DataBlob()
    crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    operation = crypt32.CryptUnprotectData if decrypt else crypt32.CryptProtectData
    operation.argtypes = [ctypes.POINTER(DataBlob), ctypes.c_void_p, ctypes.c_void_p,
                          ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(DataBlob)]
    operation.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p
    # Do not set CRYPTPROTECT_LOCAL_MACHINE: another Windows user must not decrypt.
    if not operation(ctypes.byref(source), None, None, None, None, 0x1, ctypes.byref(result)):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(result.pbData, result.cbData)
    finally:
        kernel32.LocalFree(result.pbData)


def protect_token(token: str) -> str:
    if not isinstance(token, str) or not token or len(token) > 4096:
        raise ValueError("유효하지 않은 에이전트 토큰입니다.")
    return "dpapi-v1:" + base64.b64encode(_transform(token.encode("utf-8"), False)).decode("ascii")


def unprotect_token(value: str) -> str:
    if not isinstance(value, str) or not value.startswith("dpapi-v1:"):
        raise ValueError("지원하지 않는 인증정보 형식입니다.")
    return _transform(base64.b64decode(value[9:], validate=True), True).decode("utf-8")
