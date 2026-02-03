from __future__ import annotations

import os


def to_unc(path: str) -> str:
    if not path:
        return path
    if os.name != "nt":
        return path
    if len(path) < 2 or path[1] != ":":
        return path

    import ctypes
    from ctypes import wintypes

    class UNIVERSAL_NAME_INFOW(ctypes.Structure):
        _fields_ = [("lpUniversalName", wintypes.LPWSTR)]

    WNetGetUniversalNameW = ctypes.windll.mpr.WNetGetUniversalNameW
    WNetGetUniversalNameW.argtypes = (
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.LPVOID,
        ctypes.POINTER(wintypes.DWORD),
    )
    WNetGetUniversalNameW.restype = wintypes.DWORD

    UNIVERSAL_NAME_INFO_LEVEL = 0x00000001
    ERROR_MORE_DATA = 234
    NO_ERROR = 0

    normalized = path if len(path) > 2 else f"{path}\\"
    buf_size = wintypes.DWORD(2048)
    buf = ctypes.create_string_buffer(buf_size.value)
    res = WNetGetUniversalNameW(
        normalized, UNIVERSAL_NAME_INFO_LEVEL, buf, ctypes.byref(buf_size)
    )
    if res == ERROR_MORE_DATA:
        buf = ctypes.create_string_buffer(buf_size.value)
        res = WNetGetUniversalNameW(
            normalized, UNIVERSAL_NAME_INFO_LEVEL, buf, ctypes.byref(buf_size)
        )
    if res != NO_ERROR:
        message = ctypes.FormatError(res).strip() or "sem mensagem"
        raise OSError(
            f"Nao consegui resolver o drive mapeado para UNC (erro {res}: {message})."
        )

    info = ctypes.cast(buf, ctypes.POINTER(UNIVERSAL_NAME_INFOW)).contents
    return info.lpUniversalName


def try_to_unc(path: str) -> tuple[str, OSError | None]:
    try:
        return to_unc(path), None
    except OSError as exc:
        return path, exc
