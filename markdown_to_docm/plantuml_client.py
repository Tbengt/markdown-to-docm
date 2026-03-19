from __future__ import annotations

import zlib

import requests

_B64 = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-_"


def _encode64(data: bytes) -> str:
    result = []
    for i in range(0, len(data), 3):
        chunk = data[i : i + 3]
        b = chunk + bytes(3 - len(chunk))
        result.append(_B64[(b[0] & 0xFC) >> 2])
        result.append(_B64[((b[0] & 0x03) << 4) | ((b[1] & 0xF0) >> 4)])
        result.append(_B64[((b[1] & 0x0F) << 2) | ((b[2] & 0xFC) >> 6)])
        result.append(_B64[b[2] & 0x3F])
    return "".join(result)


def encode_plantuml(text: str) -> str:
    """Encode PlantUML source text for use in a PlantUML server URL."""
    data = text.encode("utf-8")
    compress = zlib.compressobj(9, zlib.DEFLATED, -15)
    compressed = compress.compress(data) + compress.flush()
    return _encode64(compressed)


def fetch_diagram(uml_text: str, server_url: str) -> bytes:
    """Fetch a PNG rendering of a PlantUML diagram from the given server."""
    encoded = encode_plantuml(uml_text)
    url = f"{server_url}/png/{encoded}"
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    return response.content
