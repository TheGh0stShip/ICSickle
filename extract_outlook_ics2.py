#!/usr/bin/env python3
"""
extract_outlook_ics.py

Extract the hidden file embedded by embed_outlook_ics.py from an Outlook-focused
.ics using X-PADDING and restore the original filename + extension 1:1.

Hardening:
- Unfold both CRLF+WSP and LF+WSP (email/Outlook may normalize line endings).
- Normalize X-COMPRESS values by removing whitespace so folded "n\\n one" becomes "none".
"""

import argparse
import base64
import gzip
import os
import re
import zlib
from typing import Dict, List, Optional, Tuple

def unfold_ical(text: str) -> str:
    # RFC 5545 unfolding, tolerant of LF-only files:
    # remove (CRLF or LF) followed by space/tab
    return re.sub(r"(?:\r\n|\n)[ \t]", "", text)

def normalize_compress(method: str) -> str:
    # remove ALL whitespace to recover folded tokens, e.g. "n\n one" -> "none"
    m = re.sub(r"\s+", "", (method or "")).lower()
    if m == "none":
        return "none"
    if m == "gzip":
        return "gzip"
    if m == "zlib":
        return "zlib"
    raise ValueError(f"Unknown compression method: {method}")

def decompress_bytes(data: bytes, method: str) -> bytes:
    method = normalize_compress(method)
    if method == "none":
        return data
    if method == "gzip":
        return gzip.decompress(data)
    if method == "zlib":
        return zlib.decompress(data)
    raise ValueError(f"Unknown compression method: {method}")

def safe_write_path(out_dir: str, filename: str) -> str:
    # preserve basename+extension 1:1, but prevent path traversal
    base = os.path.basename(filename)
    if base != filename or not base or base in (".", ".."):
        raise ValueError("Embedded filename is unsafe.")
    return os.path.join(out_dir, base)

def parse_params(param_str: str) -> Dict[str, str]:
    params: Dict[str, str] = {}
    for seg in param_str.split(";"):
        if "=" not in seg:
            continue
        k, v = seg.split("=", 1)
        params[k.strip().upper()] = v.strip()
    return params

def extract_xpadding(ics_text: str, prop_prefix: str = "X-PADDING") -> Tuple[str, str, bytes]:
    u = unfold_ical(ics_text)

    # X-PADDING;...:<base64>
    pat = re.compile(rf"^{re.escape(prop_prefix)};([^:]*)\:(.*)$", re.MULTILINE)

    chunks: List[Tuple[int, str]] = []
    filename: Optional[str] = None
    compress = "none"
    total_expected: Optional[int] = None

    for m in pat.finditer(u):
        params_raw = m.group(1)
        value = m.group(2)

        pmap = parse_params(params_raw)

        f = pmap.get("X-FILENAME")
        c = pmap.get("X-COMPRESS", "none")
        seq = int(pmap.get("X-SEQ", "1"))
        tot = int(pmap.get("X-TOTAL", "1"))

        if not f:
            raise ValueError("Missing X-FILENAME; cannot do 1:1 restore.")

        if filename is None:
            filename = f
        elif filename != f:
            raise ValueError("Multiple different X-FILENAME values found; refusing to guess.")

        if total_expected is None:
            total_expected = tot

        compress = c
        chunks.append((seq, value))

    if not chunks:
        raise ValueError(f"No {prop_prefix} properties found.")

    chunks.sort(key=lambda t: t[0])

    if total_expected is not None and len(chunks) != total_expected:
        raise ValueError(f"Missing chunks: expected {total_expected}, found {len(chunks)}.")

    b64 = "".join(v for _, v in chunks)

    try:
        packed = base64.b64decode(b64, validate=True)
    except TypeError:
        packed = base64.b64decode(b64)

    raw = decompress_bytes(packed, compress)

    return filename, normalize_compress(compress), raw

def main() -> None:
    ap = argparse.ArgumentParser(
        description="Extract hidden file from an Outlook-focused .ics (X-PADDING) and restore filename+extension 1:1."
    )
    ap.add_argument("--ics", required=True, help="Input .ics file")
    ap.add_argument("--out-dir", default=".", help="Directory to write extracted file into (default: current)")
    ap.add_argument("--overwrite", action="store_true", help="Allow overwriting existing output file")
    ap.add_argument("--prop-prefix", default="X-PADDING", help="Property prefix (default: X-PADDING)")
    args = ap.parse_args()

    with open(args.ics, "r", encoding="utf-8") as f:
        text = f.read()

    filename, compress, data = extract_xpadding(text, prop_prefix=args.prop_prefix)
    out_path = safe_write_path(args.out_dir, filename)

    os.makedirs(args.out_dir, exist_ok=True)

    if os.path.exists(out_path) and not args.overwrite:
        raise SystemExit(f"Refusing to overwrite existing file: {out_path} (use --overwrite)")

    with open(out_path, "wb") as f:
        f.write(data)

    print(f"✔ Extracted to {out_path}")
    print(f"  compress: {compress}")
    print(f"  bytes: {len(data):,}")

if __name__ == "__main__":
    main()
