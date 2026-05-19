#!/usr/bin/env python3
"""
embed_outlook_ics.py

Embed a LOCAL FILE invisibly inside an Outlook-focused .ics using X-PADDING
(NOT ATTACH). This avoids Outlook showing an "attachment" inside the appointment UI.

1:1 requirement:
- Embedded filename is EXACTLY the basename of --file (including extension).
- No hardcoded fallback names.

Compatibility hardening:
- Keep the parameter area short to avoid RFC 5545 folding splitting tokens like
  X-COMPRESS=none into "n" + " one" when the ICS is transported via email.
- Proper CRLF output + 75-octet line folding.
"""

import argparse
import base64
import datetime as dt
import gzip
import os
import uuid
import zlib
from typing import List

# -----------------------------
# RFC 5545 helpers (escape, fold)
# -----------------------------

def ical_escape_text(value: str) -> str:
    return (
        value.replace("\\", "\\\\")
             .replace("\r\n", "\n")
             .replace("\r", "\n")
             .replace("\n", "\\n")
             .replace(";", "\\;")
             .replace(",", "\\,")
    )

def fold_ical_line(line: str, limit_octets: int = 75) -> str:
    """
    RFC 5545: lines must be <= 75 octets; fold with CRLF + single space.
    """
    b = line.encode("utf-8")
    if len(b) <= limit_octets:
        return line

    chunks = []
    start = 0
    while start < len(b):
        end = min(start + limit_octets, len(b))
        # Avoid splitting UTF-8 continuation bytes
        while end < len(b) and (b[end] & 0b11000000) == 0b10000000:
            end -= 1
        if end <= start:
            end = min(start + limit_octets, len(b))

        chunks.append(b[start:end].decode("utf-8", errors="strict"))
        start = end

    return chunks[0] + "".join("\r\n " + c for c in chunks[1:])

def dt_utc_z(d: dt.datetime) -> str:
    if d.tzinfo is None:
        d = d.replace(tzinfo=dt.timezone.utc)
    d = d.astimezone(dt.timezone.utc)
    return d.strftime("%Y%m%dT%H%M%SZ")

def parse_utc(s: str) -> dt.datetime:
    s = s.strip().replace(" ", "T")
    if not s.endswith("Z"):
        raise ValueError("Use UTC times ending with 'Z', e.g. 2026-02-02T15:00Z")
    core = s[:-1]
    d = dt.datetime.fromisoformat(core)
    return d.replace(tzinfo=dt.timezone.utc)

def strict_basename(path: str) -> str:
    base = os.path.basename(path)
    if not base or base in (".", ".."):
        raise ValueError("Could not determine a valid basename for the input file.")
    return base

# -----------------------------
# Compression (optional)
# -----------------------------

def compress_bytes(data: bytes, method: str) -> bytes:
    method = method.lower()
    if method == "none":
        return data
    if method == "gzip":
        return gzip.compress(data, compresslevel=9)
    if method == "zlib":
        return zlib.compress(data, level=9)
    raise ValueError(f"Unknown compression method: {method}")

# -----------------------------
# Hidden X- properties (NOT ATTACH)
# -----------------------------

def make_xpadding_chunks(
    payload_bytes: bytes,
    filename: str,
    compress: str,
    chunk_bytes: int,
    prop_prefix: str = "X-PADDING"
) -> List[str]:
    """
    Emits one or more lines like:
      X-PADDING;X-FILENAME=DEADBEEF.txt;X-COMPRESS=none;X-SEQ=1;X-TOTAL=1:<b64>

    Key hardening:
    - Short param list (no VALUE=..., ENCODING=...) so folding is far less likely
      to split X-COMPRESS=none into 'n' + ' one' during transport via email.
    """
    packed = compress_bytes(payload_bytes, compress)
    b64 = base64.b64encode(packed).decode("ascii")

    if chunk_bytes <= 0:
        parts = [b64]
    else:
        # base64: 4 chars ~ 3 bytes
        chunk_chars = ((chunk_bytes + 2) // 3) * 4
        parts = [b64[i:i + chunk_chars] for i in range(0, len(b64), chunk_chars)]

    total = len(parts)

    props: List[str] = []
    for i, part in enumerate(parts, start=1):
        props.append(
            f"{prop_prefix};"
            f"X-FILENAME={filename};"
            f"X-COMPRESS={compress};"
            f"X-SEQ={i};"
            f"X-TOTAL={total}:"
            f"{part}"
        )
    return props

def build_ics(summary: str,
              description: str,
              dtstart: dt.datetime,
              dtend: dt.datetime,
              x_props: List[str]) -> str:
    uid = f"{uuid.uuid4()}@outlook-xpadding"

    lines = [
        "BEGIN:VCALENDAR",
        "PRODID:-//Microsoft Corporation//Outlook 16.0 MIMEDIR//EN",
        "VERSION:2.0",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "BEGIN:VEVENT",
        f"UID:{uid}",
        f"DTSTAMP:{dt_utc_z(dt.datetime.now(dt.timezone.utc))}",
        f"DTSTART:{dt_utc_z(dtstart)}",
        f"DTEND:{dt_utc_z(dtend)}",
        f"SUMMARY:{ical_escape_text(summary)}",
        f"DESCRIPTION:{ical_escape_text(description if description else ' ')}",
    ]

    lines.extend(x_props)
    lines += ["END:VEVENT", "END:VCALENDAR"]

    return "\r\n".join(fold_ical_line(l) for l in lines) + "\r\n"

# -----------------------------
# CLI
# -----------------------------

def main() -> None:
    ap = argparse.ArgumentParser(
        description="Embed a local file invisibly into an Outlook-focused .ics using X-PADDING (no ATTACH UI). 1:1 filename."
    )
    ap.add_argument("--summary", required=True, help="Appointment subject")
    ap.add_argument("--text", default="", help="Visible DESCRIPTION text (optional)")
    ap.add_argument("--file", required=True, help="Local file to embed (basename preserved 1:1)")
    ap.add_argument("--compress", choices=["none", "gzip", "zlib"], default="none",
                    help="Compress before base64 to reduce size (best for text; often no gain for PDF/JPG/ZIP)")
    ap.add_argument("--chunk-bytes", type=int, default=50000,
                    help="Approx raw-bytes per chunk before base64 (0 disables chunking)")
    ap.add_argument("--prop-prefix", default="X-PADDING", help="Property prefix (default: X-PADDING)")
    ap.add_argument("--start", help="UTC start like 2026-02-02T15:00Z (default: now)")
    ap.add_argument("--end", help="UTC end like 2026-02-02T16:00Z (default: start+1h)")
    ap.add_argument("--output", default="event.ics", help="Output .ics path")
    args = ap.parse_args()

    if not os.path.isfile(args.file):
        raise SystemExit(f"File not found: {args.file}")

    filename = strict_basename(args.file)  # 1:1 requirement

    with open(args.file, "rb") as f:
        payload = f.read()

    dtstart = parse_utc(args.start) if args.start else dt.datetime.now(dt.timezone.utc)
    dtend = parse_utc(args.end) if args.end else (dtstart + dt.timedelta(hours=1))

    x_props = make_xpadding_chunks(
        payload_bytes=payload,
        filename=filename,
        compress=args.compress.lower(),
        chunk_bytes=args.chunk_bytes,
        prop_prefix=args.prop_prefix,
    )

    ics = build_ics(
        summary=args.summary,
        description=args.text,
        dtstart=dtstart,
        dtend=dtend,
        x_props=x_props,
    )

    with open(args.output, "w", encoding="utf-8", newline="") as f:
        f.write(ics)

    packed_len = len(compress_bytes(payload, args.compress.lower()))
    print(f"✔ Wrote {args.output}")
    print(f"  embedded filename: {filename}")
    print(f"  raw bytes: {len(payload):,}")
    print(f"  packed bytes: {packed_len:,} (pre-base64)")
    print(f"  compress: {args.compress.lower()}")
    print(f"  chunks: {len(x_props)}")

if __name__ == "__main__":
    main()
