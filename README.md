# ICSickle

> Embed and extract payloads from Outlook `.ics` files using custom properties instead of attachments.

ICSickle explores how Microsoft Outlook handles RFC 5545 `X-` fields and uses that behavior to store arbitrary data inside calendar events without anything appearing in the UI.

## Overview

ICSickle allows you to:

- Embed any file into an `.ics` event
- Store data in `X-` properties instead of `ATTACH`
- Avoid visible attachment artifacts in Outlook
- Chunk large payloads automatically
- Optionally compress data with `gzip` or `zlib`
- Extract and restore the original file exactly

Filenames are preserved 1:1. No renaming or guesswork.

## How it works

Payload data is base64-encoded and stored in custom calendar properties:

```text
X-PADDING;X-FILENAME=payload.exe;X-COMPRESS=none;X-SEQ=1;X-TOTAL=3:<data>
```

Outlook ignores these fields in the UI, but they can survive normal calendar transport. The extractor reassembles the chunks, normalizes formatting issues, and reconstructs the original file.

## Usage

### Embed a payload

```bash
python embed_outlook_ics2.py \
  --summary "Meeting" \
  --file payload.exe \
  --output event.ics
```

### Extract a payload

```bash
python extract_outlook_ics2.py \
  --ics event.ics
```

Run either script with `-h` to see additional options for compression, chunk size, timestamps, and custom property prefixes.

## Files

| File | Description |
| --- | --- |
| `embed_outlook_ics2.py` | Embeds a local file into an Outlook-compatible `.ics` file. |
| `extract_outlook_ics2.py` | Extracts and restores the embedded file from an `.ics` file. |

## Notes

- Uses RFC 5545-compliant line folding for transport safety
- Handles CRLF/LF normalization and folded parameters
- Normalizes compression values during extraction
- Prevents path traversal when writing output
- Designed around Outlook-compatible `.ics` behavior

## Why

Outlook is permissive with non-standard calendar properties. That makes `.ics` files a useful container for data that does not present as a traditional attachment.

## Disclaimer

This project is intended for authorized security research and testing only. Do not use it on systems without permission.
