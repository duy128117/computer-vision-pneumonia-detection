"""Extract the first complete 5856-image copy from the public ZIP using ranges.

This recovery utility reuses the prefix downloaded by download_data.py and verifies
each extracted member's ZIP CRC32 and uncompressed size, without downloading duplicates.
"""

from concurrent.futures import ThreadPoolExecutor, as_completed
import hashlib
import io
import json
from pathlib import Path
import struct
import zipfile
import zlib
import requests


def main():
    url = "https://www.kaggle.com/api/v1/datasets/download/paultimothymooney/chest-xray-pneumonia"
    prefix_path = Path("data/chest-xray-pneumonia.part")
    prefix = prefix_path.stat().st_size
    full_size, central_offset = 2463365435, 2461101998
    central = Path(".tmp/zip_central").read_bytes()
    index = zipfile.ZipFile(io.BytesIO(central))
    members = [
        m
        for m in index.infolist()
        if m.filename.startswith("chest_xray/chest_xray/") and m.filename.lower().endswith(".jpeg")
    ]
    end = max(
        m.header_offset
        + central_offset
        + 30
        + len(m.filename.encode())
        + len(m.extra)
        + m.compress_size
        for m in members
    )
    folder = Path(".tmp/data_chunks")
    folder.mkdir(parents=True, exist_ok=True)
    chunk = 8 * 1024 * 1024
    starts = list(range(prefix, end, chunk))

    def fetch(start):
        stop = min(start + chunk, full_size) - 1
        path = folder / str(start)
        if path.exists() and path.stat().st_size == stop - start + 1:
            return
        for attempt in range(5):
            try:
                r = requests.get(
                    url + f"?chunk={start}",
                    headers={"Range": f"bytes={start}-{stop}"},
                    timeout=(30, 90),
                )
                r.raise_for_status()
                if (
                    r.status_code != 206
                    or r.headers.get("Content-Range") != f"bytes {start}-{stop}/{full_size}"
                ):
                    raise ValueError("Invalid byte range")
                if len(r.content) != stop - start + 1:
                    raise ValueError("Truncated byte range")
                path.write_bytes(r.content)
                return
            except (requests.RequestException, ValueError):
                if attempt == 4:
                    raise

    print(f"Unique dataset: need {len(starts)} chunks after existing prefix", flush=True)
    with ThreadPoolExecutor(max_workers=24) as pool:
        for i, future in enumerate(
            as_completed([pool.submit(fetch, start) for start in starts]), 1
        ):
            future.result()
            if i % 5 == 0:
                print(f"Unique chunks {i}/{len(starts)}", flush=True)

    def read_at(offset, length):
        parts = []
        while length:
            if offset < prefix:
                count = min(length, prefix - offset)
                with prefix_path.open("rb") as stream:
                    stream.seek(offset)
                    data = stream.read(count)
            else:
                start = prefix + ((offset - prefix) // chunk) * chunk
                count = min(length, chunk - (offset - start))
                with (folder / str(start)).open("rb") as stream:
                    stream.seek(offset - start)
                    data = stream.read(count)
            if not data:
                raise ValueError("Missing archive bytes")
            parts.append(data)
            length -= len(data)
            offset += len(data)
        return b"".join(parts)

    manifest = []
    for member in members:
        offset = member.header_offset + central_offset
        header = read_at(offset, 30)
        if header[:4] != b"PK\x03\x04":
            raise ValueError("Invalid local ZIP header")
        name_length, extra_length = struct.unpack("<HH", header[26:30])
        compressed = read_at(offset + 30 + name_length + extra_length, member.compress_size)
        data = zlib.decompress(compressed, -15) if member.compress_type == 8 else compressed
        if len(data) != member.file_size or zlib.crc32(data) != member.CRC:
            raise ValueError(f"CRC or size mismatch: {member.filename}")
        target = Path("data/chest_xray") / Path(*Path(member.filename).parts[-3:])
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
        manifest.append(dict(path=str(target), sha256=hashlib.sha256(data).hexdigest()))
    Path("data/source.json").write_text(
        json.dumps(
            dict(
                url=url,
                citation="Kermany, Zhang, Goldbaum (2018), doi:10.17632/rscbjbr9sj.2",
                license="CC BY 4.0",
                method="Unique ZIP members, verified CRC32 and size; SHA256 per extracted file",
                extracted=len(manifest),
                files=manifest,
            ),
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"Extracted and verified {len(manifest)} real X-rays", flush=True)


if __name__ == "__main__":
    main()
