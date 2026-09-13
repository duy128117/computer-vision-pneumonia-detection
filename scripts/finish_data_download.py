"""Resume an interrupted public ZIP download using byte ranges."""

from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
import shutil
import requests


def main():
    url = "https://www.kaggle.com/api/v1/datasets/download/paultimothymooney/chest-xray-pneumonia"
    part = Path("data/chest-xray-pneumonia.part")
    destination = part.with_suffix(".zip")
    prefix = part.stat().st_size
    response = requests.get(url, headers={"Range": "bytes=0-0"}, stream=True, timeout=(30, 90))
    response.raise_for_status()
    if response.status_code != 206:
        response.close()
        raise ValueError("Server does not support Range; use download_data.py.")
    size = int(response.headers["Content-Range"].split("/")[-1])
    direct = response.url
    response.close()
    folder = Path(".tmp/data_chunks")
    folder.mkdir(parents=True, exist_ok=True)
    chunk = 8 * 1024 * 1024
    starts = list(range(prefix, size, chunk))

    def fetch(start):
        end = min(start + chunk, size) - 1
        path = folder / str(start)
        if path.exists() and path.stat().st_size == end - start + 1:
            return
        for attempt in range(5):
            try:
                r = requests.get(
                    direct, headers={"Range": f"bytes={start}-{end}"}, timeout=(30, 90)
                )
                r.raise_for_status()
                if (
                    r.status_code != 206
                    or len(r.content) != end - start + 1
                    or r.headers.get("Content-Range") != f"bytes {start}-{end}/{size}"
                ):
                    raise ValueError("Invalid byte range")
                path.write_bytes(r.content)
                return
            except (requests.RequestException, ValueError):
                if attempt == 4:
                    raise

    print(f"Resume {prefix}/{size} bytes, {len(starts)} chunks", flush=True)
    with ThreadPoolExecutor(max_workers=8) as pool:
        for index, future in enumerate(
            as_completed([pool.submit(fetch, start) for start in starts]), 1
        ):
            future.result()
            if index % 10 == 0:
                print(f"Dataset chunks {index}/{len(starts)}", flush=True)
    with destination.open("wb") as output, part.open("rb") as source:
        shutil.copyfileobj(source, output)
        for start in starts:
            output.write((folder / str(start)).read_bytes())
    print("ZIP assembled; extraction will validate CRC.", flush=True)


if __name__ == "__main__":
    main()
