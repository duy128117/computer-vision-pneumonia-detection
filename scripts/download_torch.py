"""Resumable ranged download from the official PyTorch host; verify index SHA256."""

from concurrent.futures import ThreadPoolExecutor, as_completed
import hashlib
from pathlib import Path
import re
import time
import requests


def main():
    name = "torch-2.6.0+cu124-cp311-cp311-win_amd64.whl"
    index = requests.get("https://download.pytorch.org/whl/cu124/torch/", timeout=60)
    index.raise_for_status()
    links = re.findall(r'href="([^"]+)"', index.text)
    link = next(link for link in links if name.replace("+", "%2B") in link or name in link)
    expected = link.split("#sha256=")[1]
    url = "https://download.pytorch.org/whl/cu124/" + name.replace("+", "%2B")
    size = 2532350702
    folder = Path(".tmp/torch_chunks")
    folder.mkdir(parents=True, exist_ok=True)
    chunk = 8 * 1024 * 1024

    def download(start):
        end = min(start + chunk, size) - 1
        path = folder / str(start)
        if path.exists() and path.stat().st_size == end - start + 1:
            return
        for attempt in range(5):
            try:
                response = requests.get(
                    url + f"?chunk={start}",
                    headers={"Range": f"bytes={start}-{end}"},
                    timeout=(30, 90),
                )
                response.raise_for_status()
                if (
                    response.status_code != 206
                    or len(response.content) != end - start + 1
                    or response.headers.get("Content-Range") != f"bytes {start}-{end}/{size}"
                ):
                    raise ValueError("Invalid range response")
                path.write_bytes(response.content)
                return
            except (requests.RequestException, ValueError):
                if attempt == 4:
                    raise
                time.sleep(2)

    starts = list(range(0, size, chunk))
    with ThreadPoolExecutor(max_workers=12) as pool:
        for done, future in enumerate(as_completed([pool.submit(download, s) for s in starts]), 1):
            future.result()
            if done % 10 == 0:
                print(f"Torch chunks {done}/{len(starts)}", flush=True)
    output = Path(".tmp") / name
    digest = hashlib.sha256()
    with output.open("wb") as stream:
        for start in starts:
            data = (folder / str(start)).read_bytes()
            digest.update(data)
            stream.write(data)
    if digest.hexdigest() != expected:
        raise ValueError("Official wheel SHA256 mismatch")
    print(f"Verified {output}", flush=True)


if __name__ == "__main__":
    main()
