from concurrent.futures import ThreadPoolExecutor
import hashlib
from pathlib import Path
import requests


def download(name):
    output = Path(".torch/hub/checkpoints") / name
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        return
    temp = output.with_suffix(".part")
    url = "https://download.pytorch.org/models/" + name
    response = requests.get(url, headers={"Range": "bytes=0-0"}, timeout=(30, 90))
    response.raise_for_status()
    if response.status_code != 206:
        temp.write_bytes(response.content)
    else:
        size = int(response.headers["Content-Range"].split("/")[-1])
        chunk = 4 * 1024 * 1024

        def fetch(start):
            end = min(start + chunk, size) - 1
            r = requests.get(
                url + f"?chunk={start}", headers={"Range": f"bytes={start}-{end}"}, timeout=(30, 90)
            )
            r.raise_for_status()
            if (
                r.status_code != 206
                or len(r.content) != end - start + 1
                or r.headers.get("Content-Range") != f"bytes {start}-{end}/{size}"
            ):
                raise ValueError("Invalid weight byte range")
            return r.content

        with ThreadPoolExecutor(max_workers=4) as pool, temp.open("wb") as stream:
            for data in pool.map(fetch, range(0, size, chunk)):
                stream.write(data)
    with temp.open("rb") as stream:
        digest = hashlib.file_digest(stream, "sha256").hexdigest()
    if not digest.startswith(name.split("-")[-1].split(".")[0]):
        raise ValueError("Weight checksum mismatch")
    temp.replace(output)
    print(f"Verified {name}", flush=True)


if __name__ == "__main__":
    with ThreadPoolExecutor(max_workers=2) as pool:
        list(pool.map(download, ["resnet50-11ad3fa6.pth", "densenet121-a639ec97.pth"]))
