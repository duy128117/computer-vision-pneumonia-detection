import hashlib
import json
from pathlib import Path
import shutil
import zipfile
import requests


def main():
    root = Path("data")
    root.mkdir(exist_ok=True)
    archive = root / "chest-xray-pneumonia.zip"
    url = "https://www.kaggle.com/api/v1/datasets/download/paultimothymooney/chest-xray-pneumonia"
    if not archive.exists():
        temporary = archive.with_suffix(".part")
        with requests.get(url, stream=True, timeout=(30, 120)) as response:
            response.raise_for_status()
            with temporary.open("wb") as stream:
                for chunk in response.iter_content(1024 * 1024):
                    stream.write(chunk)
        if not zipfile.is_zipfile(temporary):
            raise RuntimeError("Download is not ZIP; download manually from Kaggle.")
        temporary.replace(archive)
    count = 0
    with zipfile.ZipFile(archive) as source:
        for member in source.infolist():
            parts = Path(member.filename).parts
            if (
                len(parts) < 3
                or parts[-3] not in ("train", "val", "test")
                or parts[-2] not in ("NORMAL", "PNEUMONIA")
            ):
                continue
            if Path(parts[-1]).suffix.lower() not in (".jpg", ".jpeg", ".png") or parts[
                -1
            ].startswith("._"):
                continue
            destination = root / "chest_xray" / Path(*parts[-3:])
            destination.parent.mkdir(parents=True, exist_ok=True)
            if not destination.exists():
                with source.open(member) as reader, destination.open("wb") as writer:
                    shutil.copyfileobj(reader, writer)
                count += 1
    with archive.open("rb") as stream:
        digest = hashlib.file_digest(stream, "sha256").hexdigest()
    (root / "source.json").write_text(
        json.dumps(
            dict(
                url=url,
                sha256=digest,
                citation="Kermany, Zhang, Goldbaum (2018), doi:10.17632/rscbjbr9sj.2",
                license="CC BY 4.0",
                extracted=count,
            ),
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"Extracted {count} images", flush=True)


if __name__ == "__main__":
    main()
