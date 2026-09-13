import argparse
import json
from pathlib import Path
from PIL import Image, ImageOps
from .models import load_checkpoint
from .transforms import image_transform
from .gradcam import gradcam, render
from .utils import device


def validate_image(source):
    with Image.open(source) as image:
        if image.format not in ("JPEG", "PNG"):
            raise ValueError("Chỉ hỗ trợ JPEG hoặc PNG.")
        if min(image.size) < 32 or image.width * image.height > 25_000_000:
            raise ValueError("Ảnh phải có cạnh tối thiểu 32 px và không quá 25 megapixel.")
        image.load()
        return ImageOps.exif_transpose(image).convert("L").convert("RGB")


def predict(model, metadata, image):
    tensor = image_transform()(image).unsqueeze(0).to(next(model.parameters()).device)
    probability, heat, target = gradcam(model, tensor)
    return dict(
        model=metadata["model"],
        prediction="PNEUMONIA" if target else "NORMAL",
        normal=1 - probability,
        pneumonia=probability,
        confidence=probability if target else 1 - probability,
    ), render(image, heat)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--checkpoint", required=True)
    parser.add_argument("--image", required=True)
    parser.add_argument("--output", default="results/inference")
    args = parser.parse_args()
    model, metadata = load_checkpoint(args.checkpoint, device())
    result, images = predict(model, metadata, validate_image(args.image))
    output = Path(args.output)
    output.mkdir(parents=True, exist_ok=True)
    for name, image in zip(("original", "heatmap", "overlay"), images):
        image.save(output / f"{name}.png")
    (output / "prediction.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(result)
