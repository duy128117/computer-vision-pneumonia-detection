"""Exercise real trained checkpoint inference on one validation image per class."""
import argparse
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import numpy as np
import pandas as pd
import torch
from src.inference import predict, validate_image
from src.models import load_checkpoint
from src.utils import save_json


def main(checkpoint):
    torch.set_num_threads(2)
    model, metadata = load_checkpoint(checkpoint, torch.device('cpu'))
    frame = pd.read_csv('results/manifest.csv')
    output = Path('.tmp/inference_smoke') / metadata['model']
    output.mkdir(parents=True, exist_ok=True)
    rows = []
    for label in (0, 1):
        row = frame[(frame.split == 'val') & (frame.label == label)].iloc[0]
        image = validate_image(row.path)
        result, images = predict(model, metadata, image)
        assert np.isfinite(result['pneumonia'])
        assert abs(result['normal'] + result['pneumonia'] - 1) < 1e-8
        for suffix, picture in zip(('original', 'heatmap', 'overlay'), images):
            assert picture.size == (224, 224)
            picture.save(output / f'{label}_{suffix}.png')
        rows.append(dict(path=row.path, true_label=label, **result))
    save_json(output / 'results.json', rows)
    print(f'{metadata["model"]}: real trained inference + Grad-CAM for two validation images PASS')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--checkpoint', required=True)
    main(parser.parse_args().checkpoint)
