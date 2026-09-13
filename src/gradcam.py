import numpy as np
import torch
from PIL import Image
from matplotlib import colormaps


def gradcam(model, tensor, target=None):
    """Binary target score: +logit for pneumonia, -logit for normal."""
    layer = model.layer4 if hasattr(model, "layer4") else model.features.norm5
    captured = {}

    def capture(module, inputs, output):
        # clone protects DenseNet's pre-ReLU feature values from its in-place ReLU.
        activation = output.clone()
        captured["activation"] = activation
        return activation

    handle = layer.register_forward_hook(capture)
    try:
        model.eval()
        with torch.enable_grad():
            logits = model(tensor.detach().clone().requires_grad_(True)).flatten()
            probability = float(logits[0].sigmoid().detach())
            target = int(probability >= 0.5) if target is None else int(target)
            score = logits[0] if target else -logits[0]
            activation = captured["activation"]
            gradients = torch.autograd.grad(score, activation)[0]
            heat = (gradients.mean((2, 3), keepdim=True) * activation).sum(1).relu()[0]
            heat = heat.detach().cpu().numpy()
            heat /= max(float(heat.max()), 1e-12)
            return probability, heat, target
    finally:
        handle.remove()


def render(image, heat):
    original = image.convert("L").convert("RGB").resize((224, 224))
    heat = np.asarray(Image.fromarray(heat).resize((224, 224), Image.Resampling.BILINEAR))
    color = Image.fromarray((colormaps["jet"](heat)[:, :, :3] * 255).astype(np.uint8))
    return original, color, Image.blend(original, color, 0.4)
