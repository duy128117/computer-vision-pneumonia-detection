from torchvision import transforms as T


def image_transform(train=False):
    steps = [T.Resize((224, 224)), T.Grayscale(num_output_channels=3)]
    if train:
        steps += [
            T.RandomAffine(degrees=7, translate=(0.03, 0.03), scale=(0.95, 1.05)),
            T.ColorJitter(brightness=0.08, contrast=0.08),
        ]
    return T.Compose(
        steps + [T.ToTensor(), T.Normalize([0.485, 0.456, 0.406], [0.229, 0.224, 0.225])]
    )
