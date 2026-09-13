import numpy as np
from sklearn.metrics import (
    accuracy_score,
    precision_score,
    recall_score,
    f1_score,
    roc_auc_score,
    confusion_matrix,
)


def binary_metrics(labels, probabilities, threshold=0.5):
    labels = np.asarray(labels, dtype=int)
    probabilities = np.asarray(probabilities, dtype=float)
    if not len(labels) or len(labels) != len(probabilities):
        raise ValueError("Labels and probabilities must have equal nonzero length.")
    if not np.isfinite(probabilities).all():
        raise ValueError("Non-finite predictions.")
    if not np.isin(labels, [0, 1]).all() or np.any((probabilities < 0) | (probabilities > 1)):
        raise ValueError("Expected binary labels and probabilities in [0,1].")
    if not 0 <= threshold <= 1:
        raise ValueError("Threshold must be in [0,1].")
    predicted = probabilities >= threshold
    tn, fp, fn, tp = confusion_matrix(labels, predicted, labels=[0, 1]).ravel()
    return dict(
        accuracy=float(accuracy_score(labels, predicted)),
        precision=float(precision_score(labels, predicted, zero_division=0)),
        recall=float(recall_score(labels, predicted, zero_division=0)),
        specificity=float(tn / (tn + fp)) if tn + fp else None,
        f1=float(f1_score(labels, predicted, zero_division=0)),
        auc=float(roc_auc_score(labels, probabilities)) if len(np.unique(labels)) == 2 else None,
        tn=int(tn),
        fp=int(fp),
        fn=int(fn),
        tp=int(tp),
    )
