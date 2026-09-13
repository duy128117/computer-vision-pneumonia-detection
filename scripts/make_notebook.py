import json
from pathlib import Path

cells = []


def cell(kind, source):
    value = dict(cell_type=kind, metadata={}, source=source.splitlines(keepends=True))
    if kind == "code":
        value.update(execution_count=None, outputs=[])
    cells.append(value)


cell(
    "markdown",
    "# Pneumonia Classification — thực nghiệm thật\nChạy từ bản sao project trên máy hoặc Colab GPU. Notebook dùng cùng mã nguồn CLI; không chứa số liệu dựng sẵn. Trên Colab, đặt toàn bộ project trong `/content/tgmt` trước khi chạy. Chọn Runtime → GPU.",
)
cell(
    "code",
    "from pathlib import Path\nimport os, sys, subprocess\nROOT = Path('/content/tgmt') if Path('/content/tgmt').exists() else Path.cwd()\nif ROOT.name == 'notebooks': ROOT = ROOT.parent\nos.chdir(ROOT)\nassert Path('src/train.py').exists(), 'Cần đặt notebook trong project'\nos.environ['TORCH_HOME'] = str(ROOT / '.torch')\ndef run(*args): subprocess.run([sys.executable, *args], check=True)\n",
)
cell(
    "code",
    "run('-m', 'pip', 'install', '-r', 'requirements.txt')\nimport torch\nprint(torch.__version__, 'CUDA:', torch.cuda.is_available())",
)
cell(
    "markdown",
    "## Dataset và thống kê\nNguồn: Kermany et al. (2018), doi:10.17632/rscbjbr9sj.2. Chỉ dùng ảnh X-quang thật. Nếu download yêu cầu đăng nhập, tải ZIP vào `data/chest-xray-pneumonia.zip`.",
)
cell(
    "code",
    "run('scripts/download_data.py')\nrun('-m', 'src.dataset')\nimport pandas as pd\nfrom IPython.display import display, Image\ndisplay(pd.read_csv('results/dataset_counts.csv'))\ndisplay(Image(filename='results/dataset_distribution.png'))",
)
cell(
    "markdown",
    "## Train classifier rồi fine-tune\nCấu hình YAML điều khiển LR, batch, epochs, patience và weighting. Không điều chỉnh cấu hình dựa vào test. Nếu chạy lại, dùng output mới để giữ provenance.",
)
cell("code", "run('-m', 'src.train', '--config', 'configs/resnet50.yaml')")
cell("code", "run('-m', 'src.train', '--config', 'configs/densenet121.yaml')")
cell(
    "code",
    "for name in ['resnet50', 'densenet121']:\n    run('-m', 'src.evaluate', '--checkpoint', f'checkpoints/best_{name}.pth')\nrun('-m', 'src.report')\ndisplay(pd.read_csv('results/model_comparison.csv'))",
)
cell(
    "code",
    "for name in ['resnet50', 'densenet121']:\n    display(pd.read_csv(f'logs/history_{name}.csv'))\n    for suffix in ['loss_curve','accuracy_curve','precision_recall','f1_curve','auc_curve']:\n        display(Image(filename=f'results/plots/{name}_{suffix}.png'))\ndisplay(Image(filename='results/roc/model_comparison_roc.png'))",
)
cell(
    "markdown",
    "## Phân tích TP/TN/FP/FN\nKiểm tra vùng phổi, chữ/marker, viền và background. Không suy diễn heatmap thành bằng chứng tổn thương. Ghi nhận nhóm không có mẫu, không tạo case giả.",
)
cell(
    "code",
    "for name in ['resnet50','densenet121']:\n    for case in ['TP','TN','FP','FN']:\n        images = sorted(Path(f'results/gradcam/{name}').glob(f'{case}_*_overlay.png'))\n        print(name, case, 'Không có mẫu' if not images else '')\n        for image in images: display(Image(filename=str(image)))",
)
cell(
    "markdown",
    "## Báo cáo / demo\nBáo cáo: `results/report.md`; slide: `results/presentation.pptx`. Chạy demo ở terminal: `python -m streamlit run app/app.py`. Kết luận cần cân nhắc AUC, sensitivity, specificity, F1, FN và giới hạn một nguồn dữ liệu/một seed. Grad-CAM không thay thế chẩn đoán lâm sàng.",
)
Path("notebooks/analysis.ipynb").write_text(
    json.dumps(
        dict(
            cells=cells,
            metadata=dict(
                kernelspec=dict(display_name="Python 3", language="python", name="python3")
            ),
            nbformat=4,
            nbformat_minor=5,
        ),
        ensure_ascii=False,
        indent=2,
    ),
    encoding="utf-8",
)
