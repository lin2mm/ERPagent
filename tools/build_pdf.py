"""由幻灯片定义生成 PDF 与逐页 PNG（兼容性兜底格式）。

Keynote / 预览 / 任何浏览器都能直接打开 PDF，不存在 OOXML 兼容问题。
PDF 与 PPTX 来自同一份幻灯片定义，因此内容完全一致。

用法：
    python tools/build_pdf.py            # 全部四份（含两份原 PPT 与两份售前版）
    python tools/build_pdf.py pitch      # 只出两份售前说服版（老板决策简报）
    python tools/build_pdf.py deck       # 只出两份机制版方案
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deckkit import audit_pptx
from PIL import Image
import importlib.util

PPI = 200          # PDF 渲染精度
ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")

# (构建模块, 临时目录标签, 文件基名, PNG 子目录名, 分组)
DECKS = [
    ("build_pitch_a", "pitch_a", "PITCH-01-Roll-Material-MES-Boss-Briefing", "PITCH-01-slides-png", "pitch"),
    ("build_pitch_b", "pitch_b", "PITCH-02-Custom-Door-Boss-Briefing", "PITCH-02-slides-png", "pitch"),
    ("build_deck_a", "a", "PPT-01", "PPT-01-slides-png", "deck"),
    ("build_deck_b", "b", "PPT-02", "PPT-02-slides-png", "deck"),
]

want = sys.argv[1] if len(sys.argv) > 1 else "all"

for mod, tag, deckname, pngdir, group in DECKS:
    if want not in ("all", group):
        continue
    spec = importlib.util.spec_from_file_location(
        mod, os.path.join(os.path.dirname(os.path.abspath(__file__)), f"{mod}.py"))
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)          # 会顺带重建 pptx 与低分辨率预览
    deck = m.d

    # 逐页 PNG 打包
    paths, probs = deck.render(f"/tmp/pdf_{tag}", px_per_in=PPI)
    assert not probs, probs
    zdir = os.path.join(ROOT, "download", pngdir)
    os.makedirs(zdir, exist_ok=True)
    for i, p in enumerate(paths, 1):
        Image.open(p).save(os.path.join(zdir, f"{deckname}-slide-{i:02d}.png"))

    # PDF（每页 13.33 × 7.5 英寸）
    ims = [Image.open(p).convert("RGB") for p in paths]
    pdf = os.path.join(ROOT, "pdf", f"{deckname}.pdf")
    os.makedirs(os.path.dirname(pdf), exist_ok=True)
    ims[0].save(pdf, "PDF", resolution=PPI, save_all=True, append_images=ims[1:])
    print(f"{deckname}: {len(ims)} 页 → {pdf}  ({os.path.getsize(pdf)/1024/1024:.1f} MB)")
    print(f"          逐页 PNG → {zdir}")

    # 兼容性自检
    pptx = os.path.join(ROOT, "ppt", f"{deckname}.pptx")
    if os.path.exists(pptx):
        probs = audit_pptx(pptx)
        print(f"          PPTX 兼容性: {'✓ 通过' if not probs else probs}")
