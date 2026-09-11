"""由两份 PPT 的渲染画面生成 PDF 与逐页 PNG（兼容性兜底格式）。

Keynote / 预览 / 任何浏览器都能直接打开 PDF，不存在 OOXML 兼容问题。
"""
import sys, os, glob
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deckkit import Deck, contact_sheet
from PIL import Image

PPI = 200          # PDF 渲染精度
ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")

# 复用两个构建脚本的幻灯片定义，只做渲染
import importlib.util
for mod, tag, deckname in [("build_deck_a", "a", "PPT-01"),
                           ("build_deck_b", "b", "PPT-02")]:
    spec = importlib.util.spec_from_file_location(
        mod, os.path.join(os.path.dirname(os.path.abspath(__file__)), f"{mod}.py"))
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)          # 会顺带重建 pptx 与低分辨率预览
    deck = m.d
    paths, probs = deck.render(f"/tmp/pdf_{tag}", px_per_in=PPI)
    assert not probs, probs

    # 逐页 PNG 打包
    zdir = os.path.join(ROOT, "download", f"{deckname}-slides-png")
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
