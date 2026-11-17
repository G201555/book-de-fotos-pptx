import sys, os
from pptx import Presentation
from pptx.util import Emu
D = r"K:\JOB 194300 - ROKU\2026\2. RELATÓRIOS\Semanal\11 - Novembro\W45"
for f in sorted(os.listdir(D)):
    if not f.endswith(".pptx"): continue
    p = Presentation(os.path.join(D, f))
    print("="*90); print(f, os.path.getsize(os.path.join(D,f))//1024, "KB | slides:", len(p.slides), "| size", p.slide_width, p.slide_height)
    for i, s in enumerate(p.slides, 1):
        pics = [sh for sh in s.shapes if sh.shape_type == 13]
        txt = " / ".join(sh.text_frame.text.replace("\n"," ")[:60] for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip())
        if i <= 40 or len(p.slides) < 60:
            print(f"  {i:3d} layout={s.slide_layout.name[:20]:20s} pics={len(pics):2d} shapes={len(s.shapes):2d} | {txt[:200]}")
