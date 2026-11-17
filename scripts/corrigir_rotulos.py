# Corrige rotulos de periodo desatualizados (slide 19: "EM W44" -> "EM W45"), mantendo a formatacao
from pptx import Presentation
F = r"out\Roku Report W45 - 02.11 a 06.11 - com fotos.pptx"
p = Presentation(F); n = 0
def walk(shs):
    global n
    for sh in shs:
        if sh.shape_type == 6: walk(sh.shapes); continue
        if sh.has_text_frame and "NOVAS TESTEIRAS" in sh.text_frame.text and "W44" in sh.text_frame.text:
            for par in sh.text_frame.paragraphs:
                for r in par.runs:
                    if "W44" in r.text: r.text = r.text.replace("W44", "W45"); n += 1
walk(p.slides[18].shapes)
p.save(F); print("trocas:", n)
