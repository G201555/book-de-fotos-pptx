# Insere as fotos da semana nos slides "BOOK DE FOTOS" do Roku Report e atualiza as legendas -> out/<nome> - com fotos.pptx
import os, sys, copy, pandas as pd
from PIL import Image, ImageOps
from pptx import Presentation
from pptx.util import Cm

SRC = r"origem\Roku Report W45 - 02.11 a 06.11.pptx"
OUT = r"out\Roku Report W45 - 02.11 a 06.11 - com fotos.pptx"
os.makedirs("out", exist_ok=True); os.makedirs("fotos_prontas", exist_ok=True)
idx = pd.read_pickle("fotos/indice.pkl").set_index("id")

# slide -> lista de (id da foto, legenda nova); ordem: em cima/esquerda primeiro
ESCOLHA = {
    8: [("TRE22.0", "CASAS BAHIA - MG"), ("TRE36.0", "HAVAN - SP"), ("TRE68.0", "MAGAZINE LUIZA - PR"), ("TRE81.0", "QUERO-QUERO - RS"),
        ("TRE6.0", "BEMOL - AM"), ("TRE43.0", "LASER ELETRO - PE"), ("TRE2.0", "ARM. PARAÍBA - PB"), ("TRE45.0", "LOJAS CEM - SP")],
    12: [("PAR3.0", "ATACADÃO DOS ELETROS - CABEDELO - PB"), ("PAR41.0", "LASER ELETRO - CENTRO, PAULISTA - PE")],
    13: [("MPD20.0", "MAGAZINE LUIZA - PAULISTA - PE"), ("MPD25.0", "MILLENA - PAULISTA NORTH WAY - PE")],
    15: [("EXP3.1", "TOP MÓVEIS - JOAQUIM BEZERRA - CE"), ("EXP2.0", "QUERO-QUERO - REV. AMÉRICO VESPÚCIO - RS")],
    17: [("RAC9.0", "SOLAR MAGAZINE - BARÃO 2 - CE"), ("RAC4.0", "LASER ELETRO - ELIZEU CÉSAR - PB")],
    19: [("TES2.0", "ATACADÃO - ALMIRANTE BARROSO - PB"), ("TES3.0", "LOJAS LEBES - CENTRO VIAMÃO 2 - RS")],
}


def preparar(fid):
    """Aplica a rotacao EXIF e reduz para 1600 px, para o PowerPoint mostrar certo e o arquivo nao inchar."""
    src = idx.loc[fid, "arquivo"]; dst = f"fotos_prontas/{fid}.jpg"
    im = ImageOps.exif_transpose(Image.open(src)).convert("RGB"); im.thumbnail((1600, 1600)); im.save(dst, quality=85)
    return dst, im.size


def colocar(slide, arq, tam, x, y, w, h):
    """Insere a foto preenchendo a caixa (x,y,w,h) e corta o excesso pelo centro."""
    pic = slide.shapes.add_picture(arq, x, y, w, h)
    iw, ih = tam; box, img = w / h, iw / ih
    if img > box:   # foto mais larga: corta laterais
        c = (1 - box / img) / 2; pic.crop_left = c; pic.crop_right = c
    else:           # foto mais alta: corta em cima e embaixo
        c = (1 - img / box) / 2; pic.crop_top = c; pic.crop_bottom = c
    return pic


def set_texto(shape, novo):
    """Troca o texto mantendo a formatacao do primeiro run."""
    tf = shape.text_frame; p = tf.paragraphs[0]
    runs = p.runs
    runs[0].text = novo
    for r in runs[1:]: r.text = ""
    for extra in tf.paragraphs[1:]:
        for r in extra.runs: r.text = ""


prs = Presentation(SRC)
G = Cm(0.15)   # folga entre foto e legenda
for n, fotos in ESCOLHA.items():
    s = prs.slides[n - 1]
    if n == 8:
        caps = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip() and sh.top > Cm(5)]
        caps.sort(key=lambda sh: (round(sh.top / Cm(1)), sh.left))
        for sh in [sh for sh in s.shapes if sh.shape_type == 1 and sh.auto_shape_type == 5 and sh.width < Cm(5)]:
            sh._element.getparent().remove(sh._element)   # quadrados roxos que sobraram do modelo
        linha1 = [c for c in caps if c.top < Cm(10)]; linha2 = [c for c in caps if c.top >= Cm(10)]
        topo = {1: Cm(2.6), 2: linha1[0].top + linha1[0].height + Cm(0.6)}
        for k, (c, (fid, leg)) in enumerate(zip(linha1 + linha2, fotos)):
            y0 = topo[1] if c in linha1 else topo[2]
            arq, tam = preparar(fid)
            colocar(s, arq, tam, c.left, y0, c.width, c.top - y0 - G)
            set_texto(c, leg)
    else:
        todas = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip() and sh.left > Cm(23) and sh.top > Cm(4)]
        caps = []
        for c in todas:   # legendas empilhadas na mesma posicao: fica a de cima (ultima na ordem), as escondidas saem
            mesmas = [o for o in todas if abs(o.top - c.top) < Cm(0.5)]
            if c is mesmas[-1]: caps.append(c)
            else: c._element.getparent().remove(c._element)
        caps.sort(key=lambda sh: sh.top)
        assert len(caps) == 2, (n, [c.text_frame.text for c in caps])
        y0s = [Cm(2.5), caps[0].top + caps[0].height + Cm(0.5)]
        for c, y0, (fid, leg) in zip(caps, y0s, fotos):
            arq, tam = preparar(fid)
            colocar(s, arq, tam, c.left, y0, c.width, c.top - y0 - G)
            set_texto(c, leg)
prs.save(OUT)
print("ok", OUT, os.path.getsize(OUT) // 1024, "KB")
