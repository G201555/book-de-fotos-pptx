# Extrai as fotos dos cadernos da W45 com loja/data/colaborador -> fotos/*.jpg + fotos/indice.pkl
import os, re, pandas as pd, io
from pptx import Presentation
from PIL import Image
O = "origem"; os.makedirs("fotos", exist_ok=True)
DECKS = {"treino": "FOTO TREINAMENTO.pptx", "parceiro": "ROKU TV - EXPOSIÇÃO PARCEIRO.pptx", "mpdv": "ROKU TV - MPDV.pptx",
         "expo": "Exposição Especial.pptx", "rack": "Rack Roku TV.pptx", "testeira": "MOVIMENTAÇÕES DE TESTEIRAS.pptx"}
rows = []
for k, f in DECKS.items():
    p = Presentation(os.path.join(O, f))
    for i, s in enumerate(p.slides, 1):
        pics = [sh for sh in s.shapes if sh.shape_type == 13]
        if not pics: continue
        txt = " ".join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame).replace("\n", " ")
        m = re.search(r"Data:\s*(\S+)\s+Loja:\s*(\d+)\s*-\s*(.*?)\s+Colaborador\(a\):\s*(.*)$", txt)
        for j, pic in enumerate(pics):
            blob = pic.image.blob; ext = pic.image.ext
            im = Image.open(io.BytesIO(blob)); w, h = im.size
            fn = f"fotos/{k}_s{i:03d}_{j}.{ext}"; open(fn, "wb").write(blob)
            rows.append(dict(caderno=k, slide=i, idx=j, arquivo=fn, w=w, h=h, data=m.group(1) if m else None, cod_loja=int(m.group(2)) if m else None,
                             loja=m.group(3).strip() if m else None, colab=m.group(4).strip() if m else None, texto=txt[:160]))
df = pd.DataFrame(rows); df.to_pickle("fotos/indice.pkl")
pd.set_option("display.width", 250); pd.set_option("display.max_colwidth", 60); pd.set_option("display.max_rows", 50)
print(df.groupby("caderno").agg(fotos=("arquivo", "size"), lojas=("cod_loja", "nunique"), sem_meta=("loja", lambda s: s.isna().sum())))
print(df[df.caderno != "treino"][["caderno", "slide", "idx", "w", "h", "data", "cod_loja", "loja", "colab"]].head(40).to_string(index=False))
print(df.texto.iloc[0])
