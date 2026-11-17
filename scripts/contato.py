import pandas as pd, math, sys
from PIL import Image, ImageDraw, ImageFont, ImageOps
df = pd.read_pickle("fotos/indice.pkl"); df["uf"] = df.colab.str.extract(r"Uf:\s*(\w+)")[0]
df["id"] = df.caderno.str[:3].str.upper() + df.slide.astype(str) + "." + df.idx.astype(str)
df.to_pickle("fotos/indice.pkl")
font = ImageFont.truetype("arial.ttf", 14)
def folha(sub, nome, cols=8, tw=230, th=172):
    rows = math.ceil(len(sub) / cols); W = Image.new("RGB", (cols * tw, rows * (th + 34)), "white"); d = ImageDraw.Draw(W)
    for k, r in enumerate(sub.itertuples()):
        x, y = (k % cols) * tw, (k // cols) * (th + 34)
        im = ImageOps.exif_transpose(Image.open(r.arquivo)).convert("RGB"); im.thumbnail((tw - 6, th - 4))
        W.paste(im, (x + 3, y + 2))
        d.text((x + 3, y + th), f"{r.id} {r.uf}", fill="black", font=font); d.text((x + 3, y + th + 16), str(r.loja)[:28], fill="#555", font=font)
    W.save(f"render/{nome}.jpg", quality=80); print(nome, len(sub))
for k in ["parceiro", "mpdv", "expo", "rack", "testeira"]:
    s = df[df.caderno == k]
    if k == "parceiro": s = s[s.idx == 0]
    folha(s, f"folha_{k}")
t = df[(df.caderno == "treino")]
folha(t[t.idx == 0].iloc[:48], "folha_treino1"); folha(t[t.idx == 0].iloc[48:], "folha_treino2")
