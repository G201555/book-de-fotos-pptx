# Slide 17: troca o grafico vinculado (RACK ROKU TV.xlsx, pivô filtrado em T3-26) por um grafico nativo com Q4,
# calculado da mesma base (BASE, objetivo = nova instalacao, TRIMESTRE = T4-26, por BANDEIRA). Nao altera a planilha-base.
import warnings, pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.dml.color import RGBColor
from pptx.util import Pt
warnings.filterwarnings("ignore")
F = r"out\Roku Report W45 - 02.11 a 06.11 - com fotos.pptx"
TRI = "T4-26"
d = pd.read_excel(r"K:\JOB 194300 - ROKU\2026\2. RELATÓRIOS\Bases\RACK ROKU TV.xlsx", sheet_name="BASE")
i = d[d["Qual o objetivo da sua visita hoje nesta loja?"].astype(str).str.contains("nova instala", case=False) & (d.TRIMESTRE == TRI)]
v = i.BANDEIRA.value_counts()
print(TRI, int(v.sum()), v.to_dict())

p = Presentation(F); s = p.slides[16]
ole = [sh for sh in s.shapes if sh.shape_type == 10][0]
x, y, w, h = ole.left, ole.top, ole.width, ole.height
ole._element.getparent().remove(ole._element)

cd = CategoryChartData(); cd.categories = list(v.index); cd.add_series("Racks", list(v.values))
ch = s.shapes.add_chart(XL_CHART_TYPE.PIE, x, y, w, h, cd).chart
ch.has_title = False; ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.RIGHT; ch.legend.include_in_layout = False
ch.legend.font.size = Pt(9); ch.legend.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
CORES = ["156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E", "0E2841", "B54A0B", "3B7D23"]
plot = ch.plots[0]; plot.has_data_labels = True
dl = plot.data_labels; dl.show_percentage = True; dl.show_value = False; dl.number_format = "0%"; dl.number_format_is_linked = False
dl.position = XL_LABEL_POSITION.OUTSIDE_END; dl.font.size = Pt(9); dl.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
for k, pt in enumerate(plot.series[0].points):
    pt.format.fill.solid(); pt.format.fill.fore_color.rgb = RGBColor.from_string(CORES[k % len(CORES)])
    pt.format.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

n = 0
for sh in s.shapes:   # titulo do grafico
    if sh.has_text_frame and "RACKS POR REDE" in sh.text_frame.text:
        for par in sh.text_frame.paragraphs:
            for r in par.runs:
                if "Q3" in r.text: r.text = r.text.replace("Q3", "Q4"); n += 1
p.save(F); print("titulo trocado:", n)
