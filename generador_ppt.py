import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os

# ----------------------------
# 파워포인트 생성 함수 (KR + ES)
# - 제목 슬라이드 1장
# - 가사 슬라이드: KR 한 줄 + 바로 아래 ES 한 줄
# ----------------------------
def crear_ppt(titulos_kr, bloques_dict_kr, bloques_dict_es, secuencia, estilos, resaltados):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    for i, titulo in enumerate(titulos_kr):
        # ---------- 제목 슬라이드 ----------
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(*estilos['bg_titulo'])

        tb = slide.shapes.add_textbox(Inches(1), Inches(estilos['altura_texto']), Inches(11.33), Inches(3))
        tf = tb.text_frame
        tf.clear()
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        run1 = p1.add_run()
        run1.text = titulo
        run1.font.size = Pt(estilos['tamano_titulo_kr'])
        run1.font.color.rgb = RGBColor(*estilos['color_titulo_kr'])
        p1.alignment = PP_ALIGN.CENTER

        # ---------- 가사 슬라이드 ----------
        for bloque_id in secuencia[i]:
            kr_lines = bloques_dict_kr[i].get(bloque_id, [])
            es_lines = bloques_dict_es[i].get(bloque_id, [])

            # KR 라인 수 기준으로 돌리되, ES는 없으면 빈칸 처리
            for j in range(len(kr_lines)):
                linea_kr = kr_lines[j]
                linea_es = es_lines[j] if j < len(es_lines) else ""

                slide = prs.slides.add_slide(prs.slide_layouts[6])
                slide.background.fill.solid()
                slide.background.fill.fore_color.rgb = RGBColor(*estilos['bg_letra'])

                # ✅ KR (윗줄)
                tb_kr = slide.shapes.add_textbox(
                    Inches(1),
                    Inches(estilos['altura_texto']),
                    Inches(11.33),
                    Inches(1.5),
                )
                tf_kr = tb_kr.text_frame
                tf_kr.clear()
                tf_kr.word_wrap = True

                pkr = tf_kr.paragraphs[0]
                runkr = pkr.add_run()
                runkr.text = linea_kr
                runkr.font.size = Pt(estilos['tamano_letra_kr'])

                if bloque_id in resaltados[i]:
                    runkr.font.color.rgb = RGBColor(255, 192, 0)  # 노란색
                else:
                    runkr.font.color.rgb = RGBColor(*estilos['color_letra_kr'])

                pkr.alignment = PP_ALIGN.CENTER

                # ✅ ES (바로 아래)
                if linea_es.strip():
                    tb_es = slide.shapes.add_textbox(
                        Inches(1),
                        Inches(estilos['altura_texto'] + 1.8),
                        Inches(11.33),
                        Inches(1.5),
                    )
                    tf_es = tb_es.text_frame
                    tf_es.clear()
                    tf_es.word_wrap = True

                    pes = tf_es.paragraphs[0]
                    runes = pes.add_run()
                    runes.text = linea_es
                    runes.font.size = Pt(estilos.get('tamano_letra_es', estilos['tamano_letra_kr']))
                    runes.font.color.rgb = RGBColor(*estilos.get('color_letra_es', estilos['color_letra_kr']))
                    pes.alignment = PP_ALIGN.CENTER

    return prs


# --- Streamlit UI ---
st.set_page_config(layout="wide")
st.title("피피티 잘 부탁드립니당~")

col1, col2, col3, col4 = st.columns(4)

with col1:
    num_canciones = st.number_input("찬양 개수", min_value=1, max_value=10, step=1)

with col2:
    size_titulo_kr = st.number_input("제목 글자 크기", value=36)

with col3:
    size_letra_kr = st.number_input("가사 한국어 글자 크기", value=36)

with col4:
    size_letra_es = st.number_input("가사 스페인어 글자 크기", value=28)

altura_texto = st.slider("글자 위치 (0.0이 제일 높음)", 0.0, 6.0, value=0.5, step=0.1)

color_titulo_kr = "#000000"
bg_titulo = "#FFFFFF"
color_letra_kr = "#FFFFFF"
color_letra_es = "#FFFF00"
bg_letra = "#000000"

estilos = {
    'color_titulo_kr': tuple(int(color_titulo_kr[i:i+2], 16) for i in (1, 3, 5)),
    'bg_titulo': tuple(int(bg_titulo[i:i+2], 16) for i in (1, 3, 5)),
    'altura_texto': altura_texto,

    'color_letra_kr': tuple(int(color_letra_kr[i:i+2], 16) for i in (1, 3, 5)),
    'color_letra_es': tuple(int(color_letra_es[i:i+2], 16) for i in (1, 3, 5)),
    'bg_letra': tuple(int(bg_letra[i:i+2], 16) for i in (1, 3, 5)),

    'tamano_titulo_kr': size_titulo_kr,
    'tamano_letra_kr': size_letra_kr,
    'tamano_letra_es': size_letra_es,
}

korean_titles = []
bloques_por_cancion_kr, bloques_por_cancion_es = [], []
secuencias, resaltados = [], []

for i in range(num_canciones):
    st.subheader(f"🎵 찬양 {i+1}")
    titulo = st.text_input(f"한국어 [제목] #{i+1}", key=f"kr_title_{i}")
    korean_titles.append(titulo)

    # ✅ KR/ES 전체 가사 입력
    raw_lyrics_kr = st.text_area("KR 전체 가사 붙여넣기", key=f"bloques_all_kr_{i}")
    raw_lyrics_es = st.text_area("ES 전체 가사 붙여넣기", key=f"bloques_all_es_{i}")

    # KR 블록 파싱
    bloques_kr = {}
    current_block = None
    lines = raw_lyrics_kr.split("\n")
    for line in lines + [""]:
        if line.strip() == "":
            current_block = None
            continue
        if current_block is None:
            current_block = line.strip()
            bloques_kr[current_block] = []
        else:
            bloques_kr[current_block].append(line.strip())
    bloques_por_cancion_kr.append(bloques_kr)

    # ES 블록 파싱
    bloques_es = {}
    current_block = None
    lines = raw_lyrics_es.split("\n")
    for line in lines + [""]:
        if line.strip() == "":
            current_block = None
            continue
        if current_block is None:
            current_block = line.strip()
            bloques_es[current_block] = []
        else:
            bloques_es[current_block].append(line.strip())
    bloques_por_cancion_es.append(bloques_es)

    secuencia_str = st.text_input(
        f"슬라이드 순서 (예: A,A,B,C), 띄어쓰기 없이, 대문자 소문자 예민, 쉼표로 분리",
        key=f"secuencia_{i}"
    )
    bloque_resaltado_str = st.text_input(
        f"후렴 블록들 입력 (쉼표로 분리)",
        key=f"resaltado_{i}"
    )

    bloques_resaltados = [b.strip() for b in bloque_resaltado_str.split(",") if b.strip()]
    resaltados.append(bloques_resaltados)

    # ✅ 순서는 KR 블록 기준으로 검증 (ES는 없어도 빈칸으로 나옴)
    secuencia = [s.strip() for s in secuencia_str.split(",") if s.strip() in bloques_kr]
    secuencias.append(secuencia)

if st.button("완료!"):
    ppt = crear_ppt(
        korean_titles,
        bloques_por_cancion_kr,
        bloques_por_cancion_es,
        secuencias,
        estilos,
        resaltados
    )

    ppt_path = "ppt_generado.pptx"
    ppt.save(ppt_path)

    with open(ppt_path, "rb") as f:
        st.download_button("PPT 다운로드", f, file_name=ppt_path)

    if os.path.exists(ppt_path):
        os.remove(ppt_path)
