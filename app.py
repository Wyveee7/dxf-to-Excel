import streamlit as st
import ezdxf
import pandas as pd
import tempfile
import os


def extract_text_entities(dxf_path):
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    texts = []
    for entity in msp:
        if entity.dxftype() in ["TEXT", "MTEXT"]:
            try:
                text = entity.plain_text() if entity.dxftype() == "MTEXT" else entity.text
                x, y = round(entity.dxf.insert.x, 2), round(entity.dxf.insert.y, 2)
                texts.append((x, y, text.strip()))
            except Exception:
                pass

    return texts


def group_into_table(texts, tolerance=1.0):
    from collections import defaultdict

    rows = defaultdict(list)
    for x, y, text in texts:
        matched = False
        for ry in rows:
            if abs(ry - y) < tolerance:
                rows[ry].append((x, text))
                matched = True
                break
        if not matched:
            rows[y].append((x, text))

    ordered_rows = sorted(rows.items(), key=lambda r: -r[0])

    table = []
    for _, row_items in ordered_rows:
        row = [text for _, text in sorted(row_items, key=lambda item: item[0])]
        table.append(row)

    return table


# Streamlit App
st.set_page_config(page_title="Extrair Tabela do DXF", layout="centered")
st.title("📐 Extrator de Tabelas DXF para Excel")

uploaded_file = st.file_uploader("Faça upload do arquivo DXF", type=["dxf"])

if uploaded_file is not None:
    with tempfile.TemporaryDirectory() as tmpdir:
        file_path = os.path.join(tmpdir, uploaded_file.name)

        with open(file_path, "wb") as f:
            f.write(uploaded_file.read())

        st.info("⏳ Processando o arquivo...")

        texts = extract_text_entities(file_path)
        table = group_into_table(texts)

        if table:
            df = pd.DataFrame(table)
            st.success("✅ Tabela extraída com sucesso!")
            st.dataframe(df)

            # Exporta para Excel
            excel_path = os.path.join(tmpdir, "tabela_extraida.xlsx")
            df.to_excel(excel_path, index=False, header=False)

            with open(excel_path, "rb") as f:
                st.download_button(
                    label="📥 Baixar Excel",
                    data=f,
                    file_name="tabela_extraida.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("Nenhum texto foi detectado no arquivo.")
