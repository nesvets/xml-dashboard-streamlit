# app.py
import io
import requests
import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET

# Для "фильтры как в таблицах"
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode

FEED_URL = "https://platinumlist.net/xml-feed/partnership-program"

st.set_page_config(page_title="XML → Table", layout="wide")
st.title("XML feed → таблица с фильтрами и выгрузкой в Excel")

@st.cache_data(ttl=300)  # кэш на 5 минут, чтобы не дергать XML слишком часто
def load_xml(url: str) -> ET.Element:
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return ET.fromstring(r.content)

def flatten_element(el: ET.Element, prefix: str = "") -> dict:
    """
    Разворачивает XML-элемент в плоский dict:
    - Листовые теги становятся колонками
    - Атрибуты попадают как prefix@attr
    - Вложенность кодируется через / (path)
    """
    row = {}

    # Атрибуты
    for k, v in el.attrib.items():
        row[f"{prefix}@{k}"] = v

    children = list(el)
    if not children:
        text = (el.text or "").strip()
        if prefix:
            row[prefix.rstrip("/")] = text
        return row

    # Если есть дети — рекурсивно собираем
    for ch in children:
        ch_prefix = f"{prefix}{ch.tag}/"
        ch_children = list(ch)

        if ch_children:
            row.update(flatten_element(ch, ch_prefix))
        else:
            key = f"{prefix}{ch.tag}"
            val = (ch.text or "").strip()

            # если такой ключ уже был (повторяющиеся теги), аккуратно склеим
            if key in row and row[key] != "":
                row[key] = f"{row[key]} | {val}"
            else:
                row[key] = val

        # атрибуты детей тоже добавим
        for ak, av in ch.attrib.items():
            row[f"{prefix}{ch.tag}@{ak}"] = av

    return row

def find_repeating_nodes(root: ET.Element) -> list[ET.Element]:
    """
    Пытается найти "строки" — узлы, которые повторяются (типичный паттерн: offers/offer и т.п.).
    Берём самый массово повторяющийся уровень.
    """
    from collections import defaultdict
    buckets = defaultdict(list)

    for parent in root.iter():
        tags = [c.tag for c in list(parent)]
        if not tags:
            continue
        # группируем детей по тегу
        counts = {}
        for t in tags:
            counts[t] = counts.get(t, 0) + 1
        # если есть повторяющийся тег — вероятно это "строки"
        for t, cnt in counts.items():
            if cnt >= 2:
                buckets[(parent.tag, t)].extend(parent.findall(t))

    if not buckets:
        # fallback: берём прямых детей рута как строки
        return list(root)

    # выбираем группу с максимальным количеством элементов
    best = max(buckets.items(), key=lambda kv: len(kv[1]))
    return best[1]

def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    return output.getvalue()

with st.sidebar:
    st.subheader("Источник")
    url = st.text_input("XML URL", FEED_URL)
    refresh = st.button("Обновить (сбросить кэш)")

if refresh:
    st.cache_data.clear()

try:
    root = load_xml(url)
except Exception as e:
    st.error(f"Не удалось загрузить XML: {e}")
    st.stop()

rows_nodes = find_repeating_nodes(root)
rows = [flatten_element(n, prefix=f"{n.tag}/") for n in rows_nodes]
df = pd.DataFrame(rows).fillna("")

st.caption(f"Строк: {len(df)} • Колонок: {len(df.columns)}")

# Таблица с фильтрами
gb = GridOptionsBuilder.from_dataframe(df)
gb.configure_default_column(filter=True, sortable=True, resizable=True)
gb.configure_grid_options(domLayout="normal")
grid_options = gb.build()

grid_response = AgGrid(
    df,
    gridOptions=grid_options,
    data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
    update_mode=GridUpdateMode.MODEL_CHANGED,
    fit_columns_on_grid_load=False,
    height=600,
)

df_filtered = grid_response["data"]

col1, col2 = st.columns(2)
with col1:
    st.download_button(
        "Скачать ВСЁ (Excel)",
        data=df_to_xlsx_bytes(df),
        file_name="xml_full.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with col2:
    st.download_button(
        "Скачать ОТФИЛЬТРОВАННОЕ (Excel)",
        data=df_to_xlsx_bytes(df_filtered),
        file_name="xml_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
