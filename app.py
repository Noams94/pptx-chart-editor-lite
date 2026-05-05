"""PPTX Chart Editor Lite — עריכת כל הגרפים במצגת דרך אקסל בלבד."""

import io
import re

import pandas as pd
import streamlit as st

from core.data_extractor import extract_all_charts
from core.data_writer import update_multiple_charts


st.set_page_config(
    page_title="עורך גרפים למצגות",
    page_icon="📊",
    layout="centered",
)

st.markdown("""
<style>
.stApp, .stApp * { direction: rtl; text-align: right; }
.stApp [data-testid="stFileUploader"] section { direction: rtl; }
.stApp h1, .stApp h2, .stApp h3 { text-align: right; }
.stApp .stMarkdown, .stApp .stCaption { text-align: right; }
.stApp [data-testid="stMetricValue"], .stApp [data-testid="stMetricLabel"] { direction: ltr; text-align: center; }
.stApp [data-testid="stDataFrame"] { direction: ltr; }
.stApp button { direction: rtl; }
</style>
""", unsafe_allow_html=True)


def _sanitize_sheet_name(slide_index: int, shape_name: str) -> str:
    prefix = f"Slide{slide_index + 1}_"
    clean_name = re.sub(r'[\[\]:*?/\\]', '', shape_name)
    max_name_len = 31 - len(prefix)
    return prefix + clean_name[:max_name_len]


def _build_sheet_name_map(charts_list) -> dict:
    name_map = {}
    seen = set()
    for ci in charts_list:
        sheet = _sanitize_sheet_name(ci.slide_index, ci.shape_name)
        base = sheet
        counter = 1
        while sheet in seen:
            sheet = base[:29] + f"_{counter}"
            counter += 1
        seen.add(sheet)
        name_map[ci.key] = sheet
    return name_map


def _reset():
    for key in list(st.session_state.keys()):
        del st.session_state[key]


st.title("📊 עורך גרפים למצגות")
st.caption("העלאת מצגת ← הורדת אקסל ← עריכת הנתונים ← העלאת אקסל מתוקן ← הורדת מצגת מעודכנת")
st.divider()

# === שלב 1: העלאת מצגת ===
st.subheader("1. העלאת מצגת")

uploaded = st.file_uploader(
    "בחרו קובץ PPTX",
    type=["pptx"],
    label_visibility="collapsed",
)

if uploaded is not None and st.session_state.get("file_name") != uploaded.name:
    st.session_state.pptx_bytes = uploaded.getvalue()
    st.session_state.file_name = uploaded.name
    with st.spinner("טוען נתוני גרפים..."):
        st.session_state.charts = extract_all_charts(st.session_state.pptx_bytes)
    st.session_state.sheet_map = _build_sheet_name_map(st.session_state.charts)
    st.session_state.pop("updated_bytes", None)
    st.session_state.pop("xl_export_bytes", None)
    st.rerun()

if "pptx_bytes" not in st.session_state:
    st.info("העלו קובץ מצגת כדי להתחיל.")
    st.stop()

charts = st.session_state.charts
if not charts:
    st.warning("לא נמצאו גרפים במצגת.")
    if st.button("התחלה מחדש"):
        _reset()
        st.rerun()
    st.stop()

slides_with_charts = len({c.slide_index for c in charts})
m1, m2 = st.columns(2)
m1.metric("שקופיות עם גרפים", slides_with_charts)
m2.metric("סך כל הגרפים", len(charts))

st.success(f"נטען בהצלחה: {st.session_state.file_name}")
st.divider()

# === שלב 2: הורדת אקסל ===
st.subheader("2. הורדת קובץ אקסל לעריכה")
st.caption("כל גרף בגיליון נפרד. ערכו את הערכים ושמרו את הקובץ.")

if "xl_export_bytes" not in st.session_state:
    xl_buffer = io.BytesIO()
    with pd.ExcelWriter(xl_buffer, engine="openpyxl") as writer:
        for ci in charts:
            sheet = st.session_state.sheet_map[ci.key]
            ci.dataframe.to_excel(writer, sheet_name=sheet, index=False)
    st.session_state.xl_export_bytes = xl_buffer.getvalue()

base_name = st.session_state.file_name.replace(".pptx", "")
st.download_button(
    label="⬇️ הורדת אקסל",
    data=st.session_state.xl_export_bytes,
    file_name=f"charts_{base_name}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
    type="primary",
)
st.divider()

# === שלב 3: העלאת אקסל מתוקן ===
st.subheader("3. העלאת אקסל מתוקן")

xl_file = st.file_uploader(
    "העלאת אקסל ערוך",
    type=["xlsx"],
    key="xl_import",
    label_visibility="collapsed",
)

if xl_file is not None:
    try:
        xls = pd.ExcelFile(xl_file, engine="openpyxl")
        sheet_to_chart = {v: k for k, v in st.session_state.sheet_map.items()}
        charts_by_key = {ci.key: ci for ci in charts}

        updates = []
        skipped = []
        for sheet_name in xls.sheet_names:
            chart_key = sheet_to_chart.get(sheet_name)
            if chart_key and chart_key in charts_by_key:
                ci = charts_by_key[chart_key]
                imported_df = pd.read_excel(xls, sheet_name=sheet_name)
                expected_cols = len(ci.dataframe.columns)
                if len(imported_df.columns) < expected_cols:
                    skipped.append(
                        f"גיליון '{sheet_name}': חסרות עמודות "
                        f"(צפויות {expected_cols}, נמצאו {len(imported_df.columns)})"
                    )
                    continue
                imported_df.columns = (
                    list(ci.dataframe.columns)
                    + list(imported_df.columns[expected_cols:])
                )
                updates.append((
                    ci.slide_index,
                    ci.shape_name,
                    imported_df,
                    ci.is_xy,
                    ci.series_formats,
                    ci.shape_id,
                ))
            else:
                skipped.append(f"גיליון '{sheet_name}' לא נמצאה התאמה למצגת")

        if skipped:
            with st.expander(f"⚠️ אזהרות ({len(skipped)})"):
                for msg in skipped:
                    st.warning(msg)

        if updates:
            st.info(f"מוכנים לעדכון {len(updates)} גרפים.")
            if st.button("✏️ החלת השינויים על המצגת", type="primary", use_container_width=True):
                with st.spinner(f"מעדכן {len(updates)} גרפים..."):
                    updated_bytes = update_multiple_charts(
                        st.session_state.pptx_bytes, updates,
                    )
                    st.session_state.updated_bytes = updated_bytes
                    st.success(f"עודכנו {len(updates)} גרפים.")
                    st.rerun()
        elif not skipped:
            st.error("לא נמצאו גיליונות תואמים בקובץ האקסל.")
    except Exception as e:
        st.error(f"שגיאה בקריאת הקובץ: {e}")

# === שלב 4: הורדת מצגת מעודכנת ===
if "updated_bytes" in st.session_state:
    st.divider()
    st.subheader("4. הורדת המצגת המעודכנת")
    st.download_button(
        label="⬇️ הורדת מצגת מעודכנת",
        data=st.session_state.updated_bytes,
        file_name=f"updated_{st.session_state.file_name}",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
        type="primary",
    )

st.divider()
if st.button("🔄 התחלה מחדש"):
    _reset()
    st.rerun()
