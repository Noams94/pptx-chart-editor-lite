"""PPTX Chart Editor Lite — עריכת כל הגרפים במצגת דרך אקסל בלבד."""

import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st

SAMPLE_PATH = Path(__file__).parent / "sample.pptx"

from core.data_extractor import extract_all_charts
from core.data_writer import update_multiple_charts
from core.text_extractor import extract_all_texts
from core.text_writer import update_multiple_texts

TEXTS_SHEET_NAME = "_Texts"
TEXT_COLS = [
    "שקופית",
    "כותרת_שקופית",
    "שם_צורה",
    "מזהה_צורה",
    "סוג",
    "מיקום",
    "טקסט_מקורי",
    "טקסט_חדש",
]


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

with st.expander("📖 מדריך שימוש"):
    st.markdown("""
**מה האפליקציה עושה?**
עריכת כל הגרפים **והטקסטים** במצגת PowerPoint דרך קובץ אקסל אחד — גיליון נפרד לכל גרף, וגיליון מאוחד לכל הטקסטים.

**איך זה עובד?**

1. **העלאת מצגת** — בחרו קובץ `.pptx`, או נסו את המצגת לדוגמה.
2. **הורדת אקסל** — יורד קובץ אקסל הכולל גיליון לכל גרף, וגיליון `_Texts` עם כל הטקסטים במצגת (כותרות, תיבות טקסט, טבלאות).
3. **עריכה באקסל**:
   - **גרפים** — שנו ערכים, שמות סדרות וקטגוריות בגיליון של כל גרף.
   - **טקסטים** — בגיליון `_Texts`, ערכו את עמודת **טקסט_חדש** בלבד. אל תשנו את שאר העמודות (הן משמשות לזיהוי).
4. **העלאת האקסל המתוקן** — חזרו לאפליקציה והעלו את הקובץ הערוך.
5. **הורדת המצגת המעודכנת** — תקבלו קובץ `.pptx` חדש עם הנתונים והטקסטים המעודכנים.

**מה חשוב לדעת?**

- העיצוב של המצגת, הגרפים והטקסטים נשמר — רק התוכן מתעדכן.
- אין לשנות את שמות הגיליונות באקסל ואת עמודות הזיהוי בגיליון הטקסטים.
- ניתן להוסיף שורות (קטגוריות) או לשנות ערכים, אך לא להפחית את מספר העמודות המקוריות בגרפים.
- שורות שבהן **טקסט_חדש** ריק או זהה ל**טקסט_מקורי** ידולגו.
- גיליונות או שורות שלא תואמים למצגת יסומנו באזהרה וידולגו.
""")

st.divider()

# === שלב 1: העלאת מצגת ===
st.subheader("1. העלאת מצגת")

uploaded = st.file_uploader(
    "בחרו קובץ PPTX",
    type=["pptx"],
    label_visibility="collapsed",
)


def _load_pptx(pptx_bytes: bytes, file_name: str):
    st.session_state.pptx_bytes = pptx_bytes
    st.session_state.file_name = file_name
    with st.spinner("טוען נתוני גרפים וטקסטים..."):
        st.session_state.charts = extract_all_charts(pptx_bytes)
        st.session_state.texts = extract_all_texts(pptx_bytes)
    st.session_state.sheet_map = _build_sheet_name_map(st.session_state.charts)
    st.session_state.pop("updated_bytes", None)
    st.session_state.pop("xl_export_bytes", None)


if uploaded is not None and st.session_state.get("file_name") != uploaded.name:
    _load_pptx(uploaded.getvalue(), uploaded.name)
    st.rerun()

if "pptx_bytes" not in st.session_state and SAMPLE_PATH.exists():
    st.caption("אין לכם מצגת בהישג יד? נסו את המצגת לדוגמה:")
    if st.button("📂 שימוש במצגת לדוגמה", use_container_width=True):
        _load_pptx(SAMPLE_PATH.read_bytes(), "sample.pptx")
        st.rerun()

if "pptx_bytes" not in st.session_state:
    st.info("העלו קובץ מצגת כדי להתחיל.")
    st.stop()

charts = st.session_state.charts
texts = st.session_state.get("texts", [])
if not charts and not texts:
    st.warning("לא נמצאו גרפים או טקסטים לעריכה במצגת.")
    if st.button("התחלה מחדש"):
        _reset()
        st.rerun()
    st.stop()

slides_with_charts = len({c.slide_index for c in charts})
m1, m2, m3 = st.columns(3)
m1.metric("שקופיות עם גרפים", slides_with_charts)
m2.metric("סך כל הגרפים", len(charts))
m3.metric("פריטי טקסט", len(texts))

st.success(f"נטען בהצלחה: {st.session_state.file_name}")
st.divider()

# === שלב 2: הורדת אקסל ===
st.subheader("2. הורדת קובץ אקסל לעריכה")
st.caption("כל גרף בגיליון נפרד. ערכו את הערכים ושמרו את הקובץ.")

if "xl_export_bytes" not in st.session_state:
    xl_buffer = io.BytesIO()
    with pd.ExcelWriter(xl_buffer, engine="openpyxl") as writer:
        if texts:
            text_rows = [
                {
                    "שקופית": t.slide_index + 1,
                    "כותרת_שקופית": t.slide_title,
                    "שם_צורה": t.shape_name,
                    "מזהה_צורה": t.shape_id,
                    "סוג": t.kind,
                    "מיקום": t.location_label,
                    "טקסט_מקורי": t.original_text,
                    "טקסט_חדש": t.original_text,
                }
                for t in texts
            ]
            pd.DataFrame(text_rows, columns=TEXT_COLS).to_excel(
                writer, sheet_name=TEXTS_SHEET_NAME, index=False,
            )
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
        text_updates = []
        skipped = []
        for sheet_name in xls.sheet_names:
            if sheet_name == TEXTS_SHEET_NAME:
                try:
                    tdf = pd.read_excel(xls, sheet_name=sheet_name)
                except Exception as e:
                    skipped.append(f"גיליון '{sheet_name}': שגיאת קריאה ({e})")
                    continue
                missing = [c for c in TEXT_COLS if c not in tdf.columns]
                if missing:
                    skipped.append(
                        f"גיליון '{sheet_name}': חסרות עמודות {missing}"
                    )
                    continue
                texts_by_key = {
                    (t.slide_index, int(t.shape_id), t.kind,
                     t.paragraph_index, t.table_row, t.table_col): t
                    for t in texts
                }
                for _, row in tdf.iterrows():
                    new_val = row["טקסט_חדש"]
                    orig_val = row["טקסט_מקורי"]
                    if pd.isna(new_val):
                        continue
                    new_str = str(new_val)
                    orig_str = "" if pd.isna(orig_val) else str(orig_val)
                    if new_str == orig_str:
                        continue
                    try:
                        slide_index = int(row["שקופית"]) - 1
                        shape_id = int(row["מזהה_צורה"])
                    except (ValueError, TypeError):
                        continue
                    kind = str(row["סוג"])
                    # Match back to TextInfo by location string to recover
                    # paragraph_index/table coordinates.
                    matched = None
                    for t in texts:
                        if (t.slide_index == slide_index and t.shape_id == shape_id
                                and t.kind == kind and t.location_label == str(row["מיקום"])):
                            matched = t
                            break
                    if matched is None:
                        skipped.append(
                            f"טקסט בשקופית {slide_index + 1} ({kind}): לא נמצאה התאמה"
                        )
                        continue
                    text_updates.append((
                        matched.slide_index, matched.shape_id, matched.shape_name,
                        matched.kind, matched.paragraph_index,
                        matched.table_row, matched.table_col, new_str,
                    ))
                continue

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

        if updates or text_updates:
            parts = []
            if updates:
                parts.append(f"{len(updates)} גרפים")
            if text_updates:
                parts.append(f"{len(text_updates)} טקסטים")
            st.info("מוכנים לעדכון " + " ו-".join(parts) + ".")
            if st.button("✏️ החלת השינויים על המצגת", type="primary", use_container_width=True):
                with st.spinner("מעדכן את המצגת..."):
                    current_bytes = st.session_state.pptx_bytes
                    if updates:
                        current_bytes = update_multiple_charts(current_bytes, updates)
                    if text_updates:
                        current_bytes = update_multiple_texts(current_bytes, text_updates)
                    st.session_state.updated_bytes = current_bytes
                    st.success("עודכנו " + " ו-".join(parts) + ".")
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
