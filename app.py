import base64
import tempfile
import threading
import time
from pathlib import Path

import openpyxl
import streamlit as st

from check_links import (
    detect_preferred_column_from_rules,
    load_known_names,
    process_workbook,
    update_known_names_after_success,
)


st.set_page_config(
    page_title="Excel Link Cleaner",
    page_icon="🔗",
    layout="wide",
)


def render_app_chrome() -> None:
    panda_path = Path(__file__).with_name("assets") / "red_panda_w_p.png"
    panda_html = ""

    if panda_path.exists():
        panda_src = base64.b64encode(panda_path.read_bytes()).decode("ascii")
        panda_html = (
            '<div class="panda-wrap">'
            '<a href="https://github.com/Damnti" target="_blank" rel="noopener noreferrer">'
            f'<img src="data:image/png;base64,{panda_src}" alt="Red panda logo" />'
            "</a>"
            "</div>"
        )

    st.markdown(
        """
        <style>
        .stApp {
            background:
                radial-gradient(circle at top right, rgba(244, 208, 183, 0.32), transparent 24%),
                radial-gradient(circle at left 12% top 8%, rgba(255, 244, 229, 0.92), transparent 20%),
                linear-gradient(180deg, #fffdf9 0%, #fff7f0 100%);
            color: #40261d;
        }
        .stApp, .stApp p, .stApp label, .stApp span, .stApp div, .stApp li {
            color: #40261d;
        }
        .stApp h1, .stApp h2, .stApp h3 {
            color: #5b2c1f;
        }
        [data-testid="stFileUploader"] section {
            background: #fffaf5;
            border: 1px solid rgba(132, 82, 58, 0.16);
            border-radius: 18px;
        }
        [data-testid="stFileUploader"] small,
        [data-testid="stFileUploader"] span,
        [data-testid="stFileUploader"] label,
        [data-testid="stFileUploader"] p {
            color: #6a4637 !important;
        }
        [data-testid="stExpander"] details {
            background: rgba(255, 255, 255, 0.72);
            border: 1px solid rgba(132, 82, 58, 0.12);
            border-radius: 16px;
        }
        [data-testid="stExpander"] summary {
            color: #5b2c1f !important;
        }
        [data-testid="stExpander"] details summary {
            padding-top: 0.75rem !important;
            padding-bottom: 0.75rem !important;
        }
        [data-testid="stExpander"] details > div[role="button"] + div {
            padding-top: 0.2rem !important;
        }
        [data-baseweb="popover"] ul,
        [data-baseweb="select"] [role="listbox"] {
            background: #fffaf5 !important;
            color: #40261d !important;
            border: 1px solid rgba(132, 82, 58, 0.14) !important;
        }
        [data-baseweb="popover"] li,
        [data-baseweb="select"] [role="option"] {
            background: #fffaf5 !important;
            color: #40261d !important;
        }
        [data-baseweb="popover"] li[aria-selected="true"],
        [data-baseweb="select"] [role="option"][aria-selected="true"] {
            background: #fff1e3 !important;
            color: #5b2c1f !important;
        }
        [data-baseweb="popover"] li:hover,
        [data-baseweb="select"] [role="option"]:hover {
            background: #fde7d6 !important;
            color: #5b2c1f !important;
        }
        [data-testid="stAlertContainer"] {
            color: #40261d;
        }
        [data-testid="stAlertContainer"] * {
            color: inherit !important;
        }
        [data-testid="stInfo"] {
            background: #e9f1fb !important;
            border: 1px solid rgba(88, 133, 192, 0.14) !important;
        }
        [data-testid="stSuccess"] {
            background: #e6f6df !important;
            border: 1px solid rgba(104, 158, 98, 0.14) !important;
        }
        [data-testid="stWarning"] {
            background: #fff3df !important;
            border: 1px solid rgba(191, 136, 68, 0.16) !important;
        }
        [data-testid="stDataFrame"] {
            background: #fffaf5 !important;
            border: 1px solid rgba(132, 82, 58, 0.14) !important;
            border-radius: 16px !important;
            overflow: hidden;
        }
        [data-testid="stDataFrame"] * {
            color: #40261d !important;
        }
        [data-testid="stDataFrame"] [role="columnheader"] {
            background: #f7ece1 !important;
            color: #5b2c1f !important;
            border-bottom: 1px solid rgba(132, 82, 58, 0.14) !important;
        }
        [data-testid="stDataFrame"] [role="gridcell"] {
            background: #fffaf5 !important;
            border-top: 1px solid rgba(132, 82, 58, 0.08) !important;
        }
        [data-testid="stMetricValue"] {
            color: #5b2c1f;
        }
        [data-baseweb="input"] input,
        [data-baseweb="select"] input,
        [data-testid="stNumberInput"] input,
        [data-testid="stTextInput"] input {
            color: #40261d !important;
            background: #fffaf5 !important;
            -webkit-text-fill-color: #40261d !important;
            caret-color: #40261d !important;
        }
        [data-baseweb="input"],
        [data-baseweb="select"] > div,
        [data-testid="stNumberInput"] > div > div,
        [data-testid="stTextInput"] > div > div {
            background: #fffaf5 !important;
            border-color: rgba(132, 82, 58, 0.22) !important;
        }
        [data-testid="stNumberInput"] [data-baseweb="input"] {
            background: #fffaf5 !important;
            border-radius: 12px !important;
        }
        [data-testid="stNumberInput"] button,
        [data-testid="stTextInput"] button,
        [data-baseweb="select"] button {
            color: #6a4637 !important;
        }
        [data-testid="stNumberInput"] button {
            background: #fff1e3 !important;
            border-left: 1px solid rgba(132, 82, 58, 0.16) !important;
        }
        [data-testid="stNumberInput"] button:hover {
            background: #fde7d6 !important;
        }
        [data-testid="stNumberInput"] label,
        [data-testid="stTextInput"] label,
        [data-testid="stTextArea"] label,
        [data-testid="stSelectbox"] label,
        [data-testid="stRadio"] label,
        [data-testid="stCheckbox"] label {
            color: #5b2c1f !important;
        }
        [data-testid="stNumberInput"] svg,
        [data-testid="stSelectbox"] svg {
            fill: #6a4637 !important;
        }
        [data-testid="stNumberInput"] button svg {
            fill: #8b4b35 !important;
            width: 16px !important;
            height: 16px !important;
        }
        [data-testid="stTextArea"] textarea {
            background: #fffaf5 !important;
            color: #40261d !important;
            -webkit-text-fill-color: #40261d !important;
            caret-color: #40261d !important;
            border-radius: 12px !important;
        }
        [data-testid="stTextArea"] [data-baseweb="textarea"] {
            background: #fffaf5 !important;
            border-color: rgba(132, 82, 58, 0.22) !important;
            border-radius: 12px !important;
        }
        [data-testid="stCheckbox"],
        [data-testid="stRadio"] {
            padding: 2px 0;
        }
        [data-testid="stCheckbox"] label[data-baseweb="checkbox"],
        [data-testid="stRadio"] label[data-baseweb="radio"] {
            background: transparent !important;
            gap: 10px !important;
            padding: 2px 0 !important;
            border-radius: 0 !important;
            box-shadow: none !important;
            align-items: center !important;
        }
        [data-testid="stCheckbox"] label {
            background: transparent !important;
        }
        [data-testid="stCheckbox"] label p,
        [data-testid="stCheckbox"] label span,
        [data-testid="stRadio"] label p,
        [data-testid="stRadio"] label span {
            color: #5b2c1f !important;
            font-weight: 500;
            background: transparent !important;
            border: 0 !important;
            box-shadow: none !important;
            padding: 0 !important;
            margin: 0 !important;
        }
        [data-testid="stCheckbox"] input,
        [data-testid="stRadio"] input {
            accent-color: #d86c43 !important;
        }
        [data-testid="stRadio"] div[role="radiogroup"] {
            gap: 14px !important;
        }
        [data-testid="stRadio"] label > div:first-child {
            color: #d86c43 !important;
        }
        [data-testid="stCheckbox"] label p::selection,
        [data-testid="stRadio"] label p::selection,
        [data-testid="stCheckbox"] label span::selection,
        [data-testid="stRadio"] label span::selection {
            background: rgba(216, 108, 67, 0.18) !important;
        }
        [data-testid="stFileUploader"] button,
        [data-testid="stButton"] button,
        [data-testid="stDownloadButton"] button {
            background: #171c27 !important;
            color: #e07a4a !important;
            -webkit-text-fill-color: #e07a4a !important;
            border: 1px solid rgba(216, 108, 67, 0.82) !important;
            border-radius: 12px !important;
            box-shadow: none;
        }
        [data-testid="stFileUploader"] button:hover,
        [data-testid="stButton"] button:hover,
        [data-testid="stDownloadButton"] button:hover {
            background: #1e2430 !important;
            border-color: rgba(216, 108, 67, 1) !important;
            color: #f09062 !important;
            -webkit-text-fill-color: #f09062 !important;
        }
        [data-testid="stButton"] button[kind="secondary"] {
            background: #fff1e3 !important;
            color: #7a3f2d !important;
            border: 1px solid rgba(132, 82, 58, 0.18) !important;
            box-shadow: none;
        }
        [data-testid="stButton"] button[kind="secondary"]:hover {
            background: #fde7d6 !important;
            color: #6a3527 !important;
        }
        .hero-card {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 24px;
            padding: 18px 24px;
            margin-bottom: 8px;
            border: 1px solid rgba(170, 90, 60, 0.14);
            border-radius: 20px;
            background: rgba(255, 255, 255, 0.82);
            box-shadow: 0 14px 38px rgba(116, 67, 43, 0.08);
        }
        .hero-copy {
            padding-top: 4px;
        }
        .hero-copy h1 {
            margin: 0 0 6px 0;
            color: #5b2c1f;
            font-size: 2rem;
            line-height: 1.1;
        }
        .hero-copy p {
            margin: 0;
            color: #7a4b37;
            font-size: 1rem;
        }
        .panda-wrap {
            display: flex;
            justify-content: flex-end;
            align-items: center;
            flex: 0 0 190px;
            min-width: 190px;
        }
        .panda-wrap img {
            display: block;
            max-width: 170px;
            width: 170px;
            height: auto;
            object-fit: contain;
            transition: transform 0.18s ease, filter 0.18s ease;
            filter: drop-shadow(0 10px 16px rgba(128, 66, 38, 0.10));
        }
        .panda-wrap a {
            display: inline-block;
        }
        .panda-wrap a:hover img {
            transform: translateY(-2px) scale(1.02);
            filter: drop-shadow(0 14px 20px rgba(128, 66, 38, 0.16));
        }
        .sheet-card {
            padding: 10px 12px;
            margin-bottom: 8px;
            border-radius: 12px;
            border: 1px solid rgba(170, 90, 60, 0.08);
            background: rgba(255, 255, 255, 0.66);
        }
        .settings-section {
            margin: 16px 0 10px 0;
        }
        .settings-section h4 {
            margin: 0 0 8px 0;
            color: #5b2c1f;
            font-size: 1rem;
        }
        .settings-divider {
            height: 1px;
            margin: 6px 0 12px 0;
            background: linear-gradient(90deg, rgba(200, 95, 59, 0.22), rgba(200, 95, 59, 0.04));
        }
        .sheet-title {
            margin: 0 0 4px 0;
            color: #5b2c1f;
            font-weight: 700;
        }
        .sheet-meta {
            margin: 0;
            color: #8a5d48;
            font-size: 0.92rem;
        }
        .status-list {
            display: grid;
            gap: 8px;
            padding-bottom: 18px;
        }
        .status-item {
            display: grid;
            grid-template-columns: 108px minmax(0, 1fr);
            align-items: center;
            column-gap: 12px;
            color: #40261d;
            line-height: 1.45;
        }
        .status-badge {
            display: inline-block;
            width: 96px;
            padding: 3px 10px;
            border-radius: 10px;
            border: 1px solid rgba(132, 82, 58, 0.14);
            background: #fff6ee;
            color: #7a3f2d;
            font-size: 0.9rem;
            font-weight: 700;
            text-align: center;
            font-family: "Consolas", "Courier New", monospace;
        }
        .status-note {
            color: #5f3b2e;
        }
        .diagnostic-shell {
            padding: 6px 2px 2px 2px;
        }
        .diagnostic-meta {
            margin: 10px 0 8px 0;
            color: #7a4b37;
            font-size: 0.95rem;
        }
        .columns-panel {
            margin-top: 6px;
            max-height: 160px;
            overflow-y: scroll;
            padding: 10px;
            border-radius: 12px;
            border: 1px solid rgba(170, 90, 60, 0.10);
            background: rgba(255, 252, 248, 0.88);
            scrollbar-color: rgba(200, 95, 59, 0.7) rgba(255, 241, 227, 0.9);
            scrollbar-width: thin;
            padding-right: 14px;
        }
        .columns-list {
            display: grid;
            gap: 6px;
        }
        .column-line,
        .columns-empty {
            color: #8a5d48;
            font-size: 0.92rem;
        }
        .column-line strong {
            color: #5b2c1f;
            display: inline-block;
            min-width: 30px;
        }
        .columns-panel::-webkit-scrollbar {
            width: 10px;
        }
        .columns-panel::-webkit-scrollbar-track {
            background: rgba(255, 241, 227, 0.9);
            border-radius: 999px;
        }
        .columns-panel::-webkit-scrollbar-thumb {
            background: rgba(200, 95, 59, 0.7);
            border-radius: 999px;
            border: 2px solid rgba(255, 241, 227, 0.9);
        }
        .summary-table {
            margin-top: 12px;
            border: 1px solid rgba(132, 82, 58, 0.14);
            border-radius: 16px;
            overflow: hidden;
            background: #fffaf5;
        }
        .summary-row {
            display: grid;
            grid-template-columns: minmax(0, 1fr) 140px;
            align-items: center;
        }
        .summary-row + .summary-row {
            border-top: 1px solid rgba(132, 82, 58, 0.08);
        }
        .summary-row.header {
            background: #f7ece1;
            color: #5b2c1f;
            font-weight: 700;
        }
        .summary-cell {
            padding: 12px 16px;
            color: #40261d;
        }
        .summary-cell.count {
            text-align: right;
            font-variant-numeric: tabular-nums;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
        <div class="hero-card">
            <div class="hero-copy">
                <h1>Excel Link Cleaner</h1>
                <p>Локальный интерфейс для проверки ссылок в Excel.</p>
            </div>
            {panda_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def save_uploaded_file(uploaded_file) -> str:
    suffix = Path(uploaded_file.name).suffix or ".xlsx"

    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        return tmp.name


def get_uploaded_file_cache() -> dict:
    if "uploaded_file_cache" not in st.session_state:
        st.session_state.uploaded_file_cache = {
            "signature": None,
            "path": None,
        }

    return st.session_state.uploaded_file_cache


def get_cached_uploaded_file_path(uploaded_file) -> str:
    cache = get_uploaded_file_cache()
    signature = (uploaded_file.name, uploaded_file.size)

    if cache["signature"] == signature and cache["path"]:
        cached_path = Path(cache["path"])
        if cached_path.exists():
            return str(cached_path)

    old_path = cache.get("path")
    if old_path:
        old_file = Path(old_path)
        if old_file.exists():
            old_file.unlink()

    new_path = save_uploaded_file(uploaded_file)
    cache["signature"] = signature
    cache["path"] = new_path
    return new_path


def get_workbook_info(file_path: str) -> dict:
    workbook = openpyxl.load_workbook(file_path, read_only=True)

    try:
        sheet_names = workbook.sheetnames
        columns_by_sheet = {}

        for sheet_name in sheet_names:
            sheet = workbook[sheet_name]
            headers = []

            for cell in sheet[1]:
                if cell.value is None:
                    headers.append("")
                else:
                    headers.append(str(cell.value).strip())

            columns_by_sheet[sheet_name] = headers

        return {
            "sheet_names": sheet_names,
            "columns_by_sheet": columns_by_sheet,
        }
    finally:
        workbook.close()


def read_blacklist_text(file_path: str) -> str:
    path = Path(file_path)
    if not path.exists():
        return ""

    return path.read_text(encoding="utf-8")


def write_blacklist_text(file_path: str, text: str) -> None:
    Path(file_path).write_text(text, encoding="utf-8")


def make_download_bytes(file_path: str) -> bytes:
    return Path(file_path).read_bytes()


def get_checker_state() -> dict:
    if "checker_state" not in st.session_state:
        st.session_state.checker_state = {
            "thread": None,
            "cancel_event": None,
            "job": None,
        }

    return st.session_state.checker_state


def start_background_check(job: dict, cancel_event: threading.Event) -> threading.Thread:
    def worker() -> None:
        try:
            (
                output_path,
                summary,
                processed_sheets,
                skipped_sheets,
                empty_sheets,
                remembered_column_name,
                remembered_source,
            ) = process_workbook(
                input_file=job["input_file"],
                sheet_name=job["selected_sheet"],
                process_all_sheets=job["all_sheets"],
                column_name=job["column_name"],
                column_index=job["column_index"],
                preferred_column_name=job["preferred_column_name"],
                known_names=job["known_names"],
                blacklist_file=job["blacklist_path"],
                output_file=None,
                timeout=job["timeout"],
                max_workers=job["workers"],
                add_details=job["details"],
                progress_callback=lambda sheet_title, done, total: job["progress"].update(
                    {
                        "sheet_title": sheet_title,
                        "done": done,
                        "total": total,
                    }
                ),
                cancel_event=cancel_event,
            )

            if processed_sheets > 0 and not cancel_event.is_set():
                update_known_names_after_success(
                    known_names=job["known_names"],
                    known_names_path=job["known_names_path"],
                    resolved_input_file=job["input_file"],
                    resolved_column_name=remembered_column_name,
                    explicit_input_file=job["uploaded_name"],
                    explicit_column_name=job["column_name"],
                    preferred_column_name=job["preferred_column_name"],
                    resolution_source=remembered_source,
                )

            job["status"] = "cancelled" if cancel_event.is_set() else "completed"
            job["result"] = {
                "output_path": output_path,
                "summary": summary,
                "processed_sheets": processed_sheets,
                "skipped_sheets": skipped_sheets,
                "empty_sheets": empty_sheets,
                "output_name": f"{Path(job['uploaded_name']).stem}_checked.xlsx",
            }
        except Exception as error:
            job["status"] = "error"
            job["error"] = str(error)

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    return thread


def render_workbook_overview(sheet_names: list[str], columns_by_sheet: dict) -> None:
    with st.expander("Листы и колонки", expanded=False):
        selected_sheet = st.selectbox(
            "Лист для просмотра",
            options=sheet_names,
            key="diagnostic_sheet_select",
        )

        headers = [col for col in columns_by_sheet[selected_sheet] if col]

        st.markdown('<div class="diagnostic-shell">', unsafe_allow_html=True)
        st.markdown(
            f"""
            <div class="diagnostic-meta"><strong>{selected_sheet}</strong> · колонок с заголовками: {len(headers)}</div>
            """,
            unsafe_allow_html=True,
        )

        if headers:
            column_lines = "".join(
                f"<div class='column-line'><strong>{index}.</strong> {header}</div>"
                for index, header in enumerate(headers, start=1)
            )
            st.markdown(
                f"""
                <div class="columns-panel">
                    <div class="columns-list">{column_lines}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                '<div class="columns-panel"><div class="columns-empty">На выбранном листе нет заголовков.</div></div>',
                unsafe_allow_html=True,
            )

        st.markdown("</div>", unsafe_allow_html=True)


def show_summary(summary: dict, processed_sheets: int, skipped_sheets: int, empty_sheets: int) -> None:
    st.subheader("Summary")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("processed sheets", processed_sheets)
    col2.metric("skipped sheets", skipped_sheets)
    col3.metric("empty sheets", empty_sheets)
    col4.metric("total rows", sum(summary.values()))

    ordered_statuses = [
        "ok",
        "redirect",
        "blacklist",
        "empty",
        "invalid",
        "blocked",
        "failed",
    ]

    rows_html = "".join(
        f"""
        <div class="summary-row">
            <div class="summary-cell">{status}</div>
            <div class="summary-cell count">{summary.get(status, 0)}</div>
        </div>
        """
        for status in ordered_statuses
    )

    st.markdown(
        f"""
        <div class="summary-table">
            <div class="summary-row header">
                <div class="summary-cell">status</div>
                <div class="summary-cell count">count</div>
            </div>
            {rows_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def show_status_guide() -> None:
    with st.expander("Как читать статусы", expanded=False):
        st.markdown(
            """
            <div class="status-list">
                <div class="status-item">
                    <span class="status-badge">ok</span>
                    <span class="status-note">ссылка открылась нормально.</span>
                </div>
                <div class="status-item">
                    <span class="status-badge">redirect</span>
                    <span class="status-note">ссылка жива, но ведет через редирект.</span>
                </div>
                <div class="status-item">
                    <span class="status-badge">blocked</span>
                    <span class="status-note">сайт явно ограничил доступ, обычно это 403, 429 или challenge.</span>
                </div>
                <div class="status-item">
                    <span class="status-badge">failed</span>
                    <span class="status-note">страницу не удалось надежно проверить или материал выглядит удаленным.</span>
                </div>
                <div class="status-item">
                    <span class="status-badge">blacklist</span>
                    <span class="status-note">домен в стоп-листе.</span>
                </div>
                <div class="status-item">
                    <span class="status-badge">invalid</span>
                    <span class="status-note">ссылка пустая или некорректная по формату.</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


render_app_chrome()

known_names_path = "known_names.json"
blacklist_path = "blacklist.txt"

known_names = load_known_names(known_names_path)

uploaded_file = st.file_uploader(
    "Загрузи Excel-файл",
    type=["xlsx"],
)

if uploaded_file:
    temp_input_file = get_cached_uploaded_file_path(uploaded_file)
    file_info = get_workbook_info(temp_input_file)
    checker_state = get_checker_state()

    sheet_names = file_info["sheet_names"]
    columns_by_sheet = file_info["columns_by_sheet"]

    st.success(f"Файл загружен: {uploaded_file.name}")
    render_workbook_overview(sheet_names, columns_by_sheet)
    show_status_guide()

    st.subheader("Настройки")
    st.markdown('<div class="settings-section"><h4>Листы</h4><div class="settings-divider"></div></div>', unsafe_allow_html=True)
    all_sheets_choice = st.radio(
        "Обработать все листы?",
        options=["Да", "Нет"],
        index=0,
        horizontal=True,
        key="all_sheets_radio",
    )
    all_sheets = all_sheets_choice == "Да"

    selected_sheet = None
    if not all_sheets:
        selected_sheet = st.selectbox("Выбери лист", sheet_names)

    st.markdown('<div class="settings-section"><h4>Колонка со ссылками</h4><div class="settings-divider"></div></div>', unsafe_allow_html=True)
    column_mode = st.radio(
        "Как выбрать колонку?",
        options=["Авто", "По названию", "По номеру"],
        horizontal=True,
    )

    column_name = None
    column_index = None

    if column_mode == "По названию":
        if not all_sheets and selected_sheet:
            current_headers = [col for col in columns_by_sheet[selected_sheet] if col]
        else:
            all_headers = []
            for headers in columns_by_sheet.values():
                all_headers.extend([col for col in headers if col])
            current_headers = sorted(set(all_headers))

        if current_headers:
            selected_header = st.selectbox(
                "Выбери колонку",
                options=[""] + current_headers,
                index=0,
            )
            column_name = selected_header or None
        else:
            column_name = st.text_input("Название колонки")

    elif column_mode == "По номеру":
        column_index = st.number_input(
            "Номер колонки",
            min_value=1,
            step=1,
            value=1,
        )

    st.markdown('<div class="settings-section"><h4>Проверка и результат</h4><div class="settings-divider"></div></div>', unsafe_allow_html=True)
    details_choice = st.radio(
        "Добавить tech-колонку?",
        options=["Да", "Нет"],
        index=1,
        horizontal=True,
        key="details_radio",
    )
    details = details_choice == "Да"

    timeout = st.number_input(
        "Таймаут, сек",
        min_value=1,
        max_value=60,
        value=8,
        step=1,
    )

    workers = st.number_input(
        "Количество потоков",
        min_value=1,
        max_value=64,
        value=12,
        step=1,
    )

    with st.expander("Blacklist domains", expanded=False):
        blacklist_text = st.text_area(
            "Blacklist domains",
            value=read_blacklist_text(blacklist_path),
            height=180,
            label_visibility="collapsed",
        )

        if st.button("Сохранить blacklist"):
            write_blacklist_text(blacklist_path, blacklist_text)
            st.success("blacklist.txt обновлен")

    job = checker_state.get("job")
    thread = checker_state.get("thread")
    is_running = bool(job) and job.get("status") == "running" and thread is not None and thread.is_alive()

    if is_running:
        progress = job["progress"]
        safe_total = max(progress.get("total", 0), 1)
        percent = min(progress.get("done", 0) / safe_total, 1.0)

        st.info(
            f"Идет проверка. Лист: {progress.get('sheet_title', 'подготовка')} | "
            f"уникальных ссылок: {progress.get('done', 0)}/{progress.get('total', 0)}"
        )
        st.progress(percent)

        if st.button("Остановить проверку", type="secondary"):
            checker_state["cancel_event"].set()
            st.warning("Останавливаю проверку. Уже начатые запросы закончатся по таймауту.")
            st.rerun()

        time.sleep(1)
        st.rerun()
    else:
        if st.button("Запустить проверку", type="primary"):
            preferred_column_name = None

            if column_mode == "Авто":
                preferred_column_name = detect_preferred_column_from_rules(
                    resolved_input_file=temp_input_file,
                    known_names=known_names,
                )

            job = {
                "status": "running",
                "input_file": temp_input_file,
                "uploaded_name": uploaded_file.name,
                "selected_sheet": selected_sheet,
                "all_sheets": all_sheets,
                "column_name": column_name,
                "column_index": column_index,
                "preferred_column_name": preferred_column_name,
                "known_names": known_names,
                "known_names_path": known_names_path,
                "blacklist_path": blacklist_path,
                "timeout": int(timeout),
                "workers": int(workers),
                "details": details,
                "progress": {
                    "sheet_title": "подготовка",
                    "done": 0,
                    "total": 0,
                },
                "result": None,
                "error": None,
            }

            cancel_event = threading.Event()
            checker_state["job"] = job
            checker_state["cancel_event"] = cancel_event
            checker_state["thread"] = start_background_check(job, cancel_event)
            st.rerun()

    if job and job.get("status") == "completed" and job.get("result"):
        result = job["result"]
        st.success("Проверка завершена")
        show_summary(
            result["summary"],
            result["processed_sheets"],
            result["skipped_sheets"],
            result["empty_sheets"],
        )
        st.download_button(
            label="Скачать результат",
            data=make_download_bytes(result["output_path"]),
            file_name=result["output_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if job and job.get("status") == "cancelled":
        st.warning("Проверка остановлена. Можно запустить заново с другими параметрами.")
        if job.get("result"):
            result = job["result"]
            show_summary(
                result["summary"],
                result["processed_sheets"],
                result["skipped_sheets"],
                result["empty_sheets"],
            )

    if job and job.get("status") == "error":
        st.error(f"Ошибка во время проверки: {job.get('error', 'неизвестная ошибка')}")
else:
    st.info("Сначала загрузи Excel-файл.")
