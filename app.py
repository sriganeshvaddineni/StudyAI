
import streamlit as st
import sqlite3
import re
import html
from pathlib import Path
from datetime import datetime
from io import BytesIO

from pypdf import PdfReader
from pptx import Presentation
from docx import Document

from ai_teacher import generate_lesson


# =====================================================
# CONFIG
# =====================================================

st.set_page_config(
    page_title="StudyAI",
    page_icon="📚",
    layout="wide"
)

DB_PATH = Path(__file__).parent / "studai.db"
MAX_MATERIAL_CHARS = 30000


# =====================================================
# DATABASE
# =====================================================

def db():
    return sqlite3.connect(DB_PATH)


def init_db():
    with db() as conn:
        cur = conn.cursor()

        cur.execute("""
            CREATE TABLE IF NOT EXISTS subjects(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS units(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                subject_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                UNIQUE(subject_id,name)
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS topics(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                unit_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                UNIQUE(unit_id,name)
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS completed_topics(
                topic_id INTEGER PRIMARY KEY,
                completed INTEGER NOT NULL DEFAULT 1
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS unit_materials(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                unit_id INTEGER NOT NULL,
                filename TEXT NOT NULL,
                file_type TEXT NOT NULL,
                extracted_text TEXT NOT NULL,
                uploaded_at TEXT NOT NULL
            )
        """)


init_db()


# =====================================================
# BLUE MESSAGE
# =====================================================

def blue_message(message):
    st.html(f"""
    <div style="
        background:#ffffff;
        color:#1e40af;
        border:1px solid #bfdbfe;
        border-radius:10px;
        padding:14px 16px;
        margin:10px 0;
        font-size:15px;
        font-weight:500;
        overflow-wrap:anywhere;
    ">
        {html.escape(str(message))}
    </div>
    """)


# =====================================================
# MATERIAL EXTRACTION
# =====================================================

def extract_pdf(file_bytes):
    reader = PdfReader(BytesIO(file_bytes))
    pages = []

    for page in reader.pages:
        pages.append(page.extract_text() or "")

    return "\n".join(pages).strip()


def extract_pptx(file_bytes):
    presentation = Presentation(BytesIO(file_bytes))
    content = []

    for slide_number, slide in enumerate(
        presentation.slides, start=1
    ):
        content.append(f"\nSlide {slide_number}")

        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                content.append(shape.text.strip())

            if shape.has_table:
                for row in shape.table.rows:
                    content.append(
                        " | ".join(
                            cell.text for cell in row.cells
                        )
                    )

    return "\n".join(content).strip()


def extract_docx(file_bytes):
    document = Document(BytesIO(file_bytes))
    content = []

    for paragraph in document.paragraphs:
        if paragraph.text.strip():
            content.append(paragraph.text.strip())

    for table in document.tables:
        for row in table.rows:
            content.append(
                " | ".join(cell.text for cell in row.cells)
            )

    return "\n".join(content).strip()


def extract_file_text(uploaded_file):
    file_bytes = uploaded_file.getvalue()
    extension = Path(uploaded_file.name).suffix.lower()

    if extension == ".pdf":
        return extract_pdf(file_bytes)

    if extension == ".pptx":
        return extract_pptx(file_bytes)

    if extension == ".docx":
        return extract_docx(file_bytes)

    raise ValueError("Unsupported file format.")


# =====================================================
# MATERIAL DATABASE FUNCTIONS
# =====================================================

def save_material(uid, filename, file_type, extracted_text):
    with db() as conn:
        conn.execute("""
            INSERT INTO unit_materials(
                unit_id,
                filename,
                file_type,
                extracted_text,
                uploaded_at
            )
            VALUES (?, ?, ?, ?, ?)
        """, (
            uid,
            filename,
            file_type,
            extracted_text,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))


def get_materials(uid):
    with db() as conn:
        return conn.execute("""
            SELECT id, filename, file_type,
                   extracted_text, uploaded_at
            FROM unit_materials
            WHERE unit_id=?
            ORDER BY id DESC
        """, (uid,)).fetchall()


def get_unit_material_text(uid):
    materials = get_materials(uid)

    return "\n\n".join(
        f"FILE: {filename}\n{text}"
        for _, filename, _, text, _ in materials
        if text.strip()
    )


def delete_material(material_id):
    with db() as conn:
        conn.execute(
            "DELETE FROM unit_materials WHERE id=?",
            (material_id,)
        )


# =====================================================
# SUBJECT FUNCTIONS
# =====================================================

def get_subjects():
    with db() as conn:
        return conn.execute(
            "SELECT id,name FROM subjects ORDER BY id"
        ).fetchall()


def get_subject(sid):
    with db() as conn:
        return conn.execute(
            "SELECT id,name FROM subjects WHERE id=?",
            (sid,)
        ).fetchone()


def add_subject(name):
    try:
        with db() as conn:
            conn.execute(
                "INSERT INTO subjects(name) VALUES(?)",
                (name,)
            )
        return True
    except sqlite3.IntegrityError:
        return False


def delete_subject(sid):
    with db() as conn:
        units = conn.execute(
            "SELECT id FROM units WHERE subject_id=?",
            (sid,)
        ).fetchall()

        for (uid,) in units:
            topics = conn.execute(
                "SELECT id FROM topics WHERE unit_id=?",
                (uid,)
            ).fetchall()

            for (tid,) in topics:
                conn.execute(
                    "DELETE FROM completed_topics WHERE topic_id=?",
                    (tid,)
                )

            conn.execute(
                "DELETE FROM topics WHERE unit_id=?",
                (uid,)
            )

            conn.execute(
                "DELETE FROM unit_materials WHERE unit_id=?",
                (uid,)
            )

        conn.execute(
            "DELETE FROM units WHERE subject_id=?",
            (sid,)
        )

        conn.execute(
            "DELETE FROM subjects WHERE id=?",
            (sid,)
        )


# =====================================================
# UNIT FUNCTIONS
# =====================================================

def get_units(sid):
    with db() as conn:
        return conn.execute(
            "SELECT id,name FROM units WHERE subject_id=? ORDER BY id",
            (sid,)
        ).fetchall()


def get_unit(uid):
    with db() as conn:
        return conn.execute(
            "SELECT id,name,subject_id FROM units WHERE id=?",
            (uid,)
        ).fetchone()


def add_unit(sid, name):
    try:
        with db() as conn:
            conn.execute(
                "INSERT INTO units(subject_id,name) VALUES(?,?)",
                (sid, name)
            )
        return True
    except sqlite3.IntegrityError:
        return False


def delete_unit(uid):
    with db() as conn:
        topics = conn.execute(
            "SELECT id FROM topics WHERE unit_id=?",
            (uid,)
        ).fetchall()

        for (tid,) in topics:
            conn.execute(
                "DELETE FROM completed_topics WHERE topic_id=?",
                (tid,)
            )

        conn.execute(
            "DELETE FROM topics WHERE unit_id=?",
            (uid,)
        )

        conn.execute(
            "DELETE FROM unit_materials WHERE unit_id=?",
            (uid,)
        )

        conn.execute(
            "DELETE FROM units WHERE id=?",
            (uid,)
        )


# =====================================================
# TOPIC FUNCTIONS
# =====================================================

def get_topics(uid):
    with db() as conn:
        return conn.execute(
            "SELECT id,name FROM topics WHERE unit_id=? ORDER BY id",
            (uid,)
        ).fetchall()


def get_topic(tid):
    with db() as conn:
        return conn.execute(
            "SELECT id,name,unit_id FROM topics WHERE id=?",
            (tid,)
        ).fetchone()


def add_topics_bulk(uid, text):
    names = re.split(r"[,;\n]+", text)
    added = 0

    with db() as conn:
        for name in names:
            name = name.strip()

            if name:
                try:
                    conn.execute(
                        "INSERT INTO topics(unit_id,name) VALUES(?,?)",
                        (uid, name)
                    )
                    added += 1
                except sqlite3.IntegrityError:
                    pass

    return added


def delete_topic(tid):
    with db() as conn:
        conn.execute(
            "DELETE FROM completed_topics WHERE topic_id=?",
            (tid,)
        )

        conn.execute(
            "DELETE FROM topics WHERE id=?",
            (tid,)
        )


def split_existing_topics(uid):
    split_count = 0
    new_count = 0

    with db() as conn:
        rows = conn.execute(
            "SELECT id,name FROM topics WHERE unit_id=?",
            (uid,)
        ).fetchall()

        for tid, name in rows:
            if "," not in name and ";" not in name:
                continue

            parts = [
                p.strip()
                for p in re.split(r"[,;]+", name)
                if p.strip()
            ]

            if len(parts) <= 1:
                continue

            conn.execute(
                "DELETE FROM topics WHERE id=?",
                (tid,)
            )

            conn.execute(
                "DELETE FROM completed_topics WHERE topic_id=?",
                (tid,)
            )

            split_count += 1

            for part in parts:
                try:
                    conn.execute(
                        "INSERT INTO topics(unit_id,name) VALUES(?,?)",
                        (uid, part)
                    )
                    new_count += 1
                except sqlite3.IntegrityError:
                    pass

    return split_count, new_count


# =====================================================
# COMPLETION FUNCTIONS
# =====================================================

def is_completed(tid):
    with db() as conn:
        result = conn.execute(
            "SELECT completed FROM completed_topics WHERE topic_id=?",
            (tid,)
        ).fetchone()

    return bool(result and result[0])


def set_completed(tid, completed):
    with db() as conn:
        if completed:
            conn.execute("""
                INSERT OR REPLACE INTO completed_topics(
                    topic_id, completed
                )
                VALUES (?,1)
            """, (tid,))
        else:
            conn.execute(
                "DELETE FROM completed_topics WHERE topic_id=?",
                (tid,)
            )


# =====================================================
# NAVIGATION
# =====================================================

defaults = {
    "page": "home",
    "sid": None,
    "uid": None,
    "tid": None
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


def go(page, sid=None, uid=None, tid=None):
    st.session_state.page = page

    if sid is not None:
        st.session_state.sid = sid

    if uid is not None:
        st.session_state.uid = uid

    if tid is not None:
        st.session_state.tid = tid

    st.rerun()


# =====================================================
# BLUE-WHITE DESIGN
# =====================================================

st.markdown("""
<style>
.stApp,
[data-testid="stAppViewContainer"] {
    background:#f4f7fc !important;
}

.block-container {
    max-width:1250px;
    padding-top:2rem;
    padding-bottom:3rem;
}

h1,h2,h3,h4 {
    color:#172554 !important;
}

p,label,li {
    color:#334155 !important;
}

[data-testid="stSidebar"] {
    background:#111827 !important;
}

[data-testid="stSidebar"] * {
    color:#ffffff !important;
}

.stButton button,
.stFormSubmitButton button {
    background:#2563eb !important;
    color:#ffffff !important;
    border:1px solid #2563eb !important;
    border-radius:10px !important;
    min-height:42px;
}

.stButton button p,
.stFormSubmitButton button p {
    color:#ffffff !important;
}

.stButton button:hover,
.stFormSubmitButton button:hover {
    background:#1d4ed8 !important;
}

[data-testid="stTextInput"] input,
[data-testid="stTextArea"] textarea {
    background:#ffffff !important;
    color:#1e293b !important;
    border:1px solid #cbd5e1 !important;
    caret-color:#2563eb !important;
}

[data-testid="stTextInput"] input::placeholder,
[data-testid="stTextArea"] textarea::placeholder {
    color:#64748b !important;
    opacity:1 !important;
}

[data-testid="stForm"] {
    background:#ffffff !important;
    border:1px solid #dbe3ef !important;
    border-radius:16px !important;
    padding:20px !important;
}

[data-testid="stVerticalBlockBorderWrapper"] {
    background:#ffffff !important;
    border:1px solid #dbe3ef !important;
    border-radius:16px !important;
}

[data-testid="stMetric"] {
    background:#ffffff !important;
    padding:18px !important;
    border-radius:14px !important;
    border:1px solid #dbe3ef !important;
}

[data-testid="stMetricValue"] {
    color:#172554 !important;
}

[data-testid="stMetricLabel"] {
    color:#475569 !important;
}

[data-testid="stExpander"] {
    background:#ffffff !important;
    border:1px solid #dbe3ef !important;
    border-radius:12px !important;
}

[data-testid="stExpander"] details {
    background:#ffffff !important;
    border-radius:12px !important;
}

[data-testid="stExpander"] summary {
    background:#eff6ff !important;
    color:#1e3a8a !important;
    border-radius:10px !important;
}

[data-testid="stExpander"] summary:hover {
    background:#dbeafe !important;
}

[data-testid="stExpander"] summary p,
[data-testid="stExpander"] summary span {
    color:#1e3a8a !important;
}

[data-testid="stExpander"] [data-testid="stVerticalBlock"] {
    background:#ffffff !important;
}

hr {
    border-color:#cbd5e1 !important;
}

input[type="checkbox"] {
    accent-color:#2563eb !important;
}

[data-testid="stCheckbox"] input {
    accent-color:#2563eb !important;
}

[data-testid="stCheckbox"] svg {
    color:#2563eb !important;
}

[data-testid="stCheckbox"] input[type="checkbox"] {
    appearance:none !important;
    -webkit-appearance:none !important;
    width:18px !important;
    height:18px !important;
    background-color:#ffffff !important;
    border:1px solid #94a3b8 !important;
    border-radius:4px !important;
    cursor:pointer !important;
    display:inline-grid !important;
    place-content:center !important;
}

[data-testid="stCheckbox"] input[type="checkbox"]:checked {
    background-color:#2563eb !important;
    border-color:#2563eb !important;
}

[data-testid="stCheckbox"] input[type="checkbox"]:checked::after {
    content:"✓";
    color:white;
    font-size:13px;
    font-weight:bold;
}

[data-testid="stMarkdownContainer"] table {
    width:100% !important;
    border-collapse:collapse !important;
    background-color:#ffffff !important;
    color:#1e293b !important;
}

[data-testid="stMarkdownContainer"] table th {
    background-color:#dbeafe !important;
    color:#1e3a8a !important;
    -webkit-text-fill-color:#1e3a8a !important;
    font-weight:700 !important;
    padding:12px !important;
    border:1px solid #cbd5e1 !important;
}

[data-testid="stMarkdownContainer"] table td {
    background-color:#ffffff !important;
    color:#334155 !important;
    -webkit-text-fill-color:#334155 !important;
    padding:12px !important;
    border:1px solid #cbd5e1 !important;
}

[data-testid="stMarkdownContainer"] table td p,
[data-testid="stMarkdownContainer"] table th p {
    color:inherit !important;
    -webkit-text-fill-color:inherit !important;
}

[data-testid="stFileUploader"] {
    background-color:#ffffff !important;
    border:2px dashed #2563eb !important;
    border-radius:12px !important;
    padding:15px !important;
}

[data-testid="stFileUploader"] section {
    background-color:#ffffff !important;
}

[data-testid="stFileUploader"] div {
    color:#1e293b !important;
}

[data-testid="stFileUploader"] button {
    background-color:#2563eb !important;
    color:#ffffff !important;
    border:none !important;
    border-radius:8px !important;
}

[data-testid="stFileUploader"] small {
    color:#64748b !important;
}

[data-testid="stFileUploaderFile"] {
    background-color:#ffffff !important;
    color:#1e293b !important;
    border:1px solid #bfdbfe !important;
    border-radius:10px !important;
}

[data-testid="stFileUploaderFile"] * {
    color:#1e293b !important;
    -webkit-text-fill-color:#1e293b !important;
}

[data-testid="stFileUploaderFile"] button {
    background-color:#ffffff !important;
    color:#1e293b !important;
}

[data-testid="stFileUploaderFile"] svg {
    fill:#2563eb !important;
}

[data-testid="stFileUploader"] label,
[data-testid="stFileUploader"] p,
[data-testid="stFileUploader"] span,
[data-testid="stFileUploader"] small,
[data-testid="stFileUploader"] div {
    color:#334155 !important;
    -webkit-text-fill-color:#334155 !important;
    opacity:1 !important;
}

[data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"] {
    color:#334155 !important;
    -webkit-text-fill-color:#334155 !important;
    opacity:1 !important;
}

[data-testid="stFileUploader"] button {
    background:#2563eb !important;
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
}

[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] {
    display:none !important;
}
</style>
""", unsafe_allow_html=True)


# =====================================================
# SAVED MATERIAL CARD
# =====================================================

def render_material_card(
    filename,
    file_type,
    uploaded_at,
    extracted_text
):
    safe_filename = html.escape(filename)
    safe_file_type = html.escape(file_type.upper())
    safe_uploaded_at = html.escape(uploaded_at)

    st.html(f"""
    <div style="
        background:#ffffff;
        color:#1e293b;
        border:1px solid #bfdbfe;
        border-radius:12px;
        padding:16px;
        margin-bottom:12px;
        width:100%;
        box-sizing:border-box;
    ">
        <div style="
            font-size:16px;
            font-weight:600;
            color:#1e40af;
            overflow-wrap:anywhere;
        ">
            📄 {safe_filename}
        </div>

        <div style="
            font-size:13px;
            color:#64748b;
            margin-top:6px;
        ">
            {safe_file_type}
            &nbsp; • &nbsp;
            Uploaded: {safe_uploaded_at}
        </div>

        <div style="
            font-size:13px;
            color:#475569;
            margin-top:6px;
        ">
            {len(extracted_text)} characters extracted
        </div>
    </div>
    """)


# =====================================================
# SIDEBAR
# =====================================================

with st.sidebar:
    st.markdown("# 📚 StudyAI")
    st.caption("Your personal AI learning space")
    st.divider()

    if st.button("🏠 Home", use_container_width=True):
        go("home")

    if st.button("📚 My Subjects", use_container_width=True):
        go("subjects")

    st.divider()
    st.caption("Learn smarter. Score better.")


# =====================================================
# SUBJECT CARDS
# =====================================================

def render_subject_cards(subjects):
    if not subjects:
        blue_message("No subjects yet. Add your first subject.")
        return

    cols = st.columns(3)

    for i, (sid, name) in enumerate(subjects):
        units = get_units(sid)

        topic_count = sum(
            len(get_topics(uid)) for uid, _ in units
        )

        with cols[i % 3]:
            with st.container(border=True):
                st.markdown("### 📘 " + name)

                st.caption(
                    f"{len(units)} Units • {topic_count} Topics"
                )

                if st.button(
                    "Open Subject →",
                    key=f"open_subject_{sid}",
                    use_container_width=True
                ):
                    go("subject", sid=sid)


# =====================================================
# HOME PAGE
# =====================================================

if st.session_state.page == "home":
    subjects = get_subjects()

    st.html("""
    <div class="hero">
        <h1>Welcome to StudyAI 👋</h1>
        <p>Your personal space to learn, understand and prepare for exams.</p>
    </div>
    """)

    total_units = sum(
        len(get_units(sid)) for sid, _ in subjects
    )

    total_topics = sum(
        len(get_topics(uid))
        for sid, _ in subjects
        for uid, _ in get_units(sid)
    )

    a, b, c = st.columns(3)

    a.metric("📚 Subjects", len(subjects))
    b.metric("📖 Units", total_units)
    c.metric("📝 Topics", total_topics)

    st.divider()
    st.subheader("📚 My Subjects")

    render_subject_cards(subjects)

    st.divider()
    st.subheader("➕ Add Subject")

    with st.form("home_add_subject"):
        name = st.text_input("Subject name")
        submit = st.form_submit_button("Add Subject")

        if submit:
            if name.strip():
                if add_subject(name.strip()):
                    blue_message("✅ Subject added!")
                    st.rerun()
                else:
                    blue_message("⚠️ Subject already exists.")
            else:
                blue_message("⚠️ Enter a subject name.")


# =====================================================
# MY SUBJECTS PAGE
# =====================================================

elif st.session_state.page == "subjects":
    st.title("📚 My Subjects")

    render_subject_cards(get_subjects())

    st.divider()
    st.subheader("➕ Add Subject")

    with st.form("subjects_add_form"):
        name = st.text_input("Subject name")
        submit = st.form_submit_button("Add Subject")

        if submit:
            if name.strip():
                if add_subject(name.strip()):
                    blue_message("✅ Subject added!")
                    st.rerun()
                else:
                    blue_message("⚠️ Subject already exists.")
            else:
                blue_message("⚠️ Enter a subject name.")


# =====================================================
# SUBJECT PAGE
# =====================================================

elif st.session_state.page == "subject":
    sid = st.session_state.sid
    subject = get_subject(sid) if sid is not None else None

    if subject:
        if st.button("← My Subjects"):
            go("subjects")

        st.title("📘 " + subject[1])

        units = get_units(sid)

        st.metric("Total Units", len(units))
        st.divider()
        st.subheader("📖 Your Units")

        if not units:
            blue_message("No units yet. Add your first unit below.")

        cols = st.columns(2)

        for i, (uid, name) in enumerate(units):
            with cols[i % 2]:
                with st.container(border=True):
                    st.markdown("### 📖 " + name)
                    st.caption(f"{len(get_topics(uid))} Topics")

                    if st.button(
                        "Open Unit →",
                        key=f"open_unit_{uid}",
                        use_container_width=True
                    ):
                        go("unit", sid=sid, uid=uid)

        st.divider()
        st.subheader("➕ Add Unit")

        with st.form("add_unit_form"):
            name = st.text_input(
                "Unit name",
                placeholder="Example: Unit 2 - Functions"
            )

            submit = st.form_submit_button("Add Unit")

            if submit:
                if name.strip():
                    if add_unit(sid, name.strip()):
                        blue_message("✅ Unit added!")
                        st.rerun()
                    else:
                        blue_message("⚠️ Unit already exists.")
                else:
                    blue_message("⚠️ Enter a unit name.")

        st.divider()
        st.subheader("Delete Subject")

        confirm = st.toggle(
            f"Confirm delete: {subject[1]}",
            key=f"confirm_subject_{sid}"
        )

        if st.button("Delete Subject", key=f"delete_subject_{sid}"):
            if confirm:
                delete_subject(sid)
                go("subjects")
            else:
                blue_message("Please enable confirmation first.")

    else:
        blue_message("Subject not found.")


# =====================================================
# UNIT PAGE
# =====================================================

elif st.session_state.page == "unit":
    uid = st.session_state.uid
    unit = get_unit(uid) if uid is not None else None

    if unit:
        sid = unit[2]
        subject = get_subject(sid)

        if st.button("← Back to Subject"):
            go("subject", sid=sid)

        st.caption(subject[1] if subject else "Subject")
        st.title("📖 " + unit[1])

        # UPLOAD STUDY MATERIALS
        st.divider()
        st.subheader("📂 Unit Study Materials")

        st.write(
            "Upload your unit notes, PDFs, PowerPoint presentations "
            "or Word documents."
        )

        uploaded_files = st.file_uploader(
            "Choose PDF, PPTX or DOCX files",
            type=["pdf", "pptx", "docx"],
            accept_multiple_files=True,
            key=f"upload_materials_{uid}"
        )

        if uploaded_files:
            st.markdown("**Selected files:**")

            for uploaded_file in uploaded_files:
                safe_name = html.escape(uploaded_file.name)

                st.markdown(
                    f"""
                    <div style="
                        background:#ffffff;
                        color:#1e293b;
                        border:1px solid #bfdbfe;
                        border-radius:10px;
                        padding:12px;
                        margin:6px 0;
                        overflow-wrap:anywhere;
                    ">
                        📄 <b>{safe_name}</b>
                        <span style="color:#64748b;">
                            ({uploaded_file.size / 1024:.1f} KB)
                        </span>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            if st.button(
                "📤 Upload and Save Materials",
                key=f"save_materials_{uid}"
            ):
                saved_count = 0

                for uploaded_file in uploaded_files:
                    try:
                        extracted_text = extract_file_text(uploaded_file)

                        if not extracted_text:
                            blue_message(
                                f"No selectable text found in "
                                f"{uploaded_file.name}. "
                                "Scanned PDFs may require OCR."
                            )
                            continue

                        save_material(
                            uid,
                            uploaded_file.name,
                            Path(uploaded_file.name).suffix.lower(),
                            extracted_text
                        )

                        saved_count += 1

                    except Exception as error:
                        blue_message(
                            f"Could not process "
                            f"{uploaded_file.name}: {error}"
                        )

                if saved_count:
                    blue_message(
                        f"✅ {saved_count} file(s) uploaded and saved!"
                    )
                    st.rerun()

        # SHOW SAVED MATERIALS
        # FIXED: materials variable name
        materials = get_materials(uid)

        st.markdown(f"**Uploaded files: {len(materials)}**")

        if materials:
            for (
                material_id,
                filename,
                file_type,
                extracted_text,
                uploaded_at
            ) in materials:

                render_material_card(
                    filename,
                    file_type,
                    uploaded_at,
                    extracted_text
                )

                col1, col2 = st.columns([5, 1])

                with col1:
                    with st.expander("Preview extracted text"):
                        st.text(extracted_text[:5000])

                        if len(extracted_text) > 5000:
                            st.caption(
                                "Preview limited to first 5000 characters."
                            )

                with col2:
                    if st.button(
                        "🗑️ Delete",
                        key=f"delete_material_{material_id}"
                    ):
                        delete_material(material_id)
                        st.rerun()

        else:
            blue_message(
                "No study materials uploaded for this unit yet."
            )

        # FIX EXISTING COMBINED TOPICS
        st.divider()
        st.subheader("🔧 Fix Existing Combined Topics")

        st.write(
            "Comma-separated topics ni separate topics ga convert cheyyadaniki."
        )

        if st.button(
            "✂️ Fix Existing Topics",
            key=f"split_{uid}"
        ):
            split_count, new_count = split_existing_topics(uid)

            if split_count:
                blue_message(
                    f"{split_count} combined records split ayyayi. "
                    f"{new_count} new topics created."
                )
            else:
                blue_message("Combined topics emi dorakaledu.")

            st.rerun()

        # ADD TOPICS
        st.divider()
        st.subheader("➕ Add Topics")

        with st.form(f"bulk_topics_form_{uid}"):
            st.caption("Each line or comma = one topic")

            text = st.text_area(
                "Topic names",
                placeholder="""Introduction
Basic Concepts
Applications
Advantages and Disadvantages""",
                height=180
            )

            submit = st.form_submit_button("Save Topics")

            if submit:
                if text.strip():
                    added = add_topics_bulk(uid, text)
                    blue_message(f"✅ {added} topics added!")
                    st.rerun()
                else:
                    blue_message("Enter topic names.")

        # TOPIC LIST
        st.divider()

        topics = get_topics(uid)

        st.subheader(f"📝 Your Topics ({len(topics)})")

        if not topics:
            blue_message("No topics yet.")

        for i, (tid, name) in enumerate(topics):
            with st.container(border=True):
                col1, col2 = st.columns([5, 1])

                with col1:
                    st.markdown(f"**{i+1}. {name}**")

                with col2:
                    if st.button("Open", key=f"open_topic_{tid}"):
                        go("topic", sid=sid, uid=uid, tid=tid)

        # DELETE UNIT
        st.divider()
        st.subheader("Delete Unit")

        confirm = st.toggle(
            f"Confirm delete: {unit[1]} and all its topics",
            key=f"confirm_unit_{uid}"
        )

        if st.button("Delete Unit", key=f"delete_unit_{uid}"):
            if confirm:
                delete_unit(uid)
                go("subject", sid=sid)
            else:
                blue_message("Please enable confirmation first.")

    else:
        blue_message("Unit not found.")


# =====================================================
# TOPIC LEARNING PAGE
# =====================================================

elif st.session_state.page == "topic":
    uid = st.session_state.uid
    tid = st.session_state.tid

    unit = get_unit(uid) if uid is not None else None
    topic = get_topic(tid) if tid is not None else None

    if not unit or not topic:
        blue_message("Topic not found.")

        if st.button("Go to My Subjects"):
            go("subjects")

    else:
        sid = unit[2]

        topics = get_topics(uid)
        topic_ids = [t[0] for t in topics]

        if tid not in topic_ids:
            blue_message(
                "This topic doesn't belong to the selected unit."
            )

            if st.button("Back to Unit"):
                go("unit", sid=sid, uid=uid)

        else:
            current_index = topic_ids.index(tid)
            current_name = topic[1]
            completed = is_completed(tid)

            if st.button("← Back to Unit"):
                go("unit", sid=sid, uid=uid)

            st.title(unit[1])

            st.caption(
                f"Topic {current_index + 1} of {len(topics)}"
            )

            topic_col, lesson_col = st.columns(
                [1, 2.8],
                gap="large"
            )

            # LEFT: TOPIC NAVIGATION
            with topic_col:
                st.subheader("📚 All Topics")

                completed_count = sum(
                    1 for topic_id, _ in topics
                    if is_completed(topic_id)
                )

                st.progress(
                    completed_count / len(topics) if topics else 0
                )

                st.caption(
                    f"{completed_count}/{len(topics)} Completed"
                )

                for index, (topic_id, topic_name) in enumerate(
                    topics, start=1
                ):
                    is_current = topic_id == tid
                    is_done = is_completed(topic_id)

                    icon = (
                        "✅" if is_done
                        else ("🔵" if is_current else "○")
                    )

                    label = f"{icon} {index}. {topic_name}"

                    if is_current:
                        st.html(f"""
                        <div style="
                            background:#dbeafe;
                            border:1px solid #93c5fd;
                            border-radius:10px;
                            padding:10px;
                            margin-bottom:7px;
                            color:#1d4ed8;
                            font-weight:600;">
                            {html.escape(label)}
                        </div>
                        """)
                    else:
                        if st.button(
                            label,
                            key=f"topic_nav_{topic_id}",
                            use_container_width=True
                        ):
                            go(
                                "topic",
                                sid=sid,
                                uid=uid,
                                tid=topic_id
                            )

            # RIGHT: LEARNING CONTENT
            with lesson_col:
                st.html(f"""
                <div style="
                    background:white;
                    border:1px solid #dbe4f0;
                    border-radius:15px;
                    padding:22px;
                    margin-bottom:15px;">
                    <h2 style="color:#1e40af;">
                        {html.escape(current_name)}
                    </h2>
                    <p style="color:#64748b;">
                        Topic {current_index + 1} of {len(topics)}
                    </p>
                </div>
                """)

                tab1, tab2, tab3, tab4 = st.tabs([
                    "Telugu + English",
                    "Simple English",
                    "Exam Answer",
                    "Practice"
                ])

                subject_row = get_subject(sid)

                subject_name = (
                    subject_row[1]
                    if subject_row
                    else "Subject"
                )

                # GET UNIT MATERIAL TEXT
                material_text = get_unit_material_text(uid)

                def show_ai_lesson(mode, button_label):
                    lesson_key = f"ai_lesson_{tid}_{mode}"

                    if st.button(
                        button_label,
                        key=f"generate_{tid}_{mode}",
                        use_container_width=True
                    ):
                        try:
                            with st.spinner(
                                "🤖 StudyAI is preparing your lesson..."
                            ):
                                result = generate_lesson(
                                    current_name,
                                    subject_name,
                                    unit[1],
                                    mode,
                                    material_text=material_text[
                                        :MAX_MATERIAL_CHARS
                                    ]
                                )

                            st.session_state[lesson_key] = result

                            if (
                                result
                                and not result.startswith(
                                    "AI generation failed:"
                                )
                                and not result.startswith(
                                    "All Gemini models"
                                )
                            ):
                                blue_message(
                                    "✅ Your lesson is ready!"
                                )

                                if material_text:
                                    blue_message(
                                        "📚 Your uploaded unit materials "
                                        "were provided to the AI Teacher."
                                    )
                            else:
                                blue_message(f"⚠️ {result}")

                        except Exception as error:
                            blue_message(
                                f"⚠️ AI generation failed: {error}"
                            )

                    if lesson_key in st.session_state:
                        st.markdown("---")
                        st.markdown(
                            st.session_state[lesson_key]
                        )

                with tab1:
                    st.subheader("Telugu + English Explanation")

                    st.write(
                        "Learn this topic using simple Telugu and English."
                    )

                    show_ai_lesson(
                        "telugu",
                        "✨ Generate Telugu + English Explanation"
                    )

                with tab2:
                    st.subheader("Simple English Explanation")

                    st.write(
                        "Understand the topic in beginner-friendly English."
                    )

                    show_ai_lesson(
                        "simple",
                        "✨ Generate Simple English Explanation"
                    )

                with tab3:
                    st.subheader("Exam-Ready Answer")

                    st.write(
                        "Generate a structured answer for your college exams."
                    )

                    show_ai_lesson(
                        "exam",
                        "✨ Generate Exam Answer"
                    )

                with tab4:
                    st.subheader("Practice Questions")

                    st.write(
                        "Generate MCQs, short-answer questions "
                        "and long-answer questions."
                    )

                    show_ai_lesson(
                        "practice",
                        "✨ Generate Practice Questions"
                    )

                # MARK COMPLETED
                st.divider()

                mark_done = st.checkbox(
                    "Mark this topic as completed",
                    value=completed,
                    key=f"complete_topic_{tid}"
                )

                if mark_done != completed:
                    set_completed(tid, mark_done)
                    st.rerun()

                # PREVIOUS / NEXT NAVIGATION
                st.divider()

                previous_col, progress_col, next_col = st.columns(
                    [1, 2, 1]
                )

                with previous_col:
                    if current_index > 0:
                        if st.button(
                            "← Previous",
                            use_container_width=True
                        ):
                            go(
                                "topic",
                                sid=sid,
                                uid=uid,
                                tid=topics[current_index - 1][0]
                            )

                with progress_col:
                    st.html(f"""
                    <div style="
                        text-align:center;
                        padding-top:8px;">
                        Topic {current_index + 1} / {len(topics)}
                    </div>
                    """)

                with next_col:
                    if current_index < len(topics) - 1:
                        if st.button(
                            "Next →",
                            type="primary",
                            use_container_width=True
                        ):
                            go(
                                "topic",
                                sid=sid,
                                uid=uid,
                                tid=topics[current_index + 1][0]
                            )
                    else:
                        blue_message("🎉 You reached the last topic!")

                # DELETE TOPIC
                st.divider()

                with st.expander("Delete Topic"):
                    confirm = st.toggle(
                        f"Confirm delete: {current_name}",
                        key=f"confirm_topic_{tid}"
                    )

                    if st.button(
                        "Delete Topic",
                        key=f"delete_topic_{tid}"
                    ):
                        if confirm:
                            delete_topic(tid)

                            remaining = get_topics(uid)

                            if remaining:
                                go(
                                    "topic",
                                    sid=sid,
                                    uid=uid,
                                    tid=remaining[0][0]
                                )
                            else:
                                go("unit", sid=sid, uid=uid)
                        else:
                            blue_message(
                                "Please confirm before deleting."
                            )
