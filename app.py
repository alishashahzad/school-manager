# =============================================================================
# requirements.txt  (copy these lines into a file named requirements.txt)
# -----------------------------------------------------------------------------
# streamlit>=1.50.0
# supabase>=2.5.0
# reportlab>=4.0.0
# pypdf>=4.0.0
# pandas>=2.0.0
# openpyxl>=3.1.0             # needed to read/write Excel (.xlsx) files
# psycopg2-binary>=2.9.9      # optional: lets the app auto-create tables on deploy
# =============================================================================
#
# DEPLOYMENT NOTES
# 1. Create a free project at https://supabase.com
# 2. In Streamlit Cloud -> App settings -> Secrets (or .streamlit/secrets.toml) add:
#       SUPABASE_URL = "https://xxxxxxxx.supabase.co"
#       SUPABASE_KEY = "your-service-role-key"          # use the service_role key
#       DATABASE_URL = "postgresql://postgres:PASSWORD@db.xxxxxxxx.supabase.co:5432/postgres"  # optional
#    - With DATABASE_URL set, tables are created automatically (CREATE TABLE IF NOT EXISTS).
#    - Without it, the app shows the SQL once; paste it in Supabase -> SQL Editor -> Run.
# 3. Everything (students, marks, timetable, uploaded PDF templates, layout
#    coordinates) is stored in Supabase, so nothing is lost when the host restarts.
# =============================================================================

import base64
import copy
import io
import os
from datetime import date, datetime, time as dtime

import pandas as pd
import streamlit as st
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from supabase import Client, create_client

st.set_page_config(page_title="School Management System", page_icon="🏫", layout="wide")

# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------
CLASSES = ["Nursery", "Prep"] + [f"Class {i}" for i in range(1, 11)]
DEFAULT_SUBJECTS = [
    "English", "Urdu", "Mathematics", "Science", "Social Studies",
    "Islamiat", "Computer Science", "Physics", "Chemistry", "Biology",
]
GRADE_A_MIN = 80.0   # A    : 80% and above
GRADE_B_MIN = 60.0   # B    : 60% - 79.99%
GRADE_C_MIN = 40.0   # C    : 40% - 59.99%   (below 40% = Fail)

REQUIRED_TABLES = ["timetable", "students", "results", "templates", "layout_settings"]

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS timetable (
    id          BIGSERIAL PRIMARY KEY,
    class_name  TEXT NOT NULL,
    subject     TEXT NOT NULL,
    exam_date   DATE NOT NULL,
    exam_time   TEXT NOT NULL,
    created_at  TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS students (
    roll_number  TEXT PRIMARY KEY,
    name         TEXT NOT NULL,
    father_name  TEXT NOT NULL,
    class_name   TEXT NOT NULL,
    created_at   TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS results (
    id              BIGSERIAL PRIMARY KEY,
    roll_number     TEXT NOT NULL REFERENCES students(roll_number)
                        ON UPDATE CASCADE ON DELETE CASCADE,
    class_name      TEXT NOT NULL,
    subject         TEXT NOT NULL,
    marks_obtained  NUMERIC NOT NULL,
    total_marks     NUMERIC NOT NULL,
    updated_at      TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE (roll_number, subject)
);

CREATE TABLE IF NOT EXISTS templates (
    name        TEXT PRIMARY KEY,
    file_name   TEXT,
    file_b64    TEXT NOT NULL,
    updated_at  TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS layout_settings (
    key         TEXT PRIMARY KEY,
    value       JSONB NOT NULL,
    updated_at  TIMESTAMPTZ DEFAULT NOW()
);

NOTIFY pgrst, 'reload schema';
"""

# Default stamping coordinates (PDF points, origin = bottom-left, A4 = 595 x 842).
DEFAULT_LAYOUTS = {
    "admit": {
        "fields": {
            "student_name": {"label": "Student Name", "x": 150.0, "y": 640.0, "size": 12.0},
            "father_name": {"label": "Father's Name", "x": 150.0, "y": 615.0, "size": 12.0},
            "roll_number": {"label": "Roll Number", "x": 150.0, "y": 590.0, "size": 12.0},
            "class_name": {"label": "Class", "x": 400.0, "y": 590.0, "size": 12.0},
        },
        "table": {
            "x_subject": 90.0, "x_date": 300.0, "x_time": 430.0,
            "y_start": 520.0, "row_height": 22.0, "size": 11.0, "max_rows": 12.0,
        },
    },
    "result": {
        "fields": {
            "student_name": {"label": "Student Name", "x": 150.0, "y": 650.0, "size": 12.0},
            "father_name": {"label": "Father's Name", "x": 150.0, "y": 625.0, "size": 12.0},
            "roll_number": {"label": "Roll Number", "x": 150.0, "y": 600.0, "size": 12.0},
            "class_name": {"label": "Class", "x": 400.0, "y": 600.0, "size": 12.0},
            "total_obtained": {"label": "Total Obtained", "x": 200.0, "y": 180.0, "size": 12.0},
            "total_marks": {"label": "Total Marks", "x": 330.0, "y": 180.0, "size": 12.0},
            "percentage": {"label": "Percentage", "x": 200.0, "y": 155.0, "size": 12.0},
            "grade": {"label": "Grade", "x": 330.0, "y": 155.0, "size": 12.0},
            "status": {"label": "Result Status", "x": 450.0, "y": 155.0, "size": 12.0},
        },
        "table": {
            "x_subject": 90.0, "x_obtained": 300.0, "x_total": 390.0, "x_grade": 480.0,
            "y_start": 540.0, "row_height": 22.0, "size": 11.0, "max_rows": 14.0,
        },
    },
}

DOC_TITLES = {"admit": "Admit Card", "result": "Result Card"}


# -----------------------------------------------------------------------------
# Supabase connection
# -----------------------------------------------------------------------------
def get_config() -> dict:
    cfg = {}
    for key in ("SUPABASE_URL", "SUPABASE_KEY", "DATABASE_URL"):
        val = None
        try:
            val = st.secrets.get(key)
        except Exception:
            val = None
        val = val or os.environ.get(key) or st.session_state.get(key) or ""
        cfg[key] = str(val).strip()
    return cfg


@st.cache_resource(show_spinner=False)
def _make_client(url: str, key: str) -> Client:
    return create_client(url, key)


def get_client():
    cfg = get_config()
    if not cfg["SUPABASE_URL"] or not cfg["SUPABASE_KEY"]:
        return None
    try:
        return _make_client(cfg["SUPABASE_URL"], cfg["SUPABASE_KEY"])
    except Exception as exc:  # invalid URL / key format
        st.sidebar.error(f"Could not create Supabase client: {exc}")
        return None


def sidebar_connection():
    cfg = get_config()
    configured = bool(cfg["SUPABASE_URL"] and cfg["SUPABASE_KEY"])
    with st.sidebar.expander("🔌 Supabase Connection", expanded=not configured):
        url = st.text_input("SUPABASE_URL", value=cfg["SUPABASE_URL"], placeholder="https://xxxx.supabase.co")
        key = st.text_input("SUPABASE_KEY", value=cfg["SUPABASE_KEY"], type="password")
        dburl = st.text_input(
            "DATABASE_URL (optional)", value=cfg["DATABASE_URL"], type="password",
            help="Postgres connection string. Lets the app create tables automatically.",
        )
        if st.button("Connect"):
            st.session_state["SUPABASE_URL"] = url.strip()
            st.session_state["SUPABASE_KEY"] = key.strip()
            st.session_state["DATABASE_URL"] = dburl.strip()
            st.session_state.pop("schema_ok", None)
            st.session_state.pop("ddl_done", None)
            st.rerun()
        st.caption(
            "For a permanent connection, add these values to Streamlit Secrets "
            "(or environment variables). Values typed here last only for this session."
        )
        if configured:
            st.success("Credentials loaded.")


def run_ddl(database_url: str):
    import psycopg2  # imported lazily; optional dependency

    conn = psycopg2.connect(database_url)
    try:
        with conn:
            with conn.cursor() as cur:
                cur.execute(SCHEMA_SQL)
    finally:
        conn.close()


def missing_tables(client: Client) -> list:
    missing = []
    for table in REQUIRED_TABLES:
        try:
            client.table(table).select("*").limit(1).execute()
        except Exception:
            missing.append(table)
    return missing


def ensure_schema(client: Client):
    cfg = get_config()
    if cfg["DATABASE_URL"] and not st.session_state.get("ddl_done"):
        try:
            run_ddl(cfg["DATABASE_URL"])
            st.session_state["ddl_done"] = True
            st.session_state.pop("schema_ok", None)
        except Exception as exc:
            st.sidebar.warning(f"Automatic table creation failed: {exc}")
            st.session_state["ddl_done"] = True

    if st.session_state.get("schema_ok"):
        return

    missing = missing_tables(client)
    if not missing:
        st.session_state["schema_ok"] = True
        return

    st.title("🏫 School Management System")
    st.error(f"Database tables not found: {', '.join(missing)}")
    st.markdown(
        "Open **Supabase → SQL Editor**, paste the SQL below, press **Run**, "
        "then click **Re-check**. (Or provide `DATABASE_URL` so this happens automatically.)"
    )
    st.code(SCHEMA_SQL, language="sql")
    if st.button("Re-check database"):
        st.session_state.pop("schema_ok", None)
        st.rerun()
    st.stop()


# -----------------------------------------------------------------------------
# Data helpers
# -----------------------------------------------------------------------------
def fetch_all(table: str, filters: dict = None, order: list = None) -> list:
    client = get_client()
    rows, start, page = [], 0, 1000
    while True:
        query = client.table(table).select("*")
        for col, val in (filters or {}).items():
            query = query.eq(col, val)
        for col in order or []:
            query = query.order(col)
        data = query.range(start, start + page - 1).execute().data or []
        rows.extend(data)
        if len(data) < page:
            break
        start += page
    return rows


def fmt_date(value) -> str:
    if isinstance(value, (date, datetime)):
        return value.strftime("%d-%b-%Y")
    try:
        return datetime.strptime(str(value)[:10], "%Y-%m-%d").strftime("%d-%b-%Y")
    except Exception:
        return str(value)


def fmt_time(value) -> str:
    try:
        return datetime.strptime(str(value)[:5], "%H:%M").strftime("%I:%M %p")
    except Exception:
        return str(value)


def fmt_num(value) -> str:
    return f"{float(value):g}"


def grade_for(percentage: float) -> str:
    if percentage >= GRADE_A_MIN:
        return "A"
    if percentage >= GRADE_B_MIN:
        return "B"
    if percentage >= GRADE_C_MIN:
        return "C"
    return "Fail"


def subject_options(class_name: str) -> list:
    subjects = set(DEFAULT_SUBJECTS)
    try:
        for row in fetch_all("timetable", {"class_name": class_name}):
            subjects.add(row["subject"])
    except Exception:
        pass
    return sorted(subjects)


# -----------------------------------------------------------------------------
# Template + layout storage (persisted in Supabase)
# -----------------------------------------------------------------------------
def save_template(doc_key: str, file_name: str, data: bytes):
    client = get_client()
    client.table("templates").upsert(
        {
            "name": doc_key,
            "file_name": file_name,
            "file_b64": base64.b64encode(data).decode("ascii"),
            "updated_at": datetime.utcnow().isoformat(),
        },
        on_conflict="name",
    ).execute()


def load_template(doc_key: str):
    client = get_client()
    res = client.table("templates").select("file_b64").eq("name", doc_key).limit(1).execute()
    if not res.data:
        return None
    return base64.b64decode(res.data[0]["file_b64"])


def template_info(doc_key: str):
    client = get_client()
    res = (
        client.table("templates").select("name,file_name,updated_at").eq("name", doc_key).limit(1).execute()
    )
    return res.data[0] if res.data else None


def merge_layout(default: dict, saved: dict) -> dict:
    out = copy.deepcopy(default)
    for key, val in ((saved or {}).get("fields") or {}).items():
        if key in out["fields"] and isinstance(val, dict):
            out["fields"][key].update(val)
    for key, val in ((saved or {}).get("table") or {}).items():
        if key in out["table"]:
            out["table"][key] = val
    return out


def load_layout(doc_key: str) -> dict:
    client = get_client()
    res = client.table("layout_settings").select("value").eq("key", f"layout_{doc_key}").limit(1).execute()
    saved = res.data[0]["value"] if res.data else {}
    return merge_layout(DEFAULT_LAYOUTS[doc_key], saved)


def save_layout(doc_key: str, layout: dict):
    client = get_client()
    client.table("layout_settings").upsert(
        {"key": f"layout_{doc_key}", "value": layout, "updated_at": datetime.utcnow().isoformat()},
        on_conflict="key",
    ).execute()


# -----------------------------------------------------------------------------
# PDF overlay engine (pypdf + ReportLab)
# -----------------------------------------------------------------------------
def draw_text(c: canvas.Canvas, text, x, y, size, bold=False):
    c.setFont("Helvetica-Bold" if bold else "Helvetica", float(size))
    c.drawString(float(x), float(y), str(text))


def draw_grid(c: canvas.Canvas, width: float, height: float):
    c.saveState()
    c.setLineWidth(0.3)
    c.setFont("Helvetica", 6)
    step = 50
    x = 0
    while x <= width:
        c.setStrokeColorRGB(1, 0.4, 0.4) if x % 100 == 0 else c.setStrokeColorRGB(0.75, 0.75, 0.75)
        c.line(x, 0, x, height)
        c.setFillColorRGB(1, 0, 0)
        c.drawString(x + 1, 2, str(int(x)))
        c.drawString(x + 1, height - 8, str(int(x)))
        x += step
    y = 0
    while y <= height:
        c.setStrokeColorRGB(0.4, 0.4, 1) if y % 100 == 0 else c.setStrokeColorRGB(0.75, 0.75, 0.75)
        c.line(0, y, width, y)
        c.setFillColorRGB(0, 0, 1)
        c.drawString(2, y + 1, str(int(y)))
        c.drawString(width - 18, y + 1, str(int(y)))
        y += step
    c.restoreState()


def render_pdf(template_bytes: bytes, drawers: list, grid: bool = False) -> bytes:
    """Stamp each drawer's content onto a fresh copy of the template's first page."""
    writer = PdfWriter()
    for drawer in drawers:
        base_page = PdfReader(io.BytesIO(template_bytes)).pages[0]
        width = float(base_page.mediabox.width)
        height = float(base_page.mediabox.height)

        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=(width, height))
        drawer(c)
        if grid:
            draw_grid(c, width, height)
        c.showPage()
        c.save()
        buf.seek(0)

        overlay_page = PdfReader(buf).pages[0]
        base_page.merge_page(overlay_page)
        writer.add_page(base_page)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def make_admit_drawer(student: dict, exams: list, layout: dict):
    values = {
        "student_name": student["name"],
        "father_name": student["father_name"],
        "roll_number": student["roll_number"],
        "class_name": student["class_name"],
    }

    def drawer(c):
        for key, cfg in layout["fields"].items():
            draw_text(c, values.get(key, ""), cfg["x"], cfg["y"], cfg["size"], bold=True)
        t = layout["table"]
        y = float(t["y_start"])
        for exam in exams[: int(t["max_rows"])]:
            draw_text(c, exam["subject"], t["x_subject"], y, t["size"])
            draw_text(c, fmt_date(exam["exam_date"]), t["x_date"], y, t["size"])
            draw_text(c, fmt_time(exam["exam_time"]), t["x_time"], y, t["size"])
            y -= float(t["row_height"])

    return drawer


def build_result_summary(student: dict, rows: list) -> dict:
    subjects, total_o, total_t, any_fail = [], 0.0, 0.0, False
    for row in sorted(rows, key=lambda r: r["subject"]):
        obtained = float(row["marks_obtained"])
        total = float(row["total_marks"])
        pct = (obtained / total * 100.0) if total > 0 else 0.0
        grade = grade_for(pct)
        any_fail = any_fail or grade == "Fail"
        total_o += obtained
        total_t += total
        subjects.append({"subject": row["subject"], "obtained": obtained, "total": total, "grade": grade})
    pct_all = (total_o / total_t * 100.0) if total_t > 0 else 0.0
    status = "FAIL" if any_fail else "PASS"
    grade_all = "Fail" if any_fail else grade_for(pct_all)
    return {
        "student": student,
        "subjects": subjects,
        "total_obtained": total_o,
        "total_marks": total_t,
        "percentage": pct_all,
        "grade": grade_all,
        "status": status,
    }


def make_result_drawer(summary: dict, layout: dict):
    student = summary["student"]
    values = {
        "student_name": student["name"],
        "father_name": student["father_name"],
        "roll_number": student["roll_number"],
        "class_name": student["class_name"],
        "total_obtained": fmt_num(summary["total_obtained"]),
        "total_marks": fmt_num(summary["total_marks"]),
        "percentage": f"{summary['percentage']:.2f}%",
        "grade": summary["grade"],
        "status": summary["status"],
    }

    def drawer(c):
        for key, cfg in layout["fields"].items():
            draw_text(c, values.get(key, ""), cfg["x"], cfg["y"], cfg["size"], bold=True)
        t = layout["table"]
        y = float(t["y_start"])
        for sub in summary["subjects"][: int(t["max_rows"])]:
            draw_text(c, sub["subject"], t["x_subject"], y, t["size"])
            draw_text(c, fmt_num(sub["obtained"]), t["x_obtained"], y, t["size"])
            draw_text(c, fmt_num(sub["total"]), t["x_total"], y, t["size"])
            draw_text(c, sub["grade"], t["x_grade"], y, t["size"])
            y -= float(t["row_height"])

    return drawer


# -----------------------------------------------------------------------------
# Page 1: Configuration & Timetables
# -----------------------------------------------------------------------------
def page_timetable():
    st.title("🗓️ Configuration & Exam Timetables")
    client = get_client()

    st.subheader("Add Exam Schedule")
    with st.form("timetable_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        class_name = c1.selectbox("Class", CLASSES)
        subject_pick = c2.selectbox("Subject", DEFAULT_SUBJECTS)
        custom_subject = c2.text_input("Custom subject (overrides dropdown if filled)")
        c3, c4 = st.columns(2)
        exam_date = c3.date_input("Exam Date", value=date.today())
        exam_time = c4.time_input("Exam Time", value=dtime(9, 0))
        submitted = st.form_submit_button("Save Schedule")

    if submitted:
        subject = custom_subject.strip() or subject_pick
        try:
            client.table("timetable").insert(
                {
                    "class_name": class_name,
                    "subject": subject,
                    "exam_date": exam_date.isoformat(),
                    "exam_time": exam_time.strftime("%H:%M"),
                }
            ).execute()
            st.success(f"Saved: {class_name} - {subject} on {fmt_date(exam_date)} at {fmt_time(exam_time.strftime('%H:%M'))}")
        except Exception as exc:
            st.error(f"Database error: {exc}")

    st.divider()
    st.subheader("View Exam Schedules")
    filter_class = st.selectbox("Filter by class", ["All"] + CLASSES, key="tt_filter")
    try:
        rows = fetch_all(
            "timetable",
            None if filter_class == "All" else {"class_name": filter_class},
            ["class_name", "exam_date", "exam_time"],
        )
    except Exception as exc:
        st.error(f"Database error: {exc}")
        return

    if not rows:
        st.info("No exam schedules found.")
        return

    df = pd.DataFrame(
        [
            {
                "ID": r["id"],
                "Class": r["class_name"],
                "Subject": r["subject"],
                "Exam Date": fmt_date(r["exam_date"]),
                "Exam Time": fmt_time(r["exam_time"]),
            }
            for r in rows
        ]
    )
    st.dataframe(df, hide_index=True, width="stretch")

    with st.expander("Delete a schedule entry"):
        labels = {
            f"#{r['id']} | {r['class_name']} | {r['subject']} | {fmt_date(r['exam_date'])}": r["id"] for r in rows
        }
        choice = st.selectbox("Select entry", list(labels.keys()))
        if st.button("Delete selected entry"):
            try:
                client.table("timetable").delete().eq("id", labels[choice]).execute()
                st.success("Entry deleted.")
                st.rerun()
            except Exception as exc:
                st.error(f"Database error: {exc}")


# -----------------------------------------------------------------------------
# Page 2: Student Registration
# -----------------------------------------------------------------------------
def page_students():
    st.title("🎓 Student Registration")
    client = get_client()

    with st.form("student_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        roll = c1.text_input("Roll Number (unique)")
        name = c2.text_input("Student Name")
        father = c1.text_input("Father's Name")
        class_name = c2.selectbox("Class", CLASSES)
        allow_update = st.checkbox("Update the student if this Roll Number already exists")
        submitted = st.form_submit_button("Register Student")

    if submitted:
        roll, name, father = roll.strip(), name.strip(), father.strip()
        if not roll or not name or not father:
            st.error("Roll Number, Name and Father's Name are required.")
        else:
            try:
                existing = client.table("students").select("roll_number").eq("roll_number", roll).execute().data
                if existing and not allow_update:
                    st.error(f"Roll Number '{roll}' already exists. Tick the update box to overwrite it.")
                else:
                    client.table("students").upsert(
                        {"roll_number": roll, "name": name, "father_name": father, "class_name": class_name},
                        on_conflict="roll_number",
                    ).execute()
                    if existing:
                        client.table("results").update({"class_name": class_name}).eq("roll_number", roll).execute()
                    st.success(f"Student '{name}' (Roll {roll}) saved in {class_name}.")
            except Exception as exc:
                st.error(f"Database error: {exc}")

    st.divider()
    st.subheader("Registered Students")
    filter_class = st.selectbox("Filter by class", ["All"] + CLASSES, key="stu_filter")
    try:
        rows = fetch_all(
            "students",
            None if filter_class == "All" else {"class_name": filter_class},
            ["class_name", "roll_number"],
        )
    except Exception as exc:
        st.error(f"Database error: {exc}")
        return

    if not rows:
        st.info("No students registered yet.")
        return

    df = pd.DataFrame(
        [
            {"Roll Number": r["roll_number"], "Name": r["name"], "Father's Name": r["father_name"], "Class": r["class_name"]}
            for r in rows
        ]
    )
    st.dataframe(df, hide_index=True, width="stretch")
    st.caption(f"Total: {len(rows)} student(s)")

    with st.expander("Delete a student (also deletes their marks)"):
        labels = {f"{r['roll_number']} | {r['name']} | {r['class_name']}": r["roll_number"] for r in rows}
        choice = st.selectbox("Select student", list(labels.keys()))
        if st.button("Delete selected student"):
            try:
                client.table("students").delete().eq("roll_number", labels[choice]).execute()
                st.success("Student deleted.")
                st.rerun()
            except Exception as exc:
                st.error(f"Database error: {exc}")


# -----------------------------------------------------------------------------
# Page 3: Marks Entry
# -----------------------------------------------------------------------------
def page_marks():
    st.title("📝 Marks Entry")
    client = get_client()

    c1, c2, c3 = st.columns([1, 1, 1])
    class_name = c1.selectbox("Class", CLASSES, key="marks_class")
    subject_pick = c2.selectbox("Subject", subject_options(class_name), key="marks_subject")
    custom_subject = c3.text_input("Or enter a new subject", key="marks_custom_subject")
    subject = custom_subject.strip() or subject_pick

    try:
        students = fetch_all("students", {"class_name": class_name}, ["roll_number"])
        existing = {
            r["roll_number"]: r for r in fetch_all("results", {"class_name": class_name, "subject": subject})
        }
    except Exception as exc:
        st.error(f"Database error: {exc}")
        return

    if not students:
        st.info(f"No students registered in {class_name}.")
        return

    records = []
    for s in students:
        prev = existing.get(s["roll_number"])
        records.append(
            {
                "Roll Number": s["roll_number"],
                "Student Name": s["name"],
                "Marks Obtained": float(prev["marks_obtained"]) if prev else None,
                "Total Marks": float(prev["total_marks"]) if prev else 100.0,
            }
        )
    df = pd.DataFrame(records)

    st.caption(f"Entering marks for **{class_name} - {subject}**. Leave 'Marks Obtained' empty to skip a student.")
    edited = st.data_editor(
        df,
        key=f"marks_editor_{class_name}_{subject}",
        hide_index=True,
        width="stretch",
        disabled=["Roll Number", "Student Name"],
        column_config={
            "Marks Obtained": st.column_config.NumberColumn("Marks Obtained", min_value=0.0, step=0.5),
            "Total Marks": st.column_config.NumberColumn("Total Marks", min_value=1.0, step=1.0),
        },
    )

    if st.button("💾 Save Marks", type="primary"):
        payload, errors = [], []
        for _, row in edited.iterrows():
            obtained, total = row["Marks Obtained"], row["Total Marks"]
            if pd.isna(obtained):
                continue
            if pd.isna(total) or float(total) <= 0:
                errors.append(f"Roll {row['Roll Number']}: Total Marks must be greater than 0.")
                continue
            if float(obtained) < 0 or float(obtained) > float(total):
                errors.append(f"Roll {row['Roll Number']}: Marks Obtained must be between 0 and Total Marks.")
                continue
            payload.append(
                {
                    "roll_number": row["Roll Number"],
                    "class_name": class_name,
                    "subject": subject,
                    "marks_obtained": float(obtained),
                    "total_marks": float(total),
                    "updated_at": datetime.utcnow().isoformat(),
                }
            )
        if errors:
            for err in errors:
                st.error(err)
        elif not payload:
            st.warning("No marks entered.")
        else:
            try:
                client.table("results").upsert(payload, on_conflict="roll_number,subject").execute()
                st.success(f"Saved marks for {len(payload)} student(s) in {class_name} - {subject}.")
            except Exception as exc:
                st.error(f"Database error: {exc}")


# -----------------------------------------------------------------------------
# Page 4: Custom Layout Upload (PDF Overlays)
# -----------------------------------------------------------------------------
def sample_drawer(doc_key: str, layout: dict):
    student = {
        "roll_number": "1001", "name": "Sample Student Name",
        "father_name": "Sample Father Name", "class_name": "Class 5",
    }
    if doc_key == "admit":
        exams = [
            {"subject": s, "exam_date": f"2026-12-{10 + i:02d}", "exam_time": "09:00"}
            for i, s in enumerate(["English", "Urdu", "Mathematics", "Science", "Islamiat", "Computer Science"])
        ]
        return make_admit_drawer(student, exams, layout)
    rows = [
        {"subject": s, "marks_obtained": m, "total_marks": 100}
        for s, m in [("English", 78), ("Urdu", 85), ("Mathematics", 91), ("Science", 66), ("Islamiat", 88)]
    ]
    return make_result_drawer(build_result_summary(student, rows), layout)


def layout_editor(doc_key: str):
    layout = load_layout(doc_key)
    new_layout = copy.deepcopy(layout)

    with st.form(f"layout_form_{doc_key}"):
        st.markdown("**Single fields** — X / Y position (points from the bottom-left corner) and font size")
        h1, h2, h3, h4 = st.columns([2, 1, 1, 1])
        h1.caption("Field")
        h2.caption("X")
        h3.caption("Y")
        h4.caption("Font size")
        for key, cfg in new_layout["fields"].items():
            c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
            c1.markdown(cfg["label"])
            cfg["x"] = c2.number_input("X", value=float(cfg["x"]), step=1.0, key=f"{doc_key}_{key}_x", label_visibility="collapsed")
            cfg["y"] = c3.number_input("Y", value=float(cfg["y"]), step=1.0, key=f"{doc_key}_{key}_y", label_visibility="collapsed")
            cfg["size"] = c4.number_input("Size", value=float(cfg["size"]), step=1.0, min_value=4.0, key=f"{doc_key}_{key}_s", label_visibility="collapsed")

        st.markdown("**Table block** — column X positions, first row Y, row spacing")
        keys = list(new_layout["table"].keys())
        cols = st.columns(4)
        for i, key in enumerate(keys):
            new_layout["table"][key] = cols[i % 4].number_input(
                key.replace("_", " ").title(),
                value=float(new_layout["table"][key]),
                step=1.0,
                key=f"{doc_key}_table_{key}",
            )

        save_clicked = st.form_submit_button("Save Layout Coordinates")

    if save_clicked:
        try:
            save_layout(doc_key, new_layout)
            st.success(f"{DOC_TITLES[doc_key]} layout saved.")
            layout = new_layout
        except Exception as exc:
            st.error(f"Database error: {exc}")

    grid = st.checkbox("Show coordinate grid on preview", value=True, key=f"{doc_key}_grid")
    if st.button(f"Generate {DOC_TITLES[doc_key]} Preview", key=f"{doc_key}_preview_btn"):
        try:
            template = load_template(doc_key)
            if template is None:
                st.error("Upload and save a template PDF first.")
            else:
                pdf = render_pdf(template, [sample_drawer(doc_key, load_layout(doc_key))], grid=grid)
                st.session_state[f"{doc_key}_preview_pdf"] = pdf
        except Exception as exc:
            st.error(f"Could not build preview: {exc}")

    if st.session_state.get(f"{doc_key}_preview_pdf"):
        st.download_button(
            f"⬇️ Download {DOC_TITLES[doc_key]} Preview",
            data=st.session_state[f"{doc_key}_preview_pdf"],
            file_name=f"{doc_key}_preview.pdf",
            mime="application/pdf",
            key=f"{doc_key}_preview_dl",
        )


def template_section(doc_key: str, uploader_label: str):
    st.subheader(DOC_TITLES[doc_key] + " Template")
    try:
        info = template_info(doc_key)
    except Exception as exc:
        st.error(f"Database error: {exc}")
        return

    if info:
        st.success(f"Stored template: **{info.get('file_name') or 'template.pdf'}** (updated {str(info.get('updated_at'))[:19]})")
    else:
        st.warning("No template uploaded yet.")

    uploaded = st.file_uploader(uploader_label, type=["pdf"], key=f"{doc_key}_uploader")
    if uploaded is not None:
        data = uploaded.getvalue()
        try:
            reader = PdfReader(io.BytesIO(data))
            page = reader.pages[0]
            st.info(
                f"Valid PDF — {len(reader.pages)} page(s). First page size: "
                f"{float(page.mediabox.width):.0f} x {float(page.mediabox.height):.0f} pt "
                "(only the first page is used as the background)."
            )
        except Exception as exc:
            st.error(f"This is not a readable PDF: {exc}")
            return
        if st.button(f"Save {DOC_TITLES[doc_key]} Template", key=f"{doc_key}_save_tpl"):
            try:
                save_template(doc_key, uploaded.name, data)
                st.success("Template saved permanently to Supabase.")
                st.rerun()
            except Exception as exc:
                st.error(f"Database error: {exc}")

    with st.expander(f"✏️ {DOC_TITLES[doc_key]} stamping coordinates & preview", expanded=False):
        layout_editor(doc_key)


def page_layouts():
    st.title("🎨 Custom Layout Upload (PDF Overlays)")
    st.caption(
        "Upload your own card designs. The app stamps student data on top of them at the "
        "coordinates configured below. Use the preview grid to find the exact positions."
    )
    template_section("admit", "Upload Base Admit Card Design (PDF format)")
    st.divider()
    template_section("result", "Upload Base Result Card Design (PDF format)")


# -----------------------------------------------------------------------------
# Page 5: Document Generator
# -----------------------------------------------------------------------------
def page_generator():
    st.title("🖨️ Document Generator")
    class_name = st.selectbox("Class", CLASSES, key="gen_class")
    c1, c2 = st.columns(2)

    # ---- Admit cards -------------------------------------------------------
    if c1.button("Generate Bulk Admit Cards", type="primary"):
        try:
            template = load_template("admit")
            if template is None:
                st.error("No Admit Card template uploaded. Go to 'Custom Layout Upload'.")
            else:
                students = fetch_all("students", {"class_name": class_name}, ["roll_number"])
                exams = fetch_all("timetable", {"class_name": class_name}, ["exam_date", "exam_time"])
                if not students:
                    st.error(f"No students registered in {class_name}.")
                elif not exams:
                    st.error(f"No exam timetable found for {class_name}.")
                else:
                    layout = load_layout("admit")
                    with st.spinner("Generating admit cards..."):
                        drawers = [make_admit_drawer(s, exams, layout) for s in students]
                        pdf = render_pdf(template, drawers)
                    st.session_state["admit_pdf"] = {"class": class_name, "data": pdf, "count": len(students)}
        except Exception as exc:
            st.error(f"Could not generate admit cards: {exc}")

    # ---- Result cards ------------------------------------------------------
    if c2.button("Generate Bulk Result Cards", type="primary"):
        try:
            template = load_template("result")
            if template is None:
                st.error("No Result Card template uploaded. Go to 'Custom Layout Upload'.")
            else:
                students = fetch_all("students", {"class_name": class_name}, ["roll_number"])
                results = fetch_all("results", {"class_name": class_name})
                if not students:
                    st.error(f"No students registered in {class_name}.")
                else:
                    by_roll = {}
                    for r in results:
                        by_roll.setdefault(r["roll_number"], []).append(r)
                    layout = load_layout("result")
                    drawers, skipped = [], []
                    for s in students:
                        rows = by_roll.get(s["roll_number"])
                        if not rows:
                            skipped.append(f"{s['roll_number']} ({s['name']})")
                            continue
                        drawers.append(make_result_drawer(build_result_summary(s, rows), layout))
                    if skipped:
                        st.warning("Skipped (no marks entered): " + ", ".join(skipped))
                    if not drawers:
                        st.error("No marks found for this class.")
                    else:
                        with st.spinner("Generating result cards..."):
                            pdf = render_pdf(template, drawers)
                        st.session_state["result_pdf"] = {"class": class_name, "data": pdf, "count": len(drawers)}
        except Exception as exc:
            st.error(f"Could not generate result cards: {exc}")

    st.divider()
    safe_class = class_name.replace(" ", "_")

    admit = st.session_state.get("admit_pdf")
    if admit:
        st.success(f"Admit cards ready for {admit['class']} ({admit['count']} student(s)).")
        st.download_button(
            "⬇️ Download Bulk Admit Cards (PDF)",
            data=admit["data"],
            file_name=f"admit_cards_{admit['class'].replace(' ', '_')}.pdf",
            mime="application/pdf",
            key="dl_admit",
        )

    result = st.session_state.get("result_pdf")
    if result:
        st.success(f"Result cards ready for {result['class']} ({result['count']} student(s)).")
        st.download_button(
            "⬇️ Download Bulk Result Cards (PDF)",
            data=result["data"],
            file_name=f"result_cards_{result['class'].replace(' ', '_')}.pdf",
            mime="application/pdf",
            key="dl_result",
        )

    st.caption(
        f"Grading: A ≥ {GRADE_A_MIN:g}%, B ≥ {GRADE_B_MIN:g}%, C ≥ {GRADE_C_MIN:g}%, "
        f"Fail < {GRADE_C_MIN:g}%. A student failing any subject is marked FAIL overall."
    )


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# Page 6: Bulk Import (Excel / CSV)
# -----------------------------------------------------------------------------
def clean_cell(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if text.lower() in ("nan", "nat", "none"):
        return ""
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def normalize_class(value) -> str:
    text = clean_cell(value)
    if not text:
        return ""
    if text.isdigit():
        text = f"Class {int(text)}"
    for c in CLASSES:
        if c.lower() == text.lower():
            return c
    return text


def read_upload(uploaded) -> pd.DataFrame:
    if uploaded.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded, dtype=str)
    else:
        df = pd.read_excel(uploaded, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how="all")


def _norm_col(name: str) -> str:
    return str(name).lower().replace("_", " ").replace("'", "").replace("\u2019", "").strip()


def find_col(df: pd.DataFrame, *names):
    lookup = {_norm_col(c): c for c in df.columns}
    for n in names:
        if _norm_col(n) in lookup:
            return lookup[_norm_col(n)]
    return None


def to_excel_bytes(df: pd.DataFrame, sheet: str) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet, index=False)
    return buf.getvalue()


def chunked(items: list, size: int = 500):
    for i in range(0, len(items), size):
        yield items[i : i + size]


def parse_date_cell(value):
    text = clean_cell(value)
    if not text:
        return None
    try:
        return pd.to_datetime(text, dayfirst=True).date()
    except Exception:
        return None


def parse_time_cell(value) -> str:
    text = clean_cell(value)
    if not text:
        return ""
    try:
        return pd.to_datetime(text).strftime("%H:%M")
    except Exception:
        return ""


def show_import_errors(errors: list):
    if errors:
        st.error(f"{len(errors)} row(s) have problems and will be skipped:")
        st.write("\n".join(f"- {e}" for e in errors[:25]))
        if len(errors) > 25:
            st.caption(f"...and {len(errors) - 25} more.")


def import_students_tab():
    client = get_client()
    sample = pd.DataFrame(
        {
            "Roll Number": ["1", "2", "3"],
            "Student Name": ["Ali Raza", "Sara Khan", "Hina Malik"],
            "Father Name": ["Raza Ahmed", "Imran Khan", "Tariq Malik"],
            "Class": ["Class 5", "Class 5", "Class 5"],
        }
    )
    st.download_button(
        "⬇️ Download sample Excel (Students)", to_excel_bytes(sample, "Students"),
        file_name="students_sample.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_stu_sample",
    )
    st.caption("Columns needed: Roll Number, Student Name, Father Name, Class. Existing roll numbers are updated.")
    uploaded = st.file_uploader("Upload students file (.xlsx or .csv)", type=["xlsx", "csv"], key="imp_students")
    if uploaded is None:
        return
    try:
        df = read_upload(uploaded)
    except Exception as exc:
        st.error(f"Could not read the file: {exc}")
        return

    c_roll = find_col(df, "roll number", "roll no", "roll")
    c_name = find_col(df, "student name", "name")
    c_father = find_col(df, "father name", "fathers name", "father")
    c_class = find_col(df, "class")
    missing = [lbl for lbl, col in [("Roll Number", c_roll), ("Student Name", c_name), ("Father Name", c_father), ("Class", c_class)] if col is None]
    if missing:
        st.error("Missing column(s): " + ", ".join(missing))
        return

    rows, errors, seen = [], [], set()
    for idx, r in df.iterrows():
        line = idx + 2
        roll, name, father = clean_cell(r[c_roll]), clean_cell(r[c_name]), clean_cell(r[c_father])
        cls = normalize_class(r[c_class])
        if not (roll and name and father):
            errors.append(f"Row {line}: Roll Number, Student Name and Father Name are required.")
        elif cls not in CLASSES:
            errors.append(f"Row {line}: unknown class '{cls}'. Use one of: {', '.join(CLASSES)}.")
        elif roll in seen:
            errors.append(f"Row {line}: Roll Number '{roll}' appears twice in the file.")
        else:
            seen.add(roll)
            rows.append({"roll_number": roll, "name": name, "father_name": father, "class_name": cls})

    show_import_errors(errors)
    if rows:
        st.success(f"{len(rows)} student(s) ready to import.")
        st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
        if st.button("Import Students", type="primary", key="btn_imp_students"):
            try:
                old = {s["roll_number"]: s["class_name"] for s in fetch_all("students")}
                for chunk in chunked(rows):
                    client.table("students").upsert(chunk, on_conflict="roll_number").execute()
                for r in rows:
                    if r["roll_number"] in old and old[r["roll_number"]] != r["class_name"]:
                        client.table("results").update({"class_name": r["class_name"]}).eq("roll_number", r["roll_number"]).execute()
                st.success(f"Imported {len(rows)} student(s).")
            except Exception as exc:
                st.error(f"Database error: {exc}")


def import_timetable_tab():
    client = get_client()
    sample = pd.DataFrame(
        {
            "Class": ["Class 5"] * 3,
            "Subject": ["English", "Urdu", "Mathematics"],
            "Exam Date": ["2026-12-15", "2026-12-16", "2026-12-17"],
            "Exam Time": ["09:00 AM", "09:00 AM", "09:00 AM"],
        }
    )
    st.download_button(
        "⬇️ Download sample Excel (Timetable)", to_excel_bytes(sample, "Timetable"),
        file_name="timetable_sample.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_tt_sample",
    )
    st.caption("Columns needed: Class, Subject, Exam Date, Exam Time. Dates like 2026-12-15 or 15-12-2026 both work.")
    uploaded = st.file_uploader("Upload timetable file (.xlsx or .csv)", type=["xlsx", "csv"], key="imp_timetable")
    if uploaded is None:
        return
    try:
        df = read_upload(uploaded)
    except Exception as exc:
        st.error(f"Could not read the file: {exc}")
        return

    c_class = find_col(df, "class")
    c_subject = find_col(df, "subject")
    c_date = find_col(df, "exam date", "date")
    c_time = find_col(df, "exam time", "time")
    missing = [lbl for lbl, col in [("Class", c_class), ("Subject", c_subject), ("Exam Date", c_date), ("Exam Time", c_time)] if col is None]
    if missing:
        st.error("Missing column(s): " + ", ".join(missing))
        return

    rows, errors = [], []
    for idx, r in df.iterrows():
        line = idx + 2
        cls, subject = normalize_class(r[c_class]), clean_cell(r[c_subject])
        d, t = parse_date_cell(r[c_date]), parse_time_cell(r[c_time])
        if cls not in CLASSES:
            errors.append(f"Row {line}: unknown class '{cls}'.")
        elif not subject:
            errors.append(f"Row {line}: Subject is empty.")
        elif d is None:
            errors.append(f"Row {line}: Exam Date '{clean_cell(r[c_date])}' is not a valid date.")
        elif not t:
            errors.append(f"Row {line}: Exam Time '{clean_cell(r[c_time])}' is not a valid time.")
        else:
            rows.append({"class_name": cls, "subject": subject, "exam_date": d.isoformat(), "exam_time": t})

    show_import_errors(errors)
    if rows:
        st.success(f"{len(rows)} exam entr(ies) ready to import.")
        preview = pd.DataFrame(rows).assign(exam_time=lambda x: x["exam_time"].map(fmt_time))
        st.dataframe(preview, hide_index=True, width="stretch")
        replace = st.checkbox("Delete the existing timetable of these classes first (avoids duplicates)", value=True, key="tt_replace")
        if st.button("Import Timetable", type="primary", key="btn_imp_tt"):
            try:
                if replace:
                    for cls in sorted({r["class_name"] for r in rows}):
                        client.table("timetable").delete().eq("class_name", cls).execute()
                for chunk in chunked(rows):
                    client.table("timetable").insert(chunk).execute()
                st.success(f"Imported {len(rows)} exam entr(ies).")
            except Exception as exc:
                st.error(f"Database error: {exc}")


def import_marks_tab():
    client = get_client()
    sample = pd.DataFrame(
        {
            "Roll Number": ["1", "2", "3"],
            "Student Name": ["Ali Raza", "Sara Khan", "Hina Malik"],
            "English": [72, 88, 65],
            "Urdu": [65, 91, 70],
            "Mathematics": [30, 95, 62],
        }
    )
    st.download_button(
        "⬇️ Download sample Excel (Marks)", to_excel_bytes(sample, "Marks"),
        file_name="marks_sample.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_marks_sample",
    )
    st.caption(
        "One row per student: a 'Roll Number' column, then one column per subject holding the marks obtained. "
        "Students must already be registered. Empty cells are skipped."
    )
    total_marks = st.number_input("Total Marks for every subject in this file", min_value=1.0, value=100.0, step=1.0, key="imp_total")
    uploaded = st.file_uploader("Upload marks file (.xlsx or .csv)", type=["xlsx", "csv"], key="imp_marks")
    if uploaded is None:
        return
    try:
        df = read_upload(uploaded)
    except Exception as exc:
        st.error(f"Could not read the file: {exc}")
        return

    c_roll = find_col(df, "roll number", "roll no", "roll")
    if c_roll is None:
        st.error("Missing column: Roll Number")
        return
    ignore = {
        "roll number", "roll no", "roll", "student name", "name", "father name", "fathers name",
        "father", "class", "sr", "sr no", "s no", "serial", "total", "percentage", "grade", "result",
    }
    subjects = [c for c in df.columns if _norm_col(c) not in ignore and not str(c).lower().startswith("unnamed")]
    if not subjects:
        st.error("No subject columns found.")
        return
    st.info("Subjects detected: " + ", ".join(subjects))

    try:
        student_class = {s["roll_number"]: s["class_name"] for s in fetch_all("students")}
    except Exception as exc:
        st.error(f"Database error: {exc}")
        return

    payload, errors = {}, []
    for idx, r in df.iterrows():
        line = idx + 2
        roll = clean_cell(r[c_roll])
        if not roll:
            errors.append(f"Row {line}: Roll Number is empty.")
            continue
        if roll not in student_class:
            errors.append(f"Row {line}: Roll Number '{roll}' is not registered.")
            continue
        for subj in subjects:
            text = clean_cell(r[subj])
            if not text:
                continue
            try:
                val = float(text)
            except ValueError:
                errors.append(f"Row {line}, {subj}: '{text}' is not a number.")
                continue
            if val < 0 or val > total_marks:
                errors.append(f"Row {line}, {subj}: {val:g} is outside 0 - {total_marks:g}.")
                continue
            payload[(roll, subj.strip())] = {
                "roll_number": roll,
                "class_name": student_class[roll],
                "subject": subj.strip(),
                "marks_obtained": val,
                "total_marks": float(total_marks),
                "updated_at": datetime.utcnow().isoformat(),
            }

    show_import_errors(errors)
    rows = list(payload.values())
    if rows:
        st.success(f"{len(rows)} mark entr(ies) ready to import.")
        st.dataframe(
            pd.DataFrame(rows)[["roll_number", "class_name", "subject", "marks_obtained", "total_marks"]],
            hide_index=True, width="stretch",
        )
        if st.button("Import Marks", type="primary", key="btn_imp_marks"):
            try:
                for chunk in chunked(rows):
                    client.table("results").upsert(chunk, on_conflict="roll_number,subject").execute()
                st.success(f"Imported {len(rows)} mark entr(ies).")
            except Exception as exc:
                st.error(f"Database error: {exc}")


def page_import():
    st.title("📥 Bulk Import (Excel / CSV)")
    st.caption("Upload your school's existing Excel sheets. Import in this order: Students → Timetable → Marks.")
    tab1, tab2, tab3 = st.tabs(["1. Students", "2. Timetable", "3. Marks"])
    with tab1:
        import_students_tab()
    with tab2:
        import_timetable_tab()
    with tab3:
        import_marks_tab()


def main():
    sidebar_connection()
    client = get_client()
    if client is None:
        st.title("🏫 School Management System")
        st.info("Enter your SUPABASE_URL and SUPABASE_KEY in the sidebar (or Streamlit Secrets) to begin.")
        return

    ensure_schema(client)

    st.sidebar.title("🏫 School Manager")
    page = st.sidebar.radio(
        "Navigation",
        [
            "1. Configuration & Timetables",
            "2. Student Registration",
            "3. Marks Entry",
            "4. Custom Layout Upload",
            "5. Document Generator",
            "6. Bulk Import (Excel/CSV)",
        ],
    )

    if page.startswith("1"):
        page_timetable()
    elif page.startswith("2"):
        page_students()
    elif page.startswith("3"):
        page_marks()
    elif page.startswith("4"):
        page_layouts()
    elif page.startswith("5"):
        page_generator()
    else:
        page_import()


main()
