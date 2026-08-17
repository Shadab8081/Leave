"""
Leave Document Generator — Web App (Streamlit)
=================================================
Upload your employee Excel sheet in the browser, click Generate, download a
zip of every leave document, sorted by sponsor.

Run locally with:   streamlit run streamlit_app.py
Deployed for free on Streamlit Community Cloud — see DEPLOY.md.
"""

import shutil
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

import generate_leave_documents as gen

st.set_page_config(page_title="Leave Document Generator", page_icon="📄", layout="centered")

st.title("📄 Automated Leave Application & Clearance Generator")
st.write(
    "Upload your filled employee Excel sheet below. The app will detect each "
    "employee's sponsor (FMCO / Mehan / Al Falak), fill in the correct forms, "
    "and give you a single zip file with everything organized by sponsor."
)

with st.expander("Required Excel columns (click to see)"):
    st.markdown(
        "`EMP ID` · `EMP NAME` · `JOB TITTLE` · `NATIONALITY` · `ID NUMBER` · "
        "`IQAMA EXP` · `TYPE OF VACATION` · `VAC. START DATE` · `VAC. END DATE` · "
        "`NO OF DAYS` · `PROJECT` · `PHONE NUMBER` · `Hiring date` · `Efective Date`\n\n"
        "**ID NUMBER** must start with `7` (FMCO), `C` (Mehan), or `EM-` (Al Falak) "
        "so the sponsor can be detected automatically."
    )

uploaded_file = st.file_uploader("Employee Excel sheet", type=["xlsx", "xlsm"])

generate_clicked = st.button("Generate Leave Documents", type="primary", disabled=uploaded_file is None)

if generate_clicked and uploaded_file is not None:
    workdir = Path(tempfile.mkdtemp(prefix="leavegen_"))
    input_path = workdir / uploaded_file.name
    input_path.write_bytes(uploaded_file.getvalue())
    outdir = workdir / "Leave Documents"

    log_lines = []
    log_box = st.empty()

    def log(msg):
        log_lines.append(str(msg))
        log_box.code("\n".join(log_lines[-25:]), language=None)

    with st.spinner("Generating documents..."):
        try:
            generated, errors = gen.run_generation(input_path, outdir, log=log)
        except Exception as exc:
            st.error(f"Something went wrong: {exc}")
            generated, errors = [], []

    if generated:
        st.success(f"Generated documents for {len(generated)} employee(s).")
        for g in generated:
            st.write(f"- **{g['name']}** — {g['sponsor']} / {g['leave_type']}")

    if errors:
        st.warning(f"{len(errors)} issue(s) found — see Error_Report.txt in the download below.")
        with st.expander("Show error details"):
            for e in errors:
                st.text(e)

    if outdir.exists() and any(outdir.iterdir()):
        zip_path = workdir / "Leave_Documents.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file in outdir.rglob("*"):
                if file.is_file():
                    zf.write(file, file.relative_to(outdir.parent))

        st.download_button(
            "⬇️ Download all documents (zip)",
            data=zip_path.read_bytes(),
            file_name="Leave Documents.zip",
            mime="application/zip",
            type="primary",
        )
    else:
        st.error("No documents were generated. Check the error details above.")

    shutil.rmtree(workdir, ignore_errors=True)

st.divider()
st.caption(
    "This app runs entirely on the server — your Excel file isn't stored anywhere "
    "after you close this page."
)
