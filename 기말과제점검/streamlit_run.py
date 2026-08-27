import base64
import os
import subprocess
import sys
import streamlit as st
from pathlib import Path

_default_lms  = Path(__file__).parent.parent / "lms 제출물"
_default_task = Path(__file__).parent.parent / "과제점검"
LMS_DIR  = Path(os.environ.get("LMS_DIR",  str(_default_lms)))
TASK_DIR = Path(os.environ.get("TASK_DIR", str(_default_task)))
FOLDER_NAME = "기말 과제(발표용)"
PASSWORD = "Smartfarm208!!"


def _convert_with_libreoffice(ppt_path: Path) -> Path:
    subprocess.run(
        [
            "libreoffice", "--headless", "--convert-to", "pdf",
            "--outdir", str(ppt_path.parent.resolve()),
            str(ppt_path.resolve()),
        ],
        check=True,
        capture_output=True,
    )
    return ppt_path.with_suffix(".pdf")


@st.cache_data
def ppt2pdf() -> list[Path]:
    LMS_DIR.mkdir(parents=True, exist_ok=True)

    ppt_files = []
    pdf_paths = []

    for student_dir in sorted(LMS_DIR.iterdir()):
        if not student_dir.is_dir():
            continue
        target = student_dir / FOLDER_NAME
        if not target.exists():
            continue

        poster_pdf, any_pdf, ppt_file = None, None, None
        for f in target.iterdir():
            if "발표" in f.name:
                continue
            if f.suffix.lower() == ".pdf":
                if "포스터" in f.name and poster_pdf is None:
                    poster_pdf = f
                elif any_pdf is None:
                    any_pdf = f
            elif "ppt" in f.name.lower() and ppt_file is None:
                ppt_file = f

        if poster_pdf:
            pdf_paths.append(poster_pdf)
        elif any_pdf:
            pdf_paths.append(any_pdf)
        elif ppt_file:
            ppt_files.append(ppt_file)

    for ppt_path in ppt_files:
        pdf_path = ppt_path.with_suffix(".pdf")
        if pdf_path.exists():
            pdf_paths.append(pdf_path)
            continue
        try:
            pdf_path = _convert_with_libreoffice(ppt_path)
            pdf_paths.append(pdf_path)
            print(f"변환 완료: {ppt_path.parent.parent.name} / {ppt_path.name}")
        except Exception as e:
            print(f"변환 실패: {ppt_path.name} — {e}")

    return pdf_paths


def show_pdf(pdf_path: Path) -> None:
    data = base64.b64encode(pdf_path.read_bytes()).decode("utf-8")
    st.markdown(
        f'<iframe src="data:application/pdf;base64,{data}" width="100%" height="850"></iframe>',
        unsafe_allow_html=True,
    )


def _poster_view() -> None:
    with st.spinner("PDF 로딩 중..."):
        pdf_paths = ppt2pdf()

    if not pdf_paths:
        subdirs = [d.name for d in LMS_DIR.iterdir()] if LMS_DIR.exists() else []
        st.warning(f"변환된 PDF가 없습니다.\n\nLMS_DIR: `{LMS_DIR}`\n하위 항목: `{subdirs}`")
        return

    entries = {p.parent.parent.name: p for p in pdf_paths}
    st.sidebar.title("포스터 리스트")
    selected = st.sidebar.radio("학생 선택", list(entries.keys()), label_visibility="collapsed")
    st.subheader(selected)
    show_pdf(entries[selected])


def home_page() -> None:
    st.title("프원실 기말발표 포스터 모음")
    _poster_view()


def admin_page() -> None:
    st.title("프원실 기말발표 포스터 모음")

    if not st.session_state.get("admin_auth"):
        st.subheader("관리자 로그인")
        pw = st.text_input("비밀번호", type="password")
        if st.button("로그인"):
            if pw == PASSWORD:
                st.session_state["admin_auth"] = True
                st.rerun()
            else:
                st.error("비밀번호가 틀렸습니다.")
        return

    _poster_view()

    st.sidebar.divider()
    if st.sidebar.button("🔄 자료 업데이트", use_container_width=True):
        import io, contextlib, traceback
        from dotenv import load_dotenv

        log_area = st.code("", language="")
        buf = io.StringIO()

        try:
            load_dotenv(dotenv_path=TASK_DIR / ".env", override=True)
            sys.path.insert(0, str(TASK_DIR))
            import auto_login

            with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
                original_dir = os.getcwd()
                os.chdir(str(TASK_DIR))
                try:
                    auto_login.main("프로그래밍원리와실습", "16주차")
                finally:
                    os.chdir(original_dir)

        except Exception as e:
            buf.write(f"\n[오류] {traceback.format_exc()}")

        finally:
            log_area.code(buf.getvalue(), language="")
            ppt2pdf.clear()


def main():
    st.set_page_config(page_title="프원실 기말발표 포스터 모음", layout="wide")
    pg = st.navigation(
        [
            st.Page(home_page, title="포스터 모음", url_path="", default=True),
            st.Page(admin_page, title="Admin", url_path="admin"),
        ],
        position="hidden",
    )
    pg.run()


if __name__ == "__main__":
    main()
