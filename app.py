"""PDF ↔ Word Converter - Vibe Coding course mini-project."""
from __future__ import annotations

import tempfile
import hashlib
from pathlib import Path

import streamlit as st

from converter import (
    ConversionError, find_libreoffice, pdf_to_word,
    pdf_to_word_hybrid, pdf_to_word_keep_layout, word_to_pdf,
)

st.set_page_config(page_title="PDF ↔ Word 변환기", page_icon="📄", layout="centered")

st.title("📄 PDF ↔ Word 변환기")
st.write("파일을 올리고 버튼만 누르면 PDF와 Word(.docx)를 서로 변환할 수 있습니다.")
st.caption("Python · Streamlit · pdf2docx · LibreOffice | 바이브 코딩 미니 프로젝트")

mode = st.radio("변환 방향 선택", ["PDF → Word", "Word → PDF"], horizontal=True)
input_extension = ".pdf" if mode == "PDF → Word" else ".docx"
output_extension = ".docx" if mode == "PDF → Word" else ".pdf"

pdf_quality_mode = "텍스트 편집 우선"
if mode == "PDF → Word":
    pdf_quality_mode = st.radio(
        "변환 방식 선택",
        ["텍스트 편집 우선", "도형 전체 보존 + 일반 Word 편집 (개선)", "원본 모양 유지 (전체 이미지)"],
        index=1,
        horizontal=False,
        help="도형 영역 전체(큰 테두리, 내부 사각형, 선, 그림)를 한 장의 이미지로 보존합니다. 나머지 텍스트는 일반 Word 문단이라 Enter를 눌러도 비정상적으로 큰 문단 간격이 생기지 않습니다.",
    )
    if pdf_quality_mode == "원본 모양 유지 (전체 이미지)":
        st.info("📐 각 페이지 전체를 이미지로 보존합니다. 글자는 수정할 수 없습니다.")
    elif pdf_quality_mode == "도형 전체 보존 + 일반 Word 편집 (개선)":
        st.success("✅ 그림 전체(내부 작은 사각형까지)는 그대로 유지하고, 그림 밖의 텍스트는 일반 Word처럼 자유롭게 편집할 수 있습니다.")
        st.caption("엔터를 눌러도 일반 문단 간격이 유지됩니다. 그림 내부는 하나의 이미지이므로 도형/글자를 따로 편집할 수 없습니다. 연구모델·양식 PDF에 적합합니다.")
    else:
        st.caption("✏️ 텍스트 편집이 가능하지만 PDF의 도형·사각형·복잡한 서식은 누락될 수 있습니다.")

upload = st.file_uploader(
    f"{input_extension} 파일 업로드",
    type=[input_extension.lstrip(".")],
    help="텍스트 기반 PDF 및 일반적인 Word(.docx) 파일을 권장합니다.",
)

libreoffice_path = ""
if mode == "Word → PDF":
    libreoffice_found = find_libreoffice()
    if libreoffice_found:
        st.success("LibreOffice 확인 완료 — PDF 변환을 사용할 수 있습니다.")
    else:
        st.warning("LibreOffice가 아직 발견되지 않았습니다. 설치 후 다시 실행하세요.")
        with st.expander("LibreOffice 경로 직접 지정하기 (선택)"):
            libreoffice_path = st.text_input(
                "soffice.exe 경로",
                placeholder=r"C:\Program Files\LibreOffice\program\soffice.exe",
            )

if upload is not None:
    st.info(f"선택한 파일: {upload.name} ({len(upload.getvalue()) / 1024:.1f} KB)")
    if len(upload.getvalue()) > 20 * 1024 * 1024:
        st.error("이 예제 프로그램에서는 20MB 이하의 파일만 지원합니다.")

file_key = (mode, pdf_quality_mode, upload.name if upload else None, hashlib.sha256(upload.getvalue()).hexdigest()[:16] if upload else None)
if st.session_state.get("last_file_key") != file_key:
    st.session_state["last_file_key"] = file_key
    st.session_state.pop("converted", None)

convert_clicked = st.button("🔄 변환 시작", type="primary", use_container_width=True, disabled=upload is None)

if convert_clicked and upload is not None:
    if len(upload.getvalue()) > 20 * 1024 * 1024:
        st.error("20MB 이하의 파일을 선택해 주세요.")
    else:
        try:
            with st.spinner("변환 중입니다..."):
                with tempfile.TemporaryDirectory(prefix="vibe-convert-") as temp:
                    base = Path(temp)
                    input_path = base / ("input" + input_extension)
                    output_path = base / ("input" + output_extension)
                    input_path.write_bytes(upload.getvalue())
                    if mode == "PDF → Word":
                        if pdf_quality_mode == "원본 모양 유지 (전체 이미지)":
                            pdf_to_word_keep_layout(input_path, output_path)
                        elif pdf_quality_mode == "도형 전체 보존 + 일반 Word 편집 (개선)":
                            pdf_to_word_hybrid(input_path, output_path)
                        else:
                            pdf_to_word(input_path, output_path)
                    else:
                        word_to_pdf(input_path, output_path, libreoffice_path)
                    result = output_path.read_bytes()

            st.session_state["converted"] = {
                "bytes": result,
                "filename": Path(upload.name).stem + output_extension,
                "mime": (
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    if mode == "PDF → Word" else "application/pdf"
                ),
            }
        except ConversionError as exc:
            st.session_state.pop("converted", None)
            st.error(str(exc))
        except Exception as exc:
            st.session_state.pop("converted", None)
            st.error(f"예상하지 못한 오류: {exc}")

if st.session_state.get("converted"):
    result = st.session_state["converted"]
    st.success("✅ 변환 완료! 아래 버튼을 눌러 파일을 다운로드하세요.")
    st.download_button(
        "⬇️ 변환된 파일 다운로드",
        data=result["bytes"],
        file_name=result["filename"],
        mime=result["mime"],
        use_container_width=True,
    )

st.divider()
st.caption("주의: 개선된 혼합 방식은 그림 영역 전체를 이미지로 보존하고 다른 문단을 일반 Word 텍스트로 만듭니다. 복잡한 PDF에서는 원래 위치·서식과 달라질 수 있으며 그림 내부 자체는 수정할 수 없습니다. 변환 후 업로드 파일은 삭제됩니다.")
