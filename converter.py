"""Document conversion helpers for the PDF ↔ Word mini project.

PDF -> DOCX: pdf2docx (text-based PDFs; fidelity varies)
DOCX -> PDF: LibreOffice headless (installed separately)
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


class ConversionError(RuntimeError):
    """A conversion could not complete."""


def find_libreoffice(custom_path: str = "") -> str | None:
    """Find the LibreOffice executable on Windows, macOS, or Linux."""
    if custom_path.strip():
        p = Path(custom_path.strip().strip('"')).expanduser()
        if p.is_file():
            return str(p)
        return None

    for name in ("soffice", "libreoffice", "soffice.exe"):
        installed = shutil.which(name)
        if installed:
            return installed

    candidates: list[Path] = []
    if sys.platform == "win32":
        for root in (os.environ.get("PROGRAMFILES"), os.environ.get("PROGRAMFILES(X86)")):
            if root:
                candidates.append(Path(root) / "LibreOffice" / "program" / "soffice.exe")
        candidates.extend(
            [
                Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
                Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
            ]
        )
    elif sys.platform == "darwin":
        candidates.append(Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"))
    else:
        candidates.extend([Path("/usr/bin/libreoffice"), Path("/usr/bin/soffice")])

    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)
    return None


def pdf_to_word(pdf_path: Path, output_path: Path) -> None:
    """Convert a (preferably text-based) PDF to an editable DOCX."""
    try:
        from pdf2docx import Converter
    except ImportError as exc:
        raise ConversionError("pdf2docx 라이브러리가 없습니다. requirements.txt를 설치해 주세요.") from exc

    converter = None
    try:
        converter = Converter(str(pdf_path))
        converter.convert(str(output_path), start=0, end=None)
    except Exception as exc:
        raise ConversionError(f"PDF → Word 변환 실패: {exc}") from exc
    finally:
        if converter is not None:
            converter.close()

    if not output_path.is_file() or output_path.stat().st_size == 0:
        raise ConversionError("Word 파일이 생성되지 않았습니다. PDF 형식을 확인해 주세요.")


def word_to_pdf(docx_path: Path, output_path: Path, libreoffice_path: str = "") -> None:
    """Convert a DOCX to PDF with an isolated LibreOffice profile."""
    executable = find_libreoffice(libreoffice_path)
    if executable is None:
        raise ConversionError(
            "LibreOffice를 찾지 못했습니다. LibreOffice를 설치하고, "
            "필요하면 화면에서 soffice.exe의 경로를 지정해 주세요."
        )

    with tempfile.TemporaryDirectory(prefix="lo-profile-") as profile_dir:
        profile_uri = Path(profile_dir).resolve().as_uri()
        command = [
            executable,
            f"-env:UserInstallation={profile_uri}",
            "--headless",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            str(output_path.parent),
            str(docx_path),
        ]
        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=120,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            raise ConversionError("변환이 120초를 초과했습니다. 더 작은 파일로 시도해 주세요.") from exc
        except OSError as exc:
            raise ConversionError(f"LibreOffice 실행 실패: {exc}") from exc

        generated = output_path.parent / (docx_path.stem + ".pdf")
        if generated != output_path and generated.is_file():
            shutil.move(str(generated), str(output_path))

        if completed.returncode != 0 or not output_path.is_file() or output_path.stat().st_size == 0:
            detail = (completed.stderr or completed.stdout or "알 수 없는 오류").strip()[:400]
            raise ConversionError(f"Word → PDF 변환 실패: {detail}")


def pdf_to_word_keep_layout(pdf_path: Path, output_path: Path) -> None:
    """Preserve complete PDF page appearance in a DOCX by inserting page images.

    This mode intentionally trades editability for high visual fidelity. It is
    particularly useful for PDFs containing vector diagrams, boxes and charts.
    """
    from io import BytesIO

    try:
        import fitz  # PyMuPDF, installed via pdf2docx
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt
    except ImportError as exc:
        raise ConversionError("페이지 이미지 변환에 필요한 패키지가 없습니다. 라이브러리를 다시 설치해 주세요.") from exc

    try:
        with fitz.open(str(pdf_path)) as pdf:
            if pdf.needs_pass:
                raise ConversionError("비밀번호로 보호된 PDF는 변환할 수 없습니다.")
            if pdf.page_count == 0:
                raise ConversionError("페이지가 없는 PDF입니다.")
            if pdf.page_count > 60:
                raise ConversionError("레이아웃 보존 모드는 60페이지 이하의 PDF만 지원합니다.")

            word = Document()
            for idx, page in enumerate(pdf):
                # Each PDF page becomes a page in Word with matching dimensions.
                if idx == 0:
                    section = word.sections[0]
                else:
                    from docx.enum.section import WD_SECTION_START
                    section = word.add_section(WD_SECTION_START.NEW_PAGE)
                page_width = float(page.rect.width)
                page_height = float(page.rect.height)
                section.page_width = Pt(page_width)
                section.page_height = Pt(page_height)
                # Inner margins leave room for the inline picture's line height.
                section.left_margin = Pt(18)
                section.right_margin = Pt(18)
                section.top_margin = Pt(18)
                section.bottom_margin = Pt(22)
                section.header_distance = Pt(0)
                section.footer_distance = Pt(0)

                target_w = page_width - 36
                target_h = page_height - 52
                scale = min(target_w / page_width, target_h / page_height)
                image_width = page_width * scale
                # 1.8x is enough for readable text without very large DOCX files.
                pix = page.get_pixmap(matrix=fitz.Matrix(1.8, 1.8), alpha=False)
                image_bytes = pix.tobytes("png")

                paragraph = word.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.line_spacing = 1
                paragraph.add_run().add_picture(BytesIO(image_bytes), width=Pt(image_width))

            word.save(str(output_path))
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(f"레이아웃 보존 변환 실패: {exc}") from exc

    if not output_path.is_file() or output_path.stat().st_size == 0:
        raise ConversionError("Word 파일이 생성되지 않았습니다.")


def pdf_to_word_hybrid(pdf_path: Path, output_path: Path) -> None:
    """Convert a simple worksheet PDF to a *normally editable* Word document.

    All non-diagram text becomes ordinary Word paragraphs, whose Enter behaviour
    is governed only by the Normal paragraph style.  A detected figure region
    is rasterized as one image of the *whole* region, including nested boxes,
    lines, arrows, labels, and any text drawn inside the figure.

    This is a form-friendly conversion, not a general-purpose exact PDF editor:
    PDF coordinates and text flow are inherently different from Word.
    """
    from io import BytesIO
    import re

    try:
        import fitz
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.style import WD_STYLE_TYPE
        from docx.shared import Pt
    except ImportError as exc:
        raise ConversionError(
            "혼합 변환에 PyMuPDF 또는 python-docx가 필요합니다. 라이브러리를 다시 설치해 주세요."
        ) from exc

    def diagram_area(page):
        """Find the complete outer bounding region of the main vector diagram."""
        drawings = page.get_drawings()
        if not drawings:
            return None
        model_titles = page.search_for("연구모델")
        descriptions = page.search_for("설명하기")
        if model_titles:
            title_bottom = max(r.y1 for r in model_titles)
            later = [r.y0 for r in descriptions if r.y0 > title_bottom]
            region_bottom = min(later) if later else page.rect.y1 - 10
            in_region = [d['rect'] for d in drawings
                         if d['rect'].y0 >= title_bottom - 3
                         and d['rect'].y1 <= region_bottom + 3]
            # Include even *zero-width/height* vector lines when collecting
            # extents.  An outer PDF rectangle often consists of 4 thin paths.
            if len(in_region) >= 3:
                x0 = min(r.x0 for r in in_region)
                y0 = min(r.y0 for r in in_region)
                x1 = max(r.x1 for r in in_region)
                y1 = max(r.y1 for r in in_region)
                if x1 - x0 >= 100 and y1 - y0 >= 55:
                    return (fitz.Rect(x0 - 1.4, y0 - 1.4, x1 + 1.4, y1 + 1.4)
                            & page.rect)

        # Generic fallback: largest enclosed rectangle with nested objects.
        outer_boxes = []
        for d in drawings:
            rect = d['rect']
            if rect.width <= page.rect.width * .4 or rect.height <= 70:
                continue
            contained = sum((rect.contains(other['rect']) for other in drawings
                             if other is not d))
            if contained >= 2:
                outer_boxes.append(rect)
        if outer_boxes:
            r = max(outer_boxes, key=lambda r: r.width * r.height)
            return (fitz.Rect(r.x0 - 1.4, r.y0 - 1.4,
                              r.x1 + 1.4, r.y1 + 1.4) & page.rect)
        return None

    def create_text_parts(page, diagram):
        chunks = []
        data = page.get_text("dict", sort=True)
        for block in data["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                pieces = [span for span in line.get("spans", [])
                          if span.get("text", "").strip()]
                if not pieces:
                    continue
                bbox = fitz.Rect(line["bbox"])
                # Text *within* the figure is captured by the picture, never
                # omitted or duplicated as a separate Word paragraph.
                if diagram and bbox.intersects(diagram):
                    continue
                chunks.append({
                    "kind": "text", "y": bbox.y0, "bottom": bbox.y1,
                    "x": bbox.x0, "right": bbox.x1,
                    "text": "".join(piece["text"] for piece in pieces),
                    "spans": pieces,
                })
        if diagram:
            chunks.append({"kind": "diagram", "y": diagram.y0,
                           "bottom": diagram.y1, "x": diagram.x0,
                           "rect": diagram})
        chunks.sort(key=lambda item: (item["y"], item["x"]))
        return chunks

    try:
        with fitz.open(str(pdf_path)) as pdf:
            if pdf.needs_pass:
                raise ConversionError("암호로 보호된 PDF는 변환할 수 없습니다.")
            if not pdf.page_count:
                raise ConversionError("PDF에 페이지가 없습니다.")
            if pdf.page_count > 50:
                raise ConversionError("혼합 변환은 최대 50페이지까지 지원합니다.")

            doc = Document()
            first = pdf[0]
            sec = doc.sections[0]
            sec.page_width = Pt(first.rect.width)
            sec.page_height = Pt(first.rect.height)
            sec.left_margin = Pt(85 if first.rect.width > 450 else 36)
            sec.right_margin = sec.left_margin
            sec.top_margin = Pt(66)
            sec.bottom_margin = Pt(42)
            sec.header_distance = Pt(0)
            sec.footer_distance = Pt(0)

            # The essential fix: normal paragraph behavior everywhere.
            # No per-PDF-line space_before, line_spacing, or left_indent.
            normal = doc.styles['Normal']
            normal.font.name = '맑은 고딕'
            normal.font.size = Pt(10.5)
            normal.paragraph_format.space_before = Pt(0)
            normal.paragraph_format.space_after = Pt(0)
            normal.paragraph_format.line_spacing = 1.15

            # Centered headings and figures are styles whose *next* style is
            # Normal. Pressing Enter after either starts a normal paragraph.
            center = doc.styles.add_style('PDF Center Title', WD_STYLE_TYPE.PARAGRAPH)
            center.base_style = normal
            center.next_paragraph_style = normal
            center.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            figure = doc.styles.add_style('PDF Figure', WD_STYLE_TYPE.PARAGRAPH)
            figure.base_style = normal
            figure.next_paragraph_style = normal
            figure.paragraph_format.space_after = Pt(0)
            figure.paragraph_format.space_before = Pt(0)
            figure.paragraph_format.left_indent = Pt(-7)

            for page_idx, page in enumerate(pdf):
                if (abs(page.rect.width - first.rect.width) > 2 or
                    abs(page.rect.height - first.rect.height) > 2):
                    raise ConversionError('혼합 모드는 페이지 크기가 동일한 문서에 적합합니다.')
                if page_idx:
                    doc.add_page_break()

                diagram = diagram_area(page)
                chunks = create_text_parts(page, diagram)
                previous_bottom = float(sec.top_margin.pt)
                for item in chunks:
                    gap = max(0.0, item['y'] - previous_bottom)
                    # PDF blank space becomes *ordinary blank paragraphs*, not
                    # a giant hidden space_before attached to a text line.
                    # This maintains predictable Enter behaviour everywhere.
                    blanks = max(0, min(3, round((gap - 7) / 14)))
                    for _ in range(blanks):
                        doc.add_paragraph(style='Normal')

                    if item['kind'] == 'diagram':
                        pic = doc.add_paragraph(style=figure)
                        # Capture the ENTIRE crop, including all nested shapes.
                        bitmap = page.get_pixmap(
                            matrix=fitz.Matrix(2.5, 2.5),
                            clip=item['rect'], alpha=False,
                        )
                        maximum_width = (sec.page_width.pt - sec.left_margin.pt
                                         - sec.right_margin.pt)
                        width = min(float(item['rect'].width), maximum_width)
                        pic.add_run().add_picture(
                            BytesIO(bitmap.tobytes('png')), width=Pt(width)
                        )
                        previous_bottom = item['bottom']
                        continue

                    text = item['text'].strip()
                    # The PDF stores name/ID with lots of space; replace the
                    # run of spaces with default Word tabs (not special stops).
                    parts = re.split(r'\s{5,}', text)
                    if len(parts) > 1:
                        text = parts[0].rstrip() + '\t\t\t\t\t\t\t' + parts[1].lstrip()
                    if text.endswith(':'):
                        text += ' '
                    centered = (abs((item['x'] + item['right']) / 2
                                    - page.rect.width / 2) < 14 and
                                item['x'] > page.rect.width * 0.25 and
                                item['right'] < page.rect.width * 0.75)
                    p = doc.add_paragraph(style=center if centered else 'Normal')
                    run = p.add_run(text)
                    run.bold = any('Bold' in span.get('font', '')
                                   for span in item['spans'])
                    # Do not override line spacing/space-before on individual
                    # paragraphs: this is what caused huge Enter gaps.
                    previous_bottom = item['bottom']

                # Create an actually editable blank area below the worksheet's
                # explanation prompt. These are NORMAL Word paragraphs.
                if chunks and '설명하기' in ''.join(c.get('text','') for c in chunks[-2:]):
                    for _ in range(3):
                        doc.add_paragraph(style='Normal')

            doc.save(str(output_path))
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(f"도형 보존 + 일반 Word 편집 변환 실패: {exc}") from exc

    if not output_path.is_file() or output_path.stat().st_size == 0:
        raise ConversionError("변환된 Word 파일이 생성되지 않았습니다.")
