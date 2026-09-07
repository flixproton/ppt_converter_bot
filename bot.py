import io
import json
import logging
import os
import re
import threading
import time
import unicodedata
import subprocess
import tempfile
from http.server import BaseHTTPRequestHandler, HTTPServer

import fitz  # PyMuPDF
from fpdf import FPDF
import google.generativeai as genai
from pptx import Presentation
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

# -------------------------------------------------------------
# CONFIGURATION & BRANDING
# -------------------------------------------------------------
BRAND_NAME = "Yash PPT Converter Bot"
BOT_TOKEN = os.environ.get("BOT_TOKEN", "YOUR_TELEGRAM_BOT_TOKEN_HERE")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
PORT = int(os.environ.get("PORT", 8080))

if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


# -------------------------------------------------------------
# DUMMY SERVER FOR RENDER HEALTH CHECKS
# -------------------------------------------------------------
class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/plain")
        self.end_headers()
        self.wfile.write(f"{BRAND_NAME} is active!".encode("utf-8"))

    def log_message(self, format, *args):
        return


def run_dummy_server():
    server = HTTPServer(("0.0.0.0", PORT), HealthCheckHandler)
    logger.info(f"Dummy health check server running on port {PORT}")
    server.serve_forever()


# -------------------------------------------------------------
# TEXT / MARKDOWN SANITIZERS
# -------------------------------------------------------------
def clean_text_for_pdf(text: str) -> str:
    """Make source text safe for FPDF/Helvetica without changing its meaning."""
    if not text:
        return ""

    replacements = {
        "“": '"',
        "”": '"',
        "‘": "'",
        "’": "'",
        "—": "-",
        "–": "-",
        "…": "...",
        "•": "-",
        "▪": "-",
        "►": ">",
        "✔": "/",
        "✓": "/",
        "→": "->",
        "←": "<-",
        "⇒": "=>",
        "≤": "<=",
        "≥": ">=",
        "≠": "!=",
        "±": "+/-",
        "×": "x",
        "÷": "/",
        "°": " deg ",
        "\u200b": "",
        "\ufeff": "",
        "\xa0": " ",
        "\r": "",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = "".join(ch for ch in text if ch.isprintable() or ch in ["\n", "\t", " "])

    # Prevent FPDF/Helvetica problems from extremely long tokens.
    words = text.split(" ")
    safe_words = []
    for word in words:
        if len(word) > 30:
            chunks = [word[i:i + 28] for i in range(0, len(word), 28)]
            safe_words.append(" ".join(chunks))
        else:
            safe_words.append(word)

    return " ".join(safe_words).strip()


def strip_markdown(text: str) -> str:
    """Remove formatting markers that FPDF Helvetica cannot render."""
    if not text:
        return ""
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
    text = re.sub(r"__(.*?)__", r"\1", text)
    text = re.sub(r"`(.*?)`", r"\1", text)
    text = re.sub(r"^\s*#{1,6}\s*", "", text)
    return text.strip()


def normalize_source_text(text: str) -> str:
    return clean_text_for_pdf(strip_markdown(text))


def looks_like_page_number(text: str, y0: float, page_height: float) -> bool:
    t = text.strip()
    if re.fullmatch(r"\d{1,3}", t):
        return y0 > page_height * 0.78 or y0 < page_height * 0.15
    return bool(
        re.fullmatch(r"(page\s*)?\d{1,3}(\s*/\s*\d{1,3})?", t, re.I)
    )


def looks_like_bullet(text: str) -> bool:
    t = text.strip()
    return bool(
        re.match(
            r"^(?:[-*+•▪►]\s+|\d+[\.\)]\s+|[A-Za-z][\.\)]\s+)",
            t,
        )
    )


def clean_bullet_text(text: str) -> str:
    text = normalize_source_text(text)
    return re.sub(
        r"^(?:[-*+•▪►]\s+|\d+[\.\)]\s+|[A-Za-z][\.\)]\s+)",
        "",
        text,
    ).strip()


# -------------------------------------------------------------
# PROGRESS TRACKER
# IMPORTANT: THE CONVERSION PROGRESS BAR/UI IS KEPT UNCHANGED.
# -------------------------------------------------------------
class ProgressTracker:
    def __init__(self, status_msg, total_items, stage_name="Converting"):
        self.status_msg = status_msg
        self.total = max(1, total_items)
        self.stage_name = stage_name
        self.last_update_time = 0
        self.bar_length = 10

    def _generate_bar(self, current):
        fraction = min(max(current / self.total, 0.0), 1.0)
        filled = int(round(self.bar_length * fraction))
        bar = "█" * filled + "░" * (self.bar_length - filled)
        percent = int(fraction * 100)
        return f"[{bar}] {percent}%"

    async def update(self, current, detail="", force=False):
        now = time.time()
        if not force and (now - self.last_update_time < 1.5):
            return

        self.last_update_time = now
        bar_text = self._generate_bar(current)
        remaining = max(0, self.total - current)

        # Keep this conversion progress message/bar exactly as in the original.
        msg = (
            f"⚡ **{BRAND_NAME} is converting...**\n\n"
            f"`{bar_text}`\n\n"
            f"📊 **Stage:** {self.stage_name}\n"
            f"📄 **Progress:** Slide `{current}` of `{self.total}`\n"
            f"⏳ **Remaining:** `{remaining}` slide(s)\n"
            f"ℹ️ _{detail}_"
        )
        try:
            await self.status_msg.edit_text(msg, parse_mode="Markdown")
        except Exception:
            pass


# -------------------------------------------------------------
# DOWNLOAD PROGRESS HELPERS
# -------------------------------------------------------------
async def download_document_with_progress(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Telegram Bot API does not expose byte-level progress through the simple
    get_file/download_as_bytearray call. We therefore keep a separate download
    status message and do NOT touch the conversion progress bar.
    """
    document = update.message.document
    downloading_msg = await update.message.reply_text(
        "📥 *Downloading file...*\n\n`[░░░░░░░░░░]`",
        parse_mode="Markdown",
    )

    tg_file = await context.bot.get_file(document.file_id)
    raw_bytes = await tg_file.download_as_bytearray()

    try:
        await downloading_msg.edit_text(
            "📥 *Downloading file...*\n\n`[██████████] 100%`\n\n"
            "✅ Download complete.",
            parse_mode="Markdown",
        )
    except Exception:
        pass

    return bytes(raw_bytes), downloading_msg


# -------------------------------------------------------------
# OPTION 1: DIRECT B&W SLIDES CONVERSION
# -------------------------------------------------------------
async def convert_exact_slides_to_bw(file_bytes, file_ext, tracker=None):
    """Convert the original presentation to grayscale while preserving page/slide layout."""
    if file_ext == "pdf":
        src_doc = fitz.open(stream=file_bytes, filetype="pdf")
        out_doc = fitz.open()
        total_pages = len(src_doc)

        if tracker:
            tracker.total = total_pages

        for idx, page in enumerate(src_doc, start=1):
            if tracker:
                await tracker.update(
                    idx,
                    detail="Converting slide to Black & White...",
                )

            pix = page.get_pixmap(colorspace=fitz.csGRAY, dpi=110)
            img_bytes = pix.tobytes("jpeg", jpg_quality=75)
            rect = page.rect
            new_page = out_doc.new_page(width=rect.width, height=rect.height)
            new_page.insert_image(rect, stream=img_bytes)

        if tracker:
            await tracker.update(
                total_pages,
                detail="B&W conversion finished!",
                force=True,
            )

        data = out_doc.tobytes(garbage=4, deflate=True)
        src_doc.close()
        out_doc.close()
        return data

    # PPTX: first try LibreOffice conversion so the actual slide layout is
    # preserved. If LibreOffice is unavailable, fall back to a text-based PDF.
    return await convert_pptx_direct_bw(file_bytes, tracker)


async def convert_pptx_direct_bw(file_bytes, tracker=None):
    prs = Presentation(io.BytesIO(file_bytes))
    total_slides = len(prs.slides)

    if tracker:
        tracker.total = total_slides

    with tempfile.TemporaryDirectory() as tmp:
        pptx_path = os.path.join(tmp, "input.pptx")
        out_dir = os.path.join(tmp, "out")
        os.makedirs(out_dir, exist_ok=True)

        with open(pptx_path, "wb") as f:
            f.write(file_bytes)

        libreoffice = find_libreoffice()
        if libreoffice:
            try:
                subprocess.run(
                    [
                        libreoffice,
                        "--headless",
                        "--convert-to",
                        "pdf",
                        "--outdir",
                        out_dir,
                        pptx_path,
                    ],
                    check=True,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    timeout=180,
                )
                pdf_path = os.path.join(out_dir, "input.pdf")
                if os.path.exists(pdf_path):
                    src_doc = fitz.open(pdf_path)
                    out_doc = fitz.open()

                    for idx, page in enumerate(src_doc, start=1):
                        if tracker:
                            await tracker.update(
                                min(idx, total_slides),
                                detail="Converting slide to Black & White...",
                            )
                        pix = page.get_pixmap(colorspace=fitz.csGRAY, dpi=110)
                        img_bytes = pix.tobytes("jpeg", jpg_quality=75)
                        new_page = out_doc.new_page(
                            width=page.rect.width,
                            height=page.rect.height,
                        )
                        new_page.insert_image(page.rect, stream=img_bytes)

                    if tracker:
                        await tracker.update(
                            total_slides,
                            detail="B&W conversion finished!",
                            force=True,
                        )

                    result = out_doc.tobytes(garbage=4, deflate=True)
                    src_doc.close()
                    out_doc.close()
                    return result
            except Exception as exc:
                logger.warning("LibreOffice PPTX conversion failed: %s", exc)

    # Safe fallback when LibreOffice is not installed.
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=10)

    for idx, slide in enumerate(prs.slides, start=1):
        if tracker:
            await tracker.update(
                idx,
                detail="Processing slide layout...",
            )

        pdf.add_page()
        pdf.set_draw_color(180, 180, 180)
        pdf.rect(10, 10, 277, 190)

        title = ""
        if slide.shapes.title and slide.shapes.title.text:
            title = normalize_source_text(slide.shapes.title.text)

        if title:
            pdf.set_xy(15, 15)
            pdf.set_font("Helvetica", "B", 16)
            pdf.set_text_color(0, 0, 0)
            pdf.multi_cell(267, 8, title)
            pdf.ln(5)

        pdf.set_font("Helvetica", "", 12)
        pdf.set_text_color(40, 40, 40)

        for shape in slide.shapes:
            if shape == slide.shapes.title or not shape.has_text_frame:
                continue
            for p in shape.text_frame.paragraphs:
                text = normalize_source_text(p.text)
                if text:
                    pdf.set_x(15)
                    prefix = "- " if looks_like_bullet(text) else ""
                    pdf.multi_cell(267, 6, prefix + clean_bullet_text(text) if prefix else text)
                    pdf.ln(1)

    if tracker:
        await tracker.update(
            total_slides,
            detail="B&W conversion finished!",
            force=True,
        )

    return bytes(pdf.output())


def find_libreoffice():
    for candidate in ("libreoffice", "soffice"):
        try:
            result = subprocess.run(
                ["which", candidate],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=5,
            )
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout.strip()
        except Exception:
            pass
    return None


# -------------------------------------------------------------
# SOURCE EXTRACTION
# -------------------------------------------------------------
def parse_pdf_slide_local(page, page_num):
    """Extract source text in reading order without inventing content."""
    page_dict = page.get_text("dict")
    page_height = page.rect.height

    raw_items = []
    total_text_parts = []

    for block in page_dict.get("blocks", []):
        if block.get("type") != 0:
            continue

        for line in block.get("lines", []):
            line_text = ""
            max_size = 0
            x0, y0, x1, y1 = line.get("bbox", (0, 0, 0, 0))

            for span in line.get("spans", []):
                t = normalize_source_text(span.get("text", ""))
                if t:
                    line_text += (" " if line_text else "") + t
                    max_size = max(max_size, float(span.get("size", 10)))

            line_text = line_text.strip()
            if not line_text:
                continue
            if looks_like_page_number(line_text, y0, page_height):
                continue

            raw_items.append(
                {
                    "text": line_text,
                    "x0": x0,
                    "y0": y0,
                    "x1": x1,
                    "y1": y1,
                    "size": max_size,
                }
            )
            total_text_parts.append(line_text)

    if not raw_items:
        return {
            "num": page_num,
            "title": f"Slide {page_num}",
            "is_diagram": True,
            "raw_text": "",
            "elements": [],
        }

    # Candidate title = largest text near top. Do not fabricate a title if there
    # is no strong title candidate.
    top_candidates = [
        it for it in raw_items
        if it["y0"] < page_height * 0.30
    ]
    title_item = None
    if top_candidates:
        largest = max(top_candidates, key=lambda it: it["size"])
        if largest["size"] >= 16:
            title_item = largest

    title_text = title_item["text"] if title_item else ""
    content_items = [it for it in raw_items if it is not title_item]
    content_items.sort(key=lambda it: (it["y0"], it["x0"]))

    elements = []
    for item in content_items:
        txt = item["text"]
        if looks_like_bullet(txt):
            elements.append({"type": "bullet", "text": clean_bullet_text(txt)})
        elif item["size"] >= 15:
            elements.append({"type": "subheading", "text": txt})
        else:
            elements.append({"type": "paragraph", "text": txt})

    # Merge only obvious broken continuation lines. This is reorganization, not
    # content generation.
    final_elements = []
    for elem in elements:
        if (
            elem["type"] == "paragraph"
            and final_elements
            and final_elements[-1]["type"] == "paragraph"
        ):
            prev = final_elements[-1]["text"]
            # Preserve slide wording; join lines rather than rewriting them.
            if not prev.endswith((".", ":", "?", "!", ";")):
                final_elements[-1]["text"] = (prev + " " + elem["text"]).strip()
            else:
                final_elements.append(elem)
        else:
            final_elements.append(elem)

    raw_text = " ".join(total_text_parts).strip()

    # A slide with little/no text is a diagram candidate. We do not claim what
    # the diagram means; the original slide image is retained instead.
    is_diagram = len(raw_text) < 60 or not final_elements

    return {
        "num": page_num,
        "title": title_text or f"Slide {page_num}",
        "is_diagram": is_diagram,
        "raw_text": raw_text,
        "elements": final_elements,
    }


def extract_pptx_slide(slide, slide_num):
    """Extract PPTX text in shape/paragraph order."""
    title_shape = slide.shapes.title
    title = normalize_source_text(title_shape.text) if title_shape and title_shape.text else ""

    elements = []
    all_text = []

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        if shape == title_shape:
            continue

        for paragraph in shape.text_frame.paragraphs:
            txt = normalize_source_text(paragraph.text)
            if not txt:
                continue

            all_text.append(txt)

            # Detect bullets using PowerPoint paragraph metadata as well as text.
            has_bullet = False
            try:
                has_bullet = bool(paragraph.level > 0)
            except Exception:
                pass

            if looks_like_bullet(txt) or has_bullet:
                elements.append(
                    {"type": "bullet", "text": clean_bullet_text(txt)}
                )
            else:
                # Keep paragraph text intact. The notes builder decides the
                # visual style, not the content.
                elements.append({"type": "paragraph", "text": txt})

    raw_text = " ".join(all_text).strip()

    # Do not invent a description of a visual. Retain the original slide as the
    # diagram figure when there is very little text.
    is_diagram = len(raw_text) < 60

    return {
        "num": slide_num,
        "title": title or f"Slide {slide_num}",
        "is_diagram": is_diagram,
        "raw_text": raw_text,
        "elements": elements,
    }


def extract_source_slides(file_bytes, file_ext, tracker=None):
    if file_ext == "pdf":
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        slides = []
        if tracker:
            tracker.total = len(doc)

        for idx, page in enumerate(doc, start=1):
            if tracker:
                # Same conversion progress bar; only the detail text changes.
                import asyncio
                # This function is sync, so caller handles tracker updates.
            slides.append(parse_pdf_slide_local(page, idx))
        return slides, doc

    prs = Presentation(io.BytesIO(file_bytes))
    slides = [extract_pptx_slide(slide, i) for i, slide in enumerate(prs.slides, 1)]
    return slides, None


# -------------------------------------------------------------
# AI STUDY NOTES ENGINE
# -------------------------------------------------------------
def analyze_and_structure_with_ai(raw_slides_data):
    """
    AI is used only as a formatting/reorganization pass.
    It is explicitly prohibited from adding facts, examples, explanations,
    missing definitions, or conclusions not present in the source.
    """
    if not GEMINI_API_KEY:
        return None

    try:
        # A current, generally available Flash model can be overridden with an
        # environment variable if needed.
        model_name = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
        model = genai.GenerativeModel(model_name)

        slides_summary = []
        for slide in raw_slides_data:
            slides_summary.append(
                {
                    "slide_num": slide.get("num", slide.get("slide_num")),
                    "title": slide.get("title", ""),
                    "raw_text": slide.get("raw_text", ""),
                    "elements": slide.get("elements", []),
                    "is_diagram": slide.get("is_diagram", False),
                }
            )

        prompt = f"""
You are a strict source-to-notes formatter.

Your job is ONLY to reorganize and clean the exact information supplied in
Raw Slides Input. The final notes must be useful for studying, but they must
remain faithful to the source.

NON-NEGOTIABLE SOURCE RULES:
1. NEVER invent facts, definitions, examples, dates, names, explanations,
   transitions, conclusions, or relationships.
2. NEVER fill a missing/unfinished sentence with your own knowledge.
3. NEVER "correct" a fact just because it looks unusual. Preserve the source.
4. You MAY fix obvious line breaks, duplicated whitespace, and formatting noise.
5. You MAY group source statements under headings that already exist in the
   source, or use neutral structural labels such as "Key Points" only when
   the source contains corresponding points.
6. If wording is incomplete or awkward, preserve the wording instead of
   completing it.
7. Do not create a section just because a topic would normally have one.
8. Do not summarize away important source details.
9. Keep the original slide number attached to its content.
10. For diagrams/figures with little text, set is_diagram=true. Do NOT describe
    what the diagram means unless that meaning is explicitly written in the
    source text. Use an empty caption when the source gives no caption.
11. Preserve chronology, lists, categories, and comparisons when the source
    contains them.
12. Do not output Markdown tables because the PDF renderer uses a simple notes
    layout.

NOTES FORMATTING:
- Use "subheading" for an actual source heading/category.
- Use "paragraph" for definitions or continuous source prose.
- Use "bullet" for source bullet/list items.
- If a source slide has a heading followed by bullets, keep that structure.
- Do not manufacture "Key Points" or "Summary" content.
- Do not add an introduction or conclusion.

Return ONLY valid JSON matching this exact schema:
[
  {{
    "slide_num": 1,
    "title": "source title",
    "is_diagram": false,
    "caption": "",
    "elements": [
      {{"type": "subheading", "text": "source heading"}},
      {{"type": "paragraph", "text": "source text"}},
      {{"type": "bullet", "text": "source bullet"}}
    ]
  }}
]

Raw Slides Input:
{json.dumps(slides_summary, ensure_ascii=False)}
"""

        response = model.generate_content(prompt)
        text_resp = (response.text or "").strip()

        if text_resp.startswith("```"):
            text_resp = re.sub(r"^```(?:json)?\s*", "", text_resp)
            text_resp = re.sub(r"\s*```$", "", text_resp)

        data = json.loads(text_resp)
        if not isinstance(data, list):
            return None

        # Basic validation: every returned item must have source slide identity.
        validated = []
        source_by_num = {
            int(s["num"]): s for s in raw_slides_data if "num" in s
        }

        for item in data:
            try:
                slide_num = int(item["slide_num"])
            except Exception:
                continue

            if slide_num not in source_by_num:
                continue

            if not isinstance(item.get("elements", []), list):
                item["elements"] = []

            validated.append(item)

        if not validated:
            return None

        # Ensure no slide disappears because the model returned incomplete JSON.
        returned_nums = {x["slide_num"] for x in validated}
        for source in raw_slides_data:
            if source["num"] not in returned_nums:
                validated.append(
                    {
                        "slide_num": source["num"],
                        "title": source.get("title", f"Slide {source['num']}"),
                        "is_diagram": source.get("is_diagram", False),
                        "caption": "",
                        "elements": source.get("elements", []),
                    }
                )

        validated.sort(key=lambda x: x["slide_num"])
        return validated

    except Exception as exc:
        logger.warning("Gemini structuring skipped: %s", exc)
        return None


# -------------------------------------------------------------
# A4 REAL STUDY NOTES PDF BUILDER
# -------------------------------------------------------------
class BWNotesPDF(FPDF):
    def header(self):
        self.set_x(self.l_margin)
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(100, 100, 100)
        self.cell(
            self.epw,
            6,
            clean_text_for_pdf(f"Study Notes | {BRAND_NAME}"),
            border=0,
            align="R",
        )
        self.ln(8)

    def footer(self):
        self.set_y(-12)
        self.set_x(self.l_margin)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(120, 120, 120)
        self.cell(
            self.epw,
            8,
            clean_text_for_pdf(
                f"Page {self.page_no()} | Generated by {BRAND_NAME}"
            ),
            align="C",
        )


def ensure_space(pdf, needed_mm=12):
    if pdf.get_y() + needed_mm > pdf.h - pdf.b_margin:
        pdf.add_page()


def safe_write_text(
    pdf,
    text,
    font_size=10,
    is_bold=False,
    is_italic=False,
    color=(30, 30, 30),
    prefix="",
    indent=0,
):
    if not text:
        return

    style = "B" if is_bold else ("I" if is_italic else "")
    pdf.set_font("Helvetica", style, font_size)
    pdf.set_text_color(*color)

    full_text = normalize_source_text(f"{prefix}{text}")
    if not full_text:
        return

    line_height = max(4.5, font_size * 0.48)
    pdf.set_x(pdf.l_margin + indent)

    try:
        pdf.multi_cell(
            w=max(20, pdf.epw - indent),
            h=line_height,
            text=full_text,
        )
        pdf.ln(1)
    except Exception:
        pdf.set_x(pdf.l_margin + indent)
        pdf.multi_cell(
            w=max(20, pdf.epw - indent),
            h=line_height,
            text=full_text[:80],
        )
        pdf.ln(1)


def render_source_pdf_figure(src_doc, slide_num, pdf):
    if not src_doc or not (0 < slide_num <= len(src_doc)):
        return False

    try:
        page = src_doc[slide_num - 1]
        pix = page.get_pixmap(colorspace=fitz.csGRAY, dpi=110)
        img_data = pix.tobytes("jpeg", jpg_quality=75)

        img_w = min(170, pdf.epw)
        img_h = (img_w / page.rect.width) * page.rect.height

        if pdf.get_y() + img_h + 10 > pdf.h - pdf.b_margin:
            pdf.add_page()

        img_stream = io.BytesIO(img_data)
        pdf.image(
            img_stream,
            x=pdf.l_margin,
            y=pdf.get_y(),
            w=img_w,
        )
        pdf.set_y(pdf.get_y() + img_h + 3)
        return True
    except Exception as exc:
        logger.warning("Diagram embed error: %s", exc)
        return False


def convert_pptx_to_pdf_for_figures(file_bytes):
    """Use LibreOffice to create a faithful source PDF for PPTX diagram figures."""
    libreoffice = find_libreoffice()
    if not libreoffice:
        return None

    tmp = tempfile.TemporaryDirectory()
    pptx_path = os.path.join(tmp.name, "source.pptx")
    out_dir = os.path.join(tmp.name, "pdf")
    os.makedirs(out_dir, exist_ok=True)

    with open(pptx_path, "wb") as f:
        f.write(file_bytes)

    try:
        subprocess.run(
            [
                libreoffice,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                out_dir,
                pptx_path,
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=180,
        )
        pdf_path = os.path.join(out_dir, "source.pdf")
        if not os.path.exists(pdf_path):
            tmp.cleanup()
            return None

        doc = fitz.open(pdf_path)

        # Keep temp directory alive by attaching it to the document.
        doc._ppt_tmpdir = tmp
        return doc
    except Exception as exc:
        logger.warning("PPTX figure rendering skipped: %s", exc)
        tmp.cleanup()
        return None


def build_final_notes_pdf(slides_data, src_pdf_doc=None):
    """
    Creates a conventional A4 study-notes document:
    title -> source content -> subheadings -> bullets -> diagrams.
    It does not generate new facts.
    """
    pdf = BWNotesPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)

    # Cover/title page
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 22)
    pdf.set_text_color(0, 0, 0)
    pdf.multi_cell(pdf.epw, 11, clean_text_for_pdf("Study Notes"))
    pdf.ln(2)

    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(80, 80, 80)
    pdf.multi_cell(
        pdf.epw,
        6,
        clean_text_for_pdf(
            "Source-faithful notes reorganized from the uploaded presentation."
        ),
    )
    pdf.ln(7)

    # Do not create a fake summary/key-points section.
    pdf.set_draw_color(190, 190, 190)
    pdf.line(pdf.l_margin, pdf.get_y(), pdf.l_margin + pdf.epw, pdf.get_y())
    pdf.ln(6)

    for item in slides_data:
        slide_num = int(item.get("slide_num", item.get("num", 1)))
        title = normalize_source_text(
            item.get("title", f"Slide {slide_num}")
        )
        is_diagram = bool(item.get("is_diagram", False))
        elements = item.get("elements", [])
        caption = normalize_source_text(item.get("caption", ""))

        ensure_space(pdf, 20)

        # Source slide marker
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_text_color(110, 110, 110)
        pdf.cell(
            pdf.epw,
            5,
            clean_text_for_pdf(f"SOURCE SLIDE {slide_num}"),
            align="L",
        )
        pdf.ln(5)

        # Main heading
        title_size = 15 if len(title) <= 55 else 12
        safe_write_text(
            pdf,
            title,
            font_size=title_size,
            is_bold=True,
            color=(0, 0, 0),
        )

        # Diagram/figure comes before text when the slide is primarily visual.
        if is_diagram and src_pdf_doc:
            if render_source_pdf_figure(src_pdf_doc, slide_num, pdf):
                if caption:
                    safe_write_text(
                        pdf,
                        caption,
                        font_size=8.5,
                        is_italic=True,
                        color=(90, 90, 90),
                        prefix="Figure: ",
                    )
            elif caption:
                safe_write_text(
                    pdf,
                    caption,
                    font_size=8.5,
                    is_italic=True,
                    color=(90, 90, 90),
                )

        for elem in elements:
            e_type = elem.get("type", "paragraph")
            e_text = normalize_source_text(elem.get("text", ""))

            if not e_text:
                continue

            if e_type == "subheading":
                ensure_space(pdf, 12)
                pdf.ln(2)
                safe_write_text(
                    pdf,
                    e_text,
                    font_size=11,
                    is_bold=True,
                    color=(20, 20, 20),
                )
            elif e_type == "bullet":
                ensure_space(pdf, 8)
                safe_write_text(
                    pdf,
                    e_text,
                    font_size=10,
                    color=(35, 35, 35),
                    prefix="- ",
                    indent=3,
                )
            else:
                ensure_space(pdf, 10)
                safe_write_text(
                    pdf,
                    e_text,
                    font_size=10,
                    color=(20, 20, 20),
                    indent=1,
                )

        pdf.ln(5)

    return bytes(pdf.output())


# -------------------------------------------------------------
# TELEGRAM BOT HANDLERS
# -------------------------------------------------------------
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome_text = (
        f"👋 **Welcome to {BRAND_NAME}!**\n\n"
        "📄 Send me any **PPTX** or **PDF presentation**, and choose your conversion mode:\n\n"
        "🖨️ **Option 1: Direct B&W Slides**\n"
        "• Converts your exact presentation into ink-saving Black & White (No layout changes)\n\n"
        "📝 **Option 2: AI Study Notes**\n"
        "• Reorganizes source content into real A4 study notes without inventing missing information\n"
        "• Preserves source diagrams/figures when they can be rendered\n\n"
        "🚀 **Send your presentation file now!**"
    )
    await update.message.reply_text(welcome_text, parse_mode="Markdown")


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    file_name = document.file_name or "presentation"
    file_ext = file_name.rsplit(".", 1)[-1].lower() if "." in file_name else ""

    if file_ext not in ["pptx", "pdf"]:
        await update.message.reply_text(
            "⚠️ Please upload a valid **.pptx** or **.pdf** slide file.",
            parse_mode="Markdown",
        )
        return

    try:
        file_bytes, downloading_msg = await download_document_with_progress(
            update,
            context,
        )

        context.user_data["pending_file"] = {
            "bytes": file_bytes,
            "name": file_name,
            "ext": file_ext,
        }

        keyboard = [
            [
                InlineKeyboardButton(
                    "🖨️ 1. Direct B&W Slides (Exact Copy)",
                    callback_data="mode_exact_bw",
                )
            ],
            [
                InlineKeyboardButton(
                    "📝 2. AI Study Notes (Full Restructure)",
                    callback_data="mode_ai_notes",
                )
            ],
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)

        await downloading_msg.edit_text(
            f"✅ **File received:** `{file_name}`\n\n"
            "👉 **Please choose your conversion mode:**",
            reply_markup=reply_markup,
            parse_mode="Markdown",
        )
    except Exception as exc:
        logger.error("File download error: %s", exc, exc_info=True)
        await update.message.reply_text(
            f"❌ File download failed: `{str(exc)}`",
            parse_mode="Markdown",
        )


async def handle_callback_query(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    mode = query.data
    file_info = context.user_data.get("pending_file")

    if not file_info:
        await query.edit_message_text(
            "⚠️ File session expired. Please re-upload your presentation.",
            parse_mode="Markdown",
        )
        return

    file_bytes = file_info["bytes"]
    file_name = file_info["name"]
    file_ext = file_info["ext"]
    base_name = os.path.splitext(file_name)[0]

    status_msg = await query.edit_message_text(
        f"⚡ **{BRAND_NAME} is starting...**\n`[░░░░░░░░░░] 0%`",
        parse_mode="Markdown",
    )

    tracker = None
    pdf_doc = None
    figure_doc = None

    try:
        # Preserve the original conversion progress bar and message exactly.
        total_items = 10
        if file_ext == "pdf":
            probe_doc = fitz.open(stream=file_bytes, filetype="pdf")
            total_items = max(1, len(probe_doc))
            probe_doc.close()
        else:
            probe_prs = Presentation(io.BytesIO(file_bytes))
            total_items = max(1, len(probe_prs.slides))

        tracker = ProgressTracker(
            status_msg=status_msg,
            total_items=total_items,
            stage_name="Processing",
        )

        if mode == "mode_exact_bw":
            # ---------------- MODE 1: DIRECT B&W SLIDES ----------------
            tracker.stage_name = "Generating B&W Slides"
            output_pdf_bytes = await convert_exact_slides_to_bw(
                file_bytes,
                file_ext,
                tracker=tracker,
            )
            output_filename = (
                f"{base_name}_BW_Slides_{BRAND_NAME.replace(' ', '_')}.pdf"
            )
            caption_text = (
                f"✅ **Direct Black & White Slides by {BRAND_NAME}!**\n"
                "🖨️ Exact layout preserved for ink-saving printing."
            )

        elif mode == "mode_ai_notes":
            # ---------------- MODE 2: SOURCE-FAITHFUL STUDY NOTES ----------------
            tracker.stage_name = "Scanning & Structuring"

            if file_ext == "pdf":
                pdf_doc = fitz.open(stream=file_bytes, filetype="pdf")
                tracker.total = max(1, len(pdf_doc))
                raw_slides = []

                for idx, page in enumerate(pdf_doc, start=1):
                    await tracker.update(
                        idx,
                        detail="Scanning text & diagrams...",
                    )
                    raw_slides.append(parse_pdf_slide_local(page, idx))
            else:
                prs = Presentation(io.BytesIO(file_bytes))
                tracker.total = max(1, len(prs.slides))
                raw_slides = []

                for idx, slide in enumerate(prs.slides, start=1):
                    await tracker.update(
                        idx,
                        detail="Scanning slide content...",
                    )
                    raw_slides.append(extract_pptx_slide(slide, idx))

            # AI gets only extracted source material and is not allowed to add facts.
            await tracker.update(
                tracker.total,
                detail="AI organizing notes & timelines...",
                force=True,
            )
            structured_slides = (
                analyze_and_structure_with_ai(raw_slides)
                if GEMINI_API_KEY
                else None
            )
            final_data = structured_slides if structured_slides else raw_slides

            # For PPTX, render source slides to PDF so visual/diagram slides can
            # be embedded in the notes when LibreOffice is available.
            if file_ext == "pptx":
                figure_doc = convert_pptx_to_pdf_for_figures(file_bytes)

            await tracker.update(
                tracker.total,
                detail="Building study notes PDF...",
                force=True,
            )
            output_pdf_bytes = build_final_notes_pdf(
                final_data,
                src_pdf_doc=pdf_doc or figure_doc,
            )

            output_filename = (
                f"{base_name}_Study_Notes_{BRAND_NAME.replace(' ', '_')}.pdf"
            )
            caption_text = (
                f"✅ **Study Notes by {BRAND_NAME}!**\n"
                "📝 Source-faithful notes with headings, paragraphs, bullets "
                "and original visual figures where available."
            )

        else:
            raise ValueError("Unknown conversion mode.")

        # Send the generated file to the user.
        await query.message.reply_document(
            document=io.BytesIO(output_pdf_bytes),
            filename=output_filename,
            caption=caption_text,
            parse_mode="Markdown",
        )

        await status_msg.delete()

    except Exception as exc:
        logger.error("Processing error: %s", exc, exc_info=True)
        try:
            await status_msg.edit_text(
                f"❌ Conversion failed: `{str(exc)}`",
                parse_mode="Markdown",
            )
        except Exception:
            pass

    finally:
        try:
            if pdf_doc:
                pdf_doc.close()
        except Exception:
            pass

        try:
            if figure_doc:
                figure_doc.close()
                tmp = getattr(figure_doc, "_ppt_tmpdir", None)
                if tmp:
                    tmp.cleanup()
        except Exception:
            pass

        context.user_data.pop("pending_file", None)


# -------------------------------------------------------------
# MAIN APPLICATION
# -------------------------------------------------------------
def main():
    server_thread = threading.Thread(
        target=run_dummy_server,
        daemon=True,
    )
    server_thread.start()

    if not BOT_TOKEN or BOT_TOKEN == "YOUR_TELEGRAM_BOT_TOKEN_HERE":
        print("ERROR: BOT_TOKEN is missing! Set it in your environment variables.")
        return

    app = ApplicationBuilder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", start_command))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(CallbackQueryHandler(handle_callback_query))

    logger.info(f"{BRAND_NAME} is active...")
    app.run_polling()


if __name__ == "__main__":
    main()
