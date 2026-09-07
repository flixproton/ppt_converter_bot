import io
import json
import logging
import os
import re
import threading
import time
import unicodedata
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
# TEXT SANITIZER
# -------------------------------------------------------------
def clean_text_for_pdf(text: str) -> str:
    if not text:
        return ""

    replacements = {
        '“': '"', '”': '"', '‘': "'", '’': "'",
        '—': '-', '–': '-', '…': '...', '•': '-',
        '▪': '-', '►': '>', '✔': '/', '✓': '/',
        '→': '->', '←': '<-', '⇒': '=>', '≤': '<=', '≥': '>=',
        '≠': '!=', '±': '+/-', '×': 'x', '÷': '/', '°': ' deg ',
        '\u200b': '', '\ufeff': '', '\xa0': ' ', '\r': ''
    }
    for k, v in replacements.items():
        text = text.replace(k, v)

    text = unicodedata.normalize('NFKD', text)
    text = text.encode('ascii', 'ignore').decode('ascii')
    text = "".join(ch for ch in text if ch.isprintable() or ch in ['\n', '\t', ' '])

    words = text.split(' ')
    safe_words = []
    for word in words:
        if len(word) > 30:
            chunks = [word[i:i+28] for i in range(0, len(word), 28)]
            safe_words.append(" ".join(chunks))
        else:
            safe_words.append(word)

    return " ".join(safe_words).strip()


# -------------------------------------------------------------
# PROGRESS TRACKER
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
# OPTION 1: DIRECT B&W SLIDES CONVERSION (NO LAYOUT CHANGES)
# -------------------------------------------------------------
async def convert_exact_slides_to_bw(file_bytes, file_ext, tracker=None):
    """Converts original presentation into clean grayscale/B&W without changing text or layouts."""
    if file_ext == "pdf":
        src_doc = fitz.open(stream=file_bytes, filetype="pdf")
        out_doc = fitz.open()
        total_pages = len(src_doc)
        if tracker:
            tracker.total = total_pages

        for idx, page in enumerate(src_doc, start=1):
            if tracker:
                await tracker.update(idx, detail="Converting slide to Black & White...")

            # Grayscale 110 DPI with JPEG compression for ink saving
            pix = page.get_pixmap(colorspace=fitz.csGRAY, dpi=110)
            img_bytes = pix.tobytes("jpeg", jpg_quality=75)
            rect = page.rect
            new_page = out_doc.new_page(width=rect.width, height=rect.height)
            new_page.insert_image(rect, stream=img_bytes)

        if tracker:
            await tracker.update(total_pages, detail="B&W conversion finished!", force=True)

        return out_doc.tobytes(garbage=4, deflate=True)

    else:
        # For PPTX: Build clean visual slide-by-slide B&W layout
        prs = Presentation(io.BytesIO(file_bytes))
        pdf = FPDF(orientation="L", unit="mm", format="A4")
        pdf.set_auto_page_break(auto=True, margin=10)
        total_slides = len(prs.slides)
        if tracker:
            tracker.total = total_slides

        for idx, slide in enumerate(prs.slides, start=1):
            if tracker:
                await tracker.update(idx, detail="Processing slide layout...")
            pdf.add_page()

            # Border
            pdf.set_draw_color(180, 180, 180)
            pdf.rect(10, 10, 277, 190)

            # Title
            title = ""
            if slide.shapes.title and slide.shapes.title.text:
                title = clean_text_for_pdf(slide.shapes.title.text)
                pdf.set_xy(15, 15)
                pdf.set_font("Helvetica", "B", 16)
                pdf.set_text_color(0, 0, 0)
                pdf.multi_cell(267, 8, title)
                pdf.ln(5)

            # Body text
            pdf.set_font("Helvetica", "", 12)
            pdf.set_text_color(40, 40, 40)
            for shape in slide.shapes:
                if shape != slide.shapes.title and shape.has_text_frame:
                    for p in shape.text_frame.paragraphs:
                        text = clean_text_for_pdf(p.text)
                        if text:
                            pdf.set_x(15)
                            pdf.multi_cell(267, 6, f"- {text}")
                            pdf.ln(1)

        return bytes(pdf.output())


# -------------------------------------------------------------
# OPTION 2: SMART AI STUDY NOTES ENGINE
# -------------------------------------------------------------
def analyze_and_structure_with_ai(raw_slides_data):
    if not GEMINI_API_KEY:
        return None

    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
        slides_summary = [{"slide_num": s["num"], "raw_text": s.get("raw_text", ""), "is_diagram_candidate": s.get("is_diagram", False)} for s in raw_slides_data]

        prompt = f"""
You are an expert academic notes generator. Convert these raw presentation slides into clean study notes.
Rules:
1. DEFINITIONS: Stitch split lines into complete continuous paragraphs.
2. TIMELINES: Format each year/date paired with its event (e.g. "- **1950:** Turing test", "- **1955:** Dartmouth Conference").
3. COMPARISONS / CATEGORIES: Group items under clear bold subheadings with clean bullet points.
4. DIAGRAMS / FLOWCHARTS: If a slide is primarily a flowchart, Venn diagram, or image infographic with almost no text, set `"is_diagram": true` and provide a 1-sentence `"caption"`.
5. NOISE REMOVAL: Discard isolated slide numbers (like "6", "7", "8").

Raw Slides Input:
{json.dumps(slides_summary)}

Return a strict JSON array of objects with schema:
[
  {{
    "slide_num": 1,
    "title": "Clean Slide Title",
    "is_diagram": false,
    "caption": "",
    "elements": [
       {{"type": "paragraph", "text": "Continuous definition..."}},
       {{"type": "bullet", "text": "**1950:** Turing test"}},
       {{"type": "subheading", "text": "WEAK AI"}}
    ]
  }}
]
Return ONLY raw JSON. No markdown code blocks.
"""
        response = model.generate_content(prompt)
        text_resp = response.text.strip()
        if text_resp.startswith("```"):
            text_resp = re.sub(r"^```(?:json)?\n?", "", text_resp)
            text_resp = re.sub(r"\n?```$", "", text_resp)

        return json.loads(text_resp)
    except Exception as e:
        logger.warning(f"Gemini structuring skipped: {e}")
        return None


def parse_pdf_slide_local(page, page_num):
    page_dict = page.get_text("dict")
    page_height = page.rect.height

    raw_items = []
    total_text = ""
    for b in page_dict.get("blocks", []):
        if b.get("type") != 0:
            continue
        for l in b.get("lines", []):
            line_text = ""
            max_size = 0
            x0, y0, x1, y1 = l.get("bbox", (0, 0, 0, 0))
            for s in l.get("spans", []):
                t = clean_text_for_pdf(s.get("text", ""))
                if t:
                    line_text += (" " if line_text else "") + t
                    max_size = max(max_size, s.get("size", 10))

            if not line_text:
                continue

            if re.match(r'^\d{1,3}$', line_text) and (y0 > page_height * 0.78 or y0 < page_height * 0.15):
                continue
            if re.match(r'^(page\s*)?\d{1,3}(\s*/\s*\d{1,3})?$', line_text, re.IGNORECASE):
                continue

            raw_items.append({"text": line_text, "x0": x0, "y0": y0, "size": max_size})
            total_text += " " + line_text

    is_diagram = len(total_text.strip()) < 35

    if not raw_items:
        return {
            "num": page_num,
            "title": f"Slide {page_num}",
            "is_diagram": True,
            "raw_text": "",
            "elements": [{"type": "paragraph", "text": "Visual Diagram / Flowchart"}]
        }

    top_candidates = [it for it in raw_items if it["y0"] < page_height * 0.28]
    title_item = max(top_candidates, key=lambda it: (it["size"], -it["y0"])) if top_candidates else max(raw_items, key=lambda it: it["size"])
    title_text = title_item["text"]
    content_items = [it for it in raw_items if it != title_item]

    content_items.sort(key=lambda it: (it["y0"], it["x0"]))
    rows = []
    for it in content_items:
        matched_row = None
        for row in rows:
            avg_y = sum(r["y0"] for r in row) / len(row)
            if abs(it["y0"] - avg_y) <= 14:
                matched_row = row
                break
        if matched_row is not None:
            matched_row.append(it)
        else:
            rows.append([it])

    elements = []
    for row in rows:
        row.sort(key=lambda it: it["x0"])
        if len(row) > 1:
            row_texts = [r["text"].lstrip("-*> \t") for r in row]
            elements.append({"type": "bullet", "text": f"{row_texts[0]} : " + " : ".join(row_texts[1:])})
        else:
            txt = row[0]["text"]
            if txt.startswith(("-", "*", ">")) or re.match(r'^\d+[\.\)]\s', txt):
                elements.append({"type": "bullet", "text": txt.lstrip("-*> \t")})
            else:
                elements.append({"type": "paragraph", "text": txt})

    final_elements = []
    for elem in elements:
        if elem["type"] == "paragraph" and final_elements and final_elements[-1]["type"] == "paragraph":
            prev = final_elements[-1]["text"]
            if not prev.endswith((".", ":", "?", "!")):
                final_elements[-1]["text"] = prev + " " + elem["text"]
            else:
                final_elements.append(elem)
        else:
            final_elements.append(elem)

    return {
        "num": page_num,
        "title": title_text,
        "is_diagram": is_diagram,
        "raw_text": total_text,
        "elements": final_elements
    }


# -------------------------------------------------------------
# A4 STUDY NOTES PDF BUILDER
# -------------------------------------------------------------
class BWNotesPDF(FPDF):
    def header(self):
        self.set_x(self.l_margin)
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(100, 100, 100)
        self.cell(self.epw, 6, clean_text_for_pdf(f"Study Notes | {BRAND_NAME}"), border=0, align="R")
        self.ln(8)

    def footer(self):
        self.set_y(-12)
        self.set_x(self.l_margin)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(120, 120, 120)
        self.cell(self.epw, 8, clean_text_for_pdf(f"Page {self.page_no()} | Generated by {BRAND_NAME}"), align="C")


def safe_write_text(pdf, text, font_size=10, is_bold=False, is_italic=False, color=(30, 30, 30), prefix=""):
    if not text:
        return

    style = "B" if is_bold else ("I" if is_italic else "")
    pdf.set_font("Helvetica", style, font_size)
    pdf.set_text_color(*color)
    pdf.set_x(pdf.l_margin)

    full_text = clean_text_for_pdf(f"{prefix}{text}")
    line_height = max(4.5, font_size * 0.48)

    try:
        pdf.multi_cell(w=pdf.epw, h=line_height, text=full_text)
        pdf.ln(1)
    except Exception:
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(w=pdf.epw, h=line_height, text=full_text[:80])
        pdf.ln(1)


def build_final_notes_pdf(slides_data, src_pdf_doc=None):
    pdf = BWNotesPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    for item in slides_data:
        slide_num = item.get("slide_num", item.get("num", 1))
        title = item.get("title", f"Slide {slide_num}")
        is_diagram = item.get("is_diagram", False)
        elements = item.get("elements", [])
        caption = item.get("caption", "")

        pdf.set_draw_color(220, 220, 220)
        pdf.set_x(pdf.l_margin)
        pdf.line(pdf.l_margin, pdf.get_y(), pdf.l_margin + pdf.epw, pdf.get_y())
        pdf.ln(3)

        title_len = len(title)
        title_size = 9 if title_len > 90 else (10.5 if title_len > 45 else 12)
        safe_write_text(pdf, text=title, font_size=title_size, is_bold=True, color=(0, 0, 0), prefix=f"[{slide_num}] ")

        if is_diagram and src_pdf_doc and (slide_num - 1) < len(src_pdf_doc):
            try:
                page = src_pdf_doc[slide_num - 1]
                pix = page.get_pixmap(colorspace=fitz.csGRAY, dpi=110)
                img_data = pix.tobytes("jpeg", jpg_quality=75)

                img_w = min(170, pdf.epw)
                img_h = (img_w / page.rect.width) * page.rect.height
                if pdf.get_y() + img_h > 270:
                    pdf.add_page()

                img_stream = io.BytesIO(img_data)
                pdf.image(img_stream, x=pdf.l_margin, y=pdf.get_y(), w=img_w)
                pdf.set_y(pdf.get_y() + img_h + 3)

                if caption:
                    safe_write_text(pdf, text=f"Diagram: {caption}", font_size=8.5, is_italic=True, color=(90, 90, 90))
            except Exception as e:
                logger.warning(f"Diagram embed error: {e}")

        for elem in elements:
            e_type = elem.get("type", "paragraph")
            e_text = elem.get("text", "")
            if e_type == "subheading":
                pdf.ln(1)
                safe_write_text(pdf, text=e_text, font_size=10, is_bold=True, color=(20, 20, 20))
            elif e_type == "paragraph":
                safe_write_text(pdf, text=e_text, font_size=9.5, is_bold=False, color=(20, 20, 20), prefix="   ")
            elif e_type == "bullet":
                safe_write_text(pdf, text=e_text, font_size=9.5, is_bold=False, color=(35, 35, 35), prefix="- ")

        pdf.ln(3)

    return bytes(pdf.output())


# -------------------------------------------------------------
# TELEGRAM BOT HANDLERS & CALLBACKS
# -------------------------------------------------------------
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome_text = (
        f"👋 **Welcome to {BRAND_NAME}!**\n\n"
        "📄 Send me any **PPTX** or **PDF presentation**, and choose your conversion mode:\n\n"
        "🖨️ **Option 1: Direct B&W Slides**\n"
        "• Converts your exact presentation into ink-saving Black & White (No layout changes)\n\n"
        "📝 **Option 2: AI Study Notes**\n"
        "• Scans and structures definitions, timelines & preserves diagrams as B&W figures\n\n"
        "🚀 **Send your presentation file now!**"
    )
    await update.message.reply_text(welcome_text, parse_mode="Markdown")


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    file_name = document.file_name or "presentation"
    file_ext = file_name.split(".")[-1].lower()

    if file_ext not in ["pptx", "pdf"]:
        await update.message.reply_text("⚠️ Please upload a valid **.pptx** or **.pdf** slide file.", parse_mode="Markdown")
        return

    # Download into memory and store in user_data
    downloading_msg = await update.message.reply_text("📥 *Downloading file...*", parse_mode="Markdown")
    tg_file = await context.bot.get_file(document.file_id)
    raw_bytes = await tg_file.download_as_bytearray()
    file_bytes = bytes(raw_bytes)

    context.user_data["pending_file"] = {
        "bytes": file_bytes,
        "name": file_name,
        "ext": file_ext
    }

    # Present Interactive Choice Menu
    keyboard = [
        [InlineKeyboardButton("🖨️ 1. Direct B&W Slides (Exact Copy)", callback_data="mode_exact_bw")],
        [InlineKeyboardButton("📝 2. AI Study Notes (Full Restructure)", callback_data="mode_ai_notes")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)

    await downloading_msg.edit_text(
        f"✅ **File received:** `{file_name}`\n\n"
        "👉 **Please choose your conversion mode:**",
        reply_markup=reply_markup,
        parse_mode="Markdown"
    )


async def handle_callback_query(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    mode = query.data
    file_info = context.user_data.get("pending_file")

    if not file_info:
        await query.edit_message_text("⚠️ File session expired. Please re-upload your presentation.", parse_mode="Markdown")
        return

    file_bytes = file_info["bytes"]
    file_name = file_info["name"]
    file_ext = file_info["ext"]
    base_name = os.path.splitext(file_name)[0]

    status_msg = await query.edit_message_text(f"⚡ **{BRAND_NAME} is starting...**\n`[░░░░░░░░░░] 0%`", parse_mode="Markdown")

    try:
        tracker = ProgressTracker(status_msg=status_msg, total_items=10, stage_name="Processing")

        if mode == "mode_exact_bw":
            # ---------------- MODE 1: DIRECT B&W SLIDES ----------------
            tracker.stage_name = "Generating B&W Slides"
            output_pdf_bytes = await convert_exact_slides_to_bw(file_bytes, file_ext, tracker=tracker)
            output_filename = f"{base_name}_BW_Slides_{BRAND_NAME.replace(' ', '_')}.pdf"
            caption_text = f"✅ **Direct Black & White Slides by {BRAND_NAME}!**\n🖨️ Exact layout preserved for ink-saving printing."

        else:
            # ---------------- MODE 2: AI STUDY NOTES ----------------
            tracker.stage_name = "Scanning & Structuring"
            pdf_doc = fitz.open(stream=file_bytes, filetype="pdf") if file_ext == "pdf" else None

            raw_slides = []
            if pdf_doc:
                tracker.total = len(pdf_doc)
                for idx, page in enumerate(pdf_doc, start=1):
                    await tracker.update(idx, detail="Scanning text & diagrams...")
                    raw_slides.append(parse_pdf_slide_local(page, idx))
            else:
                prs = Presentation(io.BytesIO(file_bytes))
                tracker.total = len(prs.slides)
                for idx, slide in enumerate(prs.slides, start=1):
                    await tracker.update(idx, detail="Scanning slide content...")
                    title = slide.shapes.title.text.strip() if slide.shapes.title and slide.shapes.title.text else f"Slide {idx}"
                    text_content = []
                    for s in slide.shapes:
                        if s.has_text_frame and s != slide.shapes.title:
                            for p in s.text_frame.paragraphs:
                                if p.text.strip():
                                    text_content.append(p.text.strip())
                    raw_slides.append({
                        "num": idx,
                        "title": title,
                        "is_diagram": len(text_content) < 2,
                        "raw_text": " ".join(text_content),
                        "elements": [{"type": "paragraph", "text": t} for t in text_content]
                    })

            await tracker.update(tracker.total, detail="AI organizing notes & timelines...", force=True)
            structured_slides = analyze_and_structure_with_ai(raw_slides) if GEMINI_API_KEY else None
            final_data = structured_slides if structured_slides else raw_slides

            await tracker.update(tracker.total, detail="Building study notes PDF...", force=True)
            output_pdf_bytes = build_final_notes_pdf(final_data, src_pdf_doc=pdf_doc)
            output_filename = f"{base_name}_Study_Notes_{BRAND_NAME.replace(' ', '_')}.pdf"
            caption_text = f"✅ **AI Study Notes by {BRAND_NAME}!**\n📝 Clean definitions, aligned timelines & diagram figures."

        # Send File
        await query.message.reply_document(
            document=io.BytesIO(output_pdf_bytes),
            filename=output_filename,
            caption=caption_text,
            parse_mode="Markdown"
        )
        await status_msg.delete()

    except Exception as e:
        logger.error(f"Processing error: {e}", exc_info=True)
        await status_msg.edit_text(f"❌ Conversion failed: `{str(e)}`", parse_mode="Markdown")

    finally:
        # Clear file from memory
        context.user_data.pop("pending_file", None)


# -------------------------------------------------------------
# MAIN APPLICATION
# -------------------------------------------------------------
def main():
    server_thread = threading.Thread(target=run_dummy_server, daemon=True)
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
