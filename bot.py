import io
import logging
import os
import re
import threading
import time
import unicodedata
from http.server import BaseHTTPRequestHandler, HTTPServer
import fitz  # PyMuPDF
from fpdf import FPDF
from pptx import Presentation
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
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
PORT = int(os.environ.get("PORT", 8080))

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


# -------------------------------------------------------------
# DUMMY SERVER FOR RENDER
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
    logger.info(f"Health check server running on port {PORT}")
    server.serve_forever()


# -------------------------------------------------------------
# TEXT SANITIZER & SYMBOL MAPPER
# -------------------------------------------------------------
def clean_text_for_pdf(text: str) -> str:
    """Cleans symbols, normalizes to ASCII, and breaks overly long tokens."""
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

    # Split long unbroken URLs or code blocks > 32 chars
    words = text.split(' ')
    safe_words = []
    for word in words:
        if len(word) > 32:
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
# INTELLIGENT PDF SLIDE PARSER (Fixes Timelines & Definitions)
# -------------------------------------------------------------
def parse_pdf_slide(page, page_num):
    """Scans bounding boxes, pairs timeline items, unifies paragraphs, and removes slide number noise."""
    page_dict = page.get_text("dict")
    page_height = page.rect.height

    raw_items = []
    for b in page_dict.get("blocks", []):
        if b.get("type") != 0:  # text blocks only
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

            # Remove slide numbers / page stamps at top or bottom corners
            if re.match(r'^\d{1,3}$', line_text) and (y0 > page_height * 0.78 or y0 < page_height * 0.15):
                continue
            if re.match(r'^(page\s*)?\d{1,3}(\s*/\s*\d{1,3})?$', line_text, re.IGNORECASE):
                continue

            raw_items.append({
                "text": line_text,
                "x0": x0,
                "y0": y0,
                "x1": x1,
                "y1": y1,
                "size": max_size
            })

    if not raw_items:
        return {"num": page_num, "title": f"Slide {page_num}", "elements": [], "notes": ""}

    # 1. Identify Title (top banner or largest font size)
    top_candidates = [it for it in raw_items if it["y0"] < page_height * 0.28]
    if top_candidates:
        title_item = max(top_candidates, key=lambda it: (it["size"], -it["y0"]))
    else:
        title_item = max(raw_items, key=lambda it: it["size"])

    title_text = title_item["text"]
    content_items = [it for it in raw_items if it != title_item]

    # 2. Group horizontally aligned boxes (TIMELINE & KEY-VALUE PAIRS like "1950" and "Turing test")
    content_items.sort(key=lambda it: (it["y0"], it["x0"]))
    rows = []
    for it in content_items:
        matched_row = None
        for row in rows:
            avg_y = sum(r["y0"] for r in row) / len(row)
            if abs(it["y0"] - avg_y) <= 14:  # Horizontal alignment tolerance
                matched_row = row
                break
        if matched_row is not None:
            matched_row.append(it)
        else:
            rows.append([it])

    # 3. Categorize into Paragraphs vs Bullets vs Timelines
    formatted_elements = []
    for row in rows:
        row.sort(key=lambda it: it["x0"])
        if len(row) > 1:
            # Timeline / Table Row: "1950 : Turing test"
            row_texts = [r["text"].lstrip("-*> \t") for r in row]
            combined = " : ".join(row_texts)
            formatted_elements.append(("bullet", combined))
        else:
            txt = row[0]["text"]
            if txt.startswith(("-", "*", ">")) or re.match(r'^\d+[\.\)]\s', txt):
                formatted_elements.append(("bullet", txt.lstrip("-*> \t")))
            else:
                formatted_elements.append(("paragraph", txt))

    # 4. Stitch broken sentences into continuous paragraphs (Fixes broken definitions)
    final_elements = []
    for elem_type, elem_text in formatted_elements:
        if elem_type == "paragraph" and final_elements and final_elements[-1][0] == "paragraph":
            prev_text = final_elements[-1][1]
            if not prev_text.endswith((".", ":", "?", "!")):
                final_elements[-1] = ("paragraph", prev_text + " " + elem_text)
            else:
                final_elements.append((elem_type, elem_text))
        else:
            final_elements.append((elem_type, elem_text))

    return {
        "num": page_num,
        "title": title_text,
        "elements": final_elements,
        "notes": ""
    }


async def extract_from_pdf(file_bytes, tracker=None):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    total_pages = len(doc)
    if tracker:
        tracker.total = total_pages

    slides_data = []
    for idx, page in enumerate(doc, start=1):
        if tracker:
            await tracker.update(idx, detail="Analyzing slide structure...")
        slide_info = parse_pdf_slide(page, idx)
        slides_data.append(slide_info)

    if tracker:
        await tracker.update(total_pages, detail="Structure analyzed!", force=True)

    return slides_data


# -------------------------------------------------------------
# INTELLIGENT PPTX PARSER
# -------------------------------------------------------------
async def extract_from_pptx(file_bytes, tracker=None):
    prs = Presentation(io.BytesIO(file_bytes))
    total_slides = len(prs.slides)
    if tracker:
        tracker.total = total_slides

    slides_data = []

    for idx, slide in enumerate(prs.slides, start=1):
        if tracker:
            await tracker.update(idx, detail="Extracting slide content...")

        title = ""
        if slide.shapes.title and slide.shapes.title.text:
            title = clean_text_for_pdf(slide.shapes.title.text)
        else:
            title = f"Topic / Slide {idx}"

        elements = []
        for shape in slide.shapes:
            if shape == slide.shapes.title:
                continue

            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    cleaned = clean_text_for_pdf(p.text)
                    if not cleaned or re.match(r'^\d{1,3}$', cleaned):
                        continue
                    if cleaned.startswith(("-", "*", ">")):
                        elements.append(("bullet", cleaned.lstrip("-*> \t")))
                    else:
                        elements.append(("paragraph", cleaned))

            elif shape.has_table:
                for row in shape.table.rows:
                    row_cells = [clean_text_for_pdf(c.text) for c in row.cells if clean_text_for_pdf(c.text)]
                    if row_cells:
                        elements.append(("bullet", " : ".join(row_cells)))

        # Stitch paragraphs
        stitched_elements = []
        for elem_type, elem_text in elements:
            if elem_type == "paragraph" and stitched_elements and stitched_elements[-1][0] == "paragraph":
                prev_text = stitched_elements[-1][1]
                if not prev_text.endswith((".", ":", "?", "!")):
                    stitched_elements[-1] = ("paragraph", prev_text + " " + elem_text)
                else:
                    stitched_elements.append((elem_type, elem_text))
            else:
                stitched_elements.append((elem_type, elem_text))

        notes = ""
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notes = clean_text_for_pdf(slide.notes_slide.notes_text_frame.text)

        slides_data.append({
            "num": idx,
            "title": title,
            "elements": stitched_elements,
            "notes": notes,
        })

    if tracker:
        await tracker.update(total_slides, detail="Extraction complete!", force=True)

    return slides_data


# -------------------------------------------------------------
# BLACK & WHITE STUDY NOTES GENERATOR
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


def safe_write_paragraph(pdf, text, font_size=10, is_bold=False, is_italic=False, color=(30, 30, 30), prefix=""):
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
    except Exception as e:
        logger.warning(f"Writing fallback: {e}")
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(w=pdf.epw, h=line_height, text=full_text[:80])
        pdf.ln(1)


def create_bw_text_pdf(slides_data):
    pdf = BWNotesPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    for item in slides_data:
        if not item["title"] and not item["elements"] and not item["notes"]:
            continue

        # Clean section divider
        pdf.set_draw_color(220, 220, 220)
        pdf.set_x(pdf.l_margin)
        pdf.line(pdf.l_margin, pdf.get_y(), pdf.l_margin + pdf.epw, pdf.get_y())
        pdf.ln(3)

        # Dynamic Title
        title_len = len(item["title"])
        title_font_size = 9 if title_len > 90 else (10.5 if title_len > 45 else 12)
        safe_write_paragraph(pdf, text=item["title"], font_size=title_font_size, is_bold=True, color=(0, 0, 0), prefix=f"[{item['num']}] ")

        # Elements (Paragraphs & Bullets/Timelines)
        for elem_type, elem_text in item["elements"]:
            if elem_type == "paragraph":
                # Continuous definition / normal text (no bullet prefix)
                safe_write_paragraph(pdf, text=elem_text, font_size=9.5, is_bold=False, color=(20, 20, 20), prefix="   ")
            elif elem_type == "bullet":
                # Timeline or bullet item
                safe_write_paragraph(pdf, text=elem_text, font_size=9.5, is_bold=False, color=(35, 35, 35), prefix="- ")

        # Presenter Notes
        if item["notes"]:
            safe_write_paragraph(pdf, text=item["notes"], font_size=8.5, is_italic=True, color=(80, 80, 80), prefix="Presenter Notes: ")

        pdf.ln(3)

    return bytes(pdf.output())


async def convert_image_pdf_to_bw(file_bytes, tracker=None):
    """Fallback for scanned/image decks with low file size."""
    src_doc = fitz.open(stream=file_bytes, filetype="pdf")
    out_doc = fitz.open()
    total_pages = len(src_doc)
    if tracker:
        tracker.total = total_pages

    for idx, page in enumerate(src_doc, start=1):
        if tracker:
            await tracker.update(idx, detail="Rendering compressed B&W page...")

        pix = page.get_pixmap(colorspace=fitz.csGRAY, dpi=96)
        img_bytes = pix.tobytes("jpeg", jpg_quality=60)
        rect = page.rect
        new_page = out_doc.new_page(width=rect.width, height=rect.height)
        new_page.insert_image(rect, stream=img_bytes)

    if tracker:
        await tracker.update(total_pages, detail="Rendering complete!", force=True)

    return out_doc.tobytes(garbage=4, deflate=True)


# -------------------------------------------------------------
# TELEGRAM BOT HANDLERS
# -------------------------------------------------------------
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome_text = (
        f"👋 **Welcome to {BRAND_NAME}!**\n\n"
        "📄 Send me any **PPTX** or **PDF presentation**, and I will instantly convert it into a "
        "**clean, structured Black & White study notes PDF**!\n\n"
        "✨ **Smart Features:**\n"
        "• Stitches definitions into readable paragraphs 📖\n"
        "• Aligns timelines & key-value pairs (`Year : Event`) ⏳\n"
        "• Removes slide numbers & background noise 🧼\n"
        "• Live real-time conversion progress bar 📊\n\n"
        "🚀 **Send your file to start!**"
    )
    await update.message.reply_text(welcome_text, parse_mode="Markdown")


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    file_name = document.file_name or "presentation"
    file_ext = file_name.split(".")[-1].lower()

    if file_ext not in ["pptx", "pdf"]:
        await update.message.reply_text("⚠️ Please send a valid **.pptx** or **.pdf** slide file.", parse_mode="Markdown")
        return

    status_msg = await update.message.reply_text(
        f"⚡ **{BRAND_NAME} is preparing...**\n\n"
        "`[░░░░░░░░░░] 0%`\n"
        "📥 *Downloading your file...*",
        parse_mode="Markdown"
    )

    try:
        tg_file = await context.bot.get_file(document.file_id)
        raw_bytes = await tg_file.download_as_bytearray()
        file_bytes = bytes(raw_bytes)

        output_pdf_bytes = None
        tracker = ProgressTracker(status_msg=status_msg, total_items=10, stage_name="Processing Slides")

        if file_ext == "pptx":
            slides_data = await extract_from_pptx(file_bytes, tracker=tracker)
            output_pdf_bytes = create_bw_text_pdf(slides_data)
        else:
            slides_data = await extract_from_pdf(file_bytes, tracker=tracker)
            total_elements = sum(len(s["elements"]) for s in slides_data)

            if total_elements > 0:
                try:
                    output_pdf_bytes = create_bw_text_pdf(slides_data)
                except Exception as text_err:
                    logger.warning(f"Text layout error: {text_err}. Switching to visual fallback.")
                    tracker.stage_name = "Rendering Grayscale Slides"
                    output_pdf_bytes = await convert_image_pdf_to_bw(file_bytes, tracker=tracker)
            else:
                tracker.stage_name = "Rendering Grayscale Slides"
                output_pdf_bytes = await convert_image_pdf_to_bw(file_bytes, tracker=tracker)

        try:
            await status_msg.edit_text(
                f"⚡ **{BRAND_NAME}**\n\n"
                "`[██████████] 100%`\n"
                "📤 **Finalizing and sending your study notes...**",
                parse_mode="Markdown"
            )
        except Exception:
            pass

        base_name = os.path.splitext(file_name)[0]
        output_filename = f"{base_name}_Yash_PPT_Converter_Bot.pdf"

        await update.message.reply_document(
            document=io.BytesIO(output_pdf_bytes),
            filename=output_filename,
            caption=f"✅ **Converted successfully by {BRAND_NAME}!**\n📄 Clean & Structured Black & White Study Notes.",
            parse_mode="Markdown",
        )
        await status_msg.delete()

    except Exception as e:
        logger.error(f"Error during conversion: {e}", exc_info=True)
        await status_msg.edit_text(f"❌ Conversion failed: `{str(e)}`", parse_mode="Markdown")


# -------------------------------------------------------------
# MAIN ENTRY
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

    logger.info(f"{BRAND_NAME} is active and running...")
    app.run_polling()


if __name__ == "__main__":
    main()
