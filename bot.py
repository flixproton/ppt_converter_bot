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
# DUMMY SERVER (Keeps Render Web Service alive)
# -------------------------------------------------------------
class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/plain")
        self.end_headers()
        self.wfile.write(f"{BRAND_NAME} is active and running!".encode("utf-8"))

    def log_message(self, format, *args):
        return


def run_dummy_server():
    server = HTTPServer(("0.0.0.0", PORT), HealthCheckHandler)
    logger.info(f"Dummy health check server running on port {PORT}")
    server.serve_forever()


# -------------------------------------------------------------
# PROGRESS BAR HELPER (Throttled to avoid Telegram Rate Limits)
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
        # Update at most once every 1.5 seconds unless forced to avoid Telegram FloodWait
        if not force and (now - self.last_update_time < 1.5):
            return

        self.last_update_time = now
        bar_text = self._generate_bar(current)
        remaining = max(0, self.total - current)

        msg = (
            f"⚡ **{BRAND_NAME} is working...**\n\n"
            f"`{bar_text}`\n\n"
            f"📊 **Stage:** {self.stage_name}\n"
            f"📄 **Progress:** Slide `{current}` of `{self.total}`\n"
            f"⏳ **Remaining:** `{remaining}` slide(s)\n"
            f"ℹ️ _{detail}_"
        )
        try:
            await self.status_msg.edit_text(msg, parse_mode="Markdown")
        except Exception:
            # Ignore duplicate message edits or network blips silently
            pass


# -------------------------------------------------------------
# TEXT CLEANER & SANITIZER
# -------------------------------------------------------------
def clean_text_for_pdf(text: str) -> str:
    """Sanitizes unicode, symbols, and breaks long words that exceed column widths."""
    if not text:
        return ""

    replacements = {
        '“': '"', '”': '"', '‘': "'", '’': "'",
        '—': '-', '–': '-', '…': '...', '•': '*',
        '▪': '*', '►': '>', '✔': '/', '✓': '/',
        '→': '->', '←': '<-', '⇒': '=>', '≤': '<=', '≥': '>=',
        '≠': '!=', '±': '+/-', '×': 'x', '÷': '/', '°': ' deg ',
        '\u200b': '', '\ufeff': '', '\xa0': ' ', '\r': ''
    }
    for k, v in replacements.items():
        text = text.replace(k, v)

    text = unicodedata.normalize('NFKD', text)
    text = text.encode('latin-1', 'ignore').decode('latin-1')
    text = "".join(ch for ch in text if ch.isprintable() or ch in ['\n', '\t', ' '])

    # Split continuous words/URLs longer than 32 chars
    words = text.split(' ')
    safe_words = []
    for word in words:
        if len(word) > 32:
            chunks = [word[i:i+30] for i in range(0, len(word), 30)]
            safe_words.append(" ".join(chunks))
        else:
            safe_words.append(word)

    return " ".join(safe_words).strip()


# -------------------------------------------------------------
# EXTRACTORS WITH PROGRESS HOOKS
# -------------------------------------------------------------
async def extract_from_pptx(file_bytes, tracker=None):
    prs = Presentation(io.BytesIO(file_bytes))
    total_slides = len(prs.slides)
    if tracker:
        tracker.total = total_slides

    slides_data = []

    for idx, slide in enumerate(prs.slides, start=1):
        if tracker:
            await tracker.update(idx, detail="Extracting slide text & notes...")

        title = ""
        if slide.shapes.title and slide.shapes.title.text:
            title = clean_text_for_pdf(slide.shapes.title.text)
        else:
            title = f"Topic / Slide {idx}"

        content = []
        for shape in slide.shapes:
            if shape == slide.shapes.title:
                continue
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    cleaned = clean_text_for_pdf(p.text)
                    if cleaned and cleaned not in content:
                        content.append(cleaned)
            elif shape.has_table:
                for row in shape.table.rows:
                    row_cells = [clean_text_for_pdf(c.text) for c in row.cells if clean_text_for_pdf(c.text)]
                    if row_cells:
                        content.append(" | ".join(row_cells))

        notes = ""
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notes = clean_text_for_pdf(slide.notes_slide.notes_text_frame.text)

        slides_data.append({
            "num": idx,
            "title": title,
            "content": content,
            "notes": notes,
        })

    if tracker:
        await tracker.update(total_slides, detail="Text extraction complete!", force=True)

    return slides_data


async def extract_from_pdf(file_bytes, tracker=None):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    total_pages = len(doc)
    if tracker:
        tracker.total = total_pages

    slides_data = []

    for idx, page in enumerate(doc, start=1):
        if tracker:
            await tracker.update(idx, detail="Reading slide page...")

        text = page.get_text("text").strip()
        raw_lines = [clean_text_for_pdf(line) for line in text.split("\n") if clean_text_for_pdf(line)]

        title = raw_lines[0] if raw_lines else f"Slide {idx}"
        content = raw_lines[1:] if len(raw_lines) > 1 else []

        slides_data.append({
            "num": idx,
            "title": title,
            "content": content,
            "notes": "",
        })

    if tracker:
        await tracker.update(total_pages, detail="Slide reading complete!", force=True)

    return slides_data


# -------------------------------------------------------------
# BLACK & WHITE PDF GENERATOR
# -------------------------------------------------------------
class BWNotesPDF(FPDF):
    def header(self):
        self.set_x(self.l_margin)
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(100, 100, 100)
        self.cell(self.epw, 6, f"Study Notes • {BRAND_NAME}", border=0, align="R")
        self.ln(8)

    def footer(self):
        self.set_y(-12)
        self.set_x(self.l_margin)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(120, 120, 120)
        self.cell(self.epw, 8, f"Page {self.page_no()} | Generated by {BRAND_NAME}", align="C")


def safe_write_paragraph(pdf, text, font_size=10, is_bold=False, is_italic=False, color=(30, 30, 30), prefix=""):
    if not text:
        return

    style = ""
    if is_bold and is_italic:
        style = "BI"
    elif is_bold:
        style = "B"
    elif is_italic:
        style = "I"

    pdf.set_font("Helvetica", style, font_size)
    pdf.set_text_color(*color)
    pdf.set_x(pdf.l_margin)

    full_text = f"{prefix}{text}"
    line_height = max(4.5, font_size * 0.45)

    try:
        pdf.multi_cell(w=pdf.epw, h=line_height, text=full_text)
        pdf.ln(1)
    except Exception:
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(w=pdf.epw, h=line_height, text=full_text[:120] + "...")
        pdf.ln(1)


def create_bw_text_pdf(slides_data):
    pdf = BWNotesPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    for item in slides_data:
        if not item["title"] and not item["content"] and not item["notes"]:
            continue

        pdf.set_draw_color(220, 220, 220)
        pdf.set_x(pdf.l_margin)
        pdf.line(pdf.l_margin, pdf.get_y(), pdf.l_margin + pdf.epw, pdf.get_y())
        pdf.ln(3)

        title_len = len(item["title"])
        if title_len > 90:
            title_font_size = 9
        elif title_len > 45:
            title_font_size = 10.5
        else:
            title_font_size = 12

        safe_write_paragraph(pdf, text=item["title"], font_size=title_font_size, is_bold=True, color=(0, 0, 0), prefix=f"[{item['num']}] ")

        if item["content"]:
            for bullet in item["content"]:
                safe_write_paragraph(pdf, text=bullet, font_size=9.5, is_bold=False, color=(35, 35, 35), prefix="• ")

        if item["notes"]:
            safe_write_paragraph(pdf, text=item["notes"], font_size=8.5, is_italic=True, color=(80, 80, 80), prefix="Presenter Notes: ")

        pdf.ln(3)

    return bytes(pdf.output())


async def convert_image_pdf_to_bw(file_bytes, tracker=None):
    """Fast Grayscale rendering for image/scanned slide decks."""
    src_doc = fitz.open(stream=file_bytes, filetype="pdf")
    out_doc = fitz.open()
    total_pages = len(src_doc)
    if tracker:
        tracker.total = total_pages

    for idx, page in enumerate(src_doc, start=1):
        if tracker:
            await tracker.update(idx, detail="Rendering clean B&W page...")

        # 120 DPI gives crisp reading quality with 40% faster rendering speed
        pix = page.get_pixmap(colorspace=fitz.csGRAY, dpi=120)
        img_bytes = pix.tobytes("png")
        rect = page.rect
        new_page = out_doc.new_page(width=rect.width, height=rect.height)
        new_page.insert_image(rect, stream=img_bytes)

    if tracker:
        await tracker.update(total_pages, detail="Rendering finished!", force=True)

    return out_doc.tobytes()


# -------------------------------------------------------------
# TELEGRAM BOT HANDLERS
# -------------------------------------------------------------
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome_text = (
        f"👋 **Welcome to {BRAND_NAME}!**\n\n"
        "📄 Send me any **PPTX** or **PDF presentation**, and I will instantly convert it into a "
        "**clean, ink-saving Black & White study notes PDF**!\n\n"
        "✨ **Features:**\n"
        "• Real-time conversion Progress Bar 📊\n"
        "• Strips dark background colors & images\n"
        "• Automatically fits headlines & bullet points\n"
        "• Extracts slide text, tables & presenter notes\n\n"
        "🚀 **Send your file now to get started!**"
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
        # Download document into memory
        tg_file = await context.bot.get_file(document.file_id)
        raw_bytes = await tg_file.download_as_bytearray()
        file_bytes = bytes(raw_bytes)

        output_pdf_bytes = None
        tracker = ProgressTracker(status_msg=status_msg, total_items=10, stage_name="Converting Slides")

        if file_ext == "pptx":
            slides_data = await extract_from_pptx(file_bytes, tracker=tracker)
            output_pdf_bytes = create_bw_text_pdf(slides_data)
        else:
            slides_data = await extract_from_pdf(file_bytes, tracker=tracker)
            total_text_len = sum(len(s["title"]) + sum(len(c) for c in s["content"]) for s in slides_data)

            if total_text_len > 30:
                try:
                    output_pdf_bytes = create_bw_text_pdf(slides_data)
                except Exception as text_err:
                    logger.warning(f"Text layout failed, switching to image B&W fallback: {text_err}")
                    tracker.stage_name = "Rendering Grayscale Slides"
                    output_pdf_bytes = await convert_image_pdf_to_bw(file_bytes, tracker=tracker)
            else:
                tracker.stage_name = "Rendering Grayscale Slides"
                output_pdf_bytes = await convert_image_pdf_to_bw(file_bytes, tracker=tracker)

        # Notify upload
        try:
            await status_msg.edit_text(
                f"⚡ **{BRAND_NAME}**\n\n"
                "`[██████████] 100%`\n"
                "📤 **Finalizing and sending your notes...**",
                parse_mode="Markdown"
            )
        except Exception:
            pass

        # Brand output filename
        base_name = os.path.splitext(file_name)[0]
        output_filename = f"{base_name}_Yash_PPT_Converter_Bot.pdf"

        # Send PDF back to user
        await update.message.reply_document(
            document=io.BytesIO(output_pdf_bytes),
            filename=output_filename,
            caption=f"✅ **Converted successfully by {BRAND_NAME}!**\n📄 Clean Black & White Study Notes.",
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

    logger.info(f"{BRAND_NAME} is active and listening...")
    app.run_polling()


if __name__ == "__main__":
    main()
