import io
import logging
import os
import re
import threading
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
# DUMMY SERVER FOR RENDER WEB SERVICE
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
# ROBUST TEXT SANITIZER & LONG WORD SPLITTER
# -------------------------------------------------------------
def clean_text_for_pdf(text: str) -> str:
    """Sanitizes text, converts math/unicode symbols, and splits overlong words."""
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

    # Convert unicode to standard ASCII
    text = unicodedata.normalize('NFKD', text)
    text = text.encode('latin-1', 'ignore').decode('latin-1')

    # Remove non-printable control characters
    text = "".join(ch for ch in text if ch.isprintable() or ch in ['\n', '\t', ' '])

    # CRITICAL FIX: Split long unbroken tokens (URLs/code/math) exceeding 32 chars
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
# EXTRACTION LOGIC (PPTX & PDF)
# -------------------------------------------------------------
def extract_from_pptx(file_bytes):
    """Extracts text from slides, tables, shapes, and notes in PPTX."""
    prs = Presentation(io.BytesIO(file_bytes))
    slides_data = []

    for idx, slide in enumerate(prs.slides, start=1):
        title = ""
        if slide.shapes.title and slide.shapes.title.text:
            title = clean_text_for_pdf(slide.shapes.title.text)
        else:
            title = f"Topic / Slide {idx}"

        content = []
        for shape in slide.shapes:
            if shape == slide.shapes.title:
                continue

            # Standard Text Frames
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    cleaned = clean_text_for_pdf(p.text)
                    if cleaned and cleaned not in content:
                        content.append(cleaned)

            # Tables in Slides
            elif shape.has_table:
                for row in shape.table.rows:
                    row_cells = [clean_text_for_pdf(c.text) for c in row.cells if clean_text_for_pdf(c.text)]
                    if row_cells:
                        content.append(" | ".join(row_cells))

        # Presenter Notes
        notes = ""
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notes = clean_text_for_pdf(slide.notes_slide.notes_text_frame.text)

        slides_data.append({
            "num": idx,
            "title": title,
            "content": content,
            "notes": notes,
        })
    return slides_data


def extract_from_pdf(file_bytes):
    """Extracts text blocks cleanly from PDF presentations."""
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    slides_data = []

    for idx, page in enumerate(doc, start=1):
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
    return slides_data


# -------------------------------------------------------------
# SAFE BLACK & WHITE PDF GENERATOR
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
    """Safely writes multi-line text with auto-managed widths and margins."""
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
    except Exception as e:
        logger.warning(f"Fallback writing line due to error: {e}")
        # Secondary fallback: break characters if still oversized
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(w=pdf.epw, h=line_height, text=full_text[:120] + "...")
        pdf.ln(1)


def create_bw_text_pdf(slides_data):
    """Generates structured A4 Black and White study notes with dynamic font scaling."""
    pdf = BWNotesPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    for item in slides_data:
        if not item["title"] and not item["content"] and not item["notes"]:
            continue

        # Section separator line
        pdf.set_draw_color(220, 220, 220)
        pdf.set_x(pdf.l_margin)
        pdf.line(pdf.l_margin, pdf.get_y(), pdf.l_margin + pdf.epw, pdf.get_y())
        pdf.ln(3)

        # Dynamic Title Scaling (fits any title size automatically)
        title_len = len(item["title"])
        if title_len > 90:
            title_font_size = 9
        elif title_len > 45:
            title_font_size = 10.5
        else:
            title_font_size = 12

        safe_write_paragraph(
            pdf,
            text=item["title"],
            font_size=title_font_size,
            is_bold=True,
            color=(0, 0, 0),
            prefix=f"[{item['num']}] "
        )

        # Content Bullets
        if item["content"]:
            for bullet in item["content"]:
                safe_write_paragraph(
                    pdf,
                    text=bullet,
                    font_size=9.5,
                    is_bold=False,
                    color=(35, 35, 35),
                    prefix="• "
                )

        # Speaker Notes (if any)
        if item["notes"]:
            safe_write_paragraph(
                pdf,
                text=item["notes"],
                font_size=8.5,
                is_italic=True,
                color=(80, 80, 80),
                prefix="Presenter Notes: "
            )

        pdf.ln(3)

    return bytes(pdf.output())


def convert_image_pdf_to_bw(file_bytes):
    """Universal Fallback: Converts full slide deck pages into crisp B&W/grayscale printable pages."""
    src_doc = fitz.open(stream=file_bytes, filetype="pdf")
    out_doc = fitz.open()

    for page in src_doc:
        pix = page.get_pixmap(colorspace=fitz.csGRAY, dpi=150)
        img_bytes = pix.tobytes("png")
        rect = page.rect
        new_page = out_doc.new_page(width=rect.width, height=rect.height)
        new_page.insert_image(rect, stream=img_bytes)

    return out_doc.tobytes()


# -------------------------------------------------------------
# TELEGRAM BOT HANDLERS
# -------------------------------------------------------------
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome_text = (
        f"👋 **Welcome to {BRAND_NAME}!**\n\n"
        "📄 Send me any **PPTX** or **PDF presentation**, and I will instantly convert it into a "
        "**clean, ink-saving Black & White study notes PDF**!\n\n"
        "✨ **What I do:**\n"
        "• Strips dark background colors & images\n"
        "• Automatically fits headlines & bullet points\n"
        "• Extracts slide text, tables & presenter notes\n"
        "• Generates printable A4 study sheets\n\n"
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

    status_msg = await update.message.reply_text("⏳ **Converting into clean B&W notes... Please wait.**", parse_mode="Markdown")

    try:
        # Download document into memory
        tg_file = await context.bot.get_file(document.file_id)
        raw_bytes = await tg_file.download_as_bytearray()
        file_bytes = bytes(raw_bytes)

        output_pdf_bytes = None

        if file_ext == "pptx":
            slides_data = extract_from_pptx(file_bytes)
            output_pdf_bytes = create_bw_text_pdf(slides_data)
        else:
            # Extract PDF content
            slides_data = extract_from_pdf(file_bytes)
            total_text_len = sum(len(s["title"]) + sum(len(c) for c in s["content"]) for s in slides_data)

            # If document has text, build formatted notes; if it's purely scanned images, use visual B&W engine
            if total_text_len > 30:
                try:
                    output_pdf_bytes = create_bw_text_pdf(slides_data)
                except Exception as text_err:
                    logger.warning(f"Text layout failed, using visual B&W fallback: {text_err}")
                    output_pdf_bytes = convert_image_pdf_to_bw(file_bytes)
            else:
                output_pdf_bytes = convert_image_pdf_to_bw(file_bytes)

        # Brand the output filename
        base_name = os.path.splitext(file_name)[0]
        output_filename = f"{base_name}_Yash_PPT_Converter_Bot.pdf"

        # Send file back to user
        await update.message.reply_document(
            document=io.BytesIO(output_pdf_bytes),
            filename=output_filename,
            caption=f"✅ **Converted successfully by {BRAND_NAME}!**\n📄 Clean Black & White Study Notes.",
            parse_mode="Markdown",
        )
        await status_msg.delete()

    except Exception as e:
        logger.error(f"Error during conversion: {e}", exc_info=True)
        # Ultimate fallback for any unhandled PDF issues
        if file_ext == "pdf":
            try:
                output_pdf_bytes = convert_image_pdf_to_bw(file_bytes)
                base_name = os.path.splitext(file_name)[0]
                output_filename = f"{base_name}_Yash_PPT_Converter_Bot.pdf"
                await update.message.reply_document(
                    document=io.BytesIO(output_pdf_bytes),
                    filename=output_filename,
                    caption=f"✅ **Converted successfully by {BRAND_NAME}!**\n📄 Printable Grayscale B&W Notes.",
                    parse_mode="Markdown",
                )
                await status_msg.delete()
                return
            except Exception:
                pass

        await status_msg.edit_text(f"❌ Conversion failed: `{str(e)}`", parse_mode="Markdown")


# -------------------------------------------------------------
# MAIN APP ENTRY
# -------------------------------------------------------------
def main():
    # Start background health server for Render
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
