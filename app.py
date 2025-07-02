import os
import fitz 
from flask import Flask, request, send_file, render_template, after_this_request
from werkzeug.utils import secure_filename
import atexit
import shutil

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Cleanup function to remove all uploads when the app exits
def cleanup():
    try:
        shutil.rmtree(UPLOAD_FOLDER)
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    except Exception as e:
        print(f"Error during cleanup: {e}")

atexit.register(cleanup)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        pdf_file = request.files["pdf"]
        pages_per_sheet = int(request.form["pages_per_sheet"])
        filename = secure_filename(pdf_file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        pdf_file.save(filepath)

        output_path = os.path.join(UPLOAD_FOLDER, "output_" + filename)
        process_pdf(filepath, output_path, pages_per_sheet)

        @after_this_request
        def cleanup(response):
            try:
                os.remove(filepath)
                os.remove(output_path)
            except Exception as e:
                app.logger.error(f"Error removing files: {e}")
            return response

        return send_file(output_path, as_attachment=True)

    return render_template("index.html")

def process_pdf(input_path, output_path, n):
    doc = fitz.open(input_path)
    total_pages = len(doc)

    # Split into odd/even pages
    odd_pages = [doc[i] for i in range(total_pages) if (i + 1) % 2 == 1]
    even_pages = [doc[i] for i in range(total_pages) if (i + 1) % 2 == 0]

    # Sort to preserve natural order (critical fix)
    odd_pages.sort(key=lambda p: p.number)
    even_pages = swap_neighbors(even_pages)

    # Create N-up groups
    odd_groups = make_nup_groups(odd_pages, n)
    even_groups = make_nup_groups(even_pages, n)

    # Interleave odd/even groups
    final_output = fitz.open()
    for i in range(max(len(odd_groups), len(even_groups))):
        if i < len(odd_groups):
            final_output.insert_pdf(odd_groups[i])
        if i < len(even_groups):
            final_output.insert_pdf(even_groups[i])

    # Save final file
    final_output.save(output_path, deflate=True)
    doc.close()

def swap_neighbors(pages):
    swapped = pages[:]
    for i in range(0, len(swapped) - 1, 2):
        swapped[i], swapped[i+1] = swapped[i+1], swapped[i]
    return swapped

def make_nup_groups(pages, n):
    source_path = pages[0].parent.name
    w, h = fitz.paper_size("a4")
    cols = 1 if n == 1 else 2
    rows = (n + 1) // 2
    grouped_docs = []

    for i in range(0, len(pages), n):
        out = fitz.open()
        sheet = out.new_page(width=w, height=h)
        src = fitz.open(source_path)

        for j, page in enumerate(pages[i:i+n]):
            page_number = page.number
            img_rect = fitz.Rect(
                (j % cols) * (w / cols),
                (j // cols) * (h / rows),
                ((j % cols) + 1) * (w / cols),
                ((j // cols) + 1) * (h / rows)
            )
            sheet.show_pdf_page(
                img_rect,
                src,
                page_number,
                rotate=0,
                oc=0,
                keep_proportion=True,
                overlay=True
            )

        grouped_docs.append(out)
        src.close()

    return grouped_docs

if __name__ == "__main__":
    app.run(debug=True)