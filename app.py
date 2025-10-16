from flask import Flask, request, send_file
from flask_cors import CORS
from pptx import Presentation
import tempfile
import os

app = Flask(__name__)
CORS(app)

@app.route('/generate_ppt', methods=['POST'])
def generate_ppt():
    data = request.json
    slides = data.get("slides", [])

    prs = Presentation()

    for slide_data in slides:
        slide_layout = prs.slide_layouts[1]  # Title and Content layout
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        content = slide.placeholders[1]

        title.text = slide_data.get("title", "Untitled")
        content.text = "\n".join(slide_data.get("bullets", []))

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(temp_file.name)

    return send_file(temp_file.name, as_attachment=True, download_name="generated_presentation.pptx")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=True)
