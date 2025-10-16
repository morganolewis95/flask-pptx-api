from flask import Flask, request, send_file
from flask_cors import CORS
from pptx import Presentation
from pptx.util import Inches
import tempfile
import os
import requests
from io import BytesIO

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

        image_url = slide_data.get("image_url")
        if image_url:
            try:
                response = requests.get(image_url)
                image_stream = BytesIO(response.content)
                slide.shapes.add_picture(image_stream, Inches(1), Inches(3.5), width=Inches(6))
            except Exception as e:
                print(f"Error adding image from {image_url}: {e}")

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(temp_file.name)

    return send_file(temp_file.name, as_attachment=True, download_name="generated_presentation.pptx")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=True)
