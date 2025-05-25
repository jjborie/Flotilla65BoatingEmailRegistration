import mammoth
import docx
import base64
import os
import re

# Configuration
DOCX_PATH = "1 Enrollment Confirmation Boating Course June 14, 2025 CRYC.docx"
OUTPUT_HTML_PATH = "confirmation.html"

def extract_images_fallback(docx_path):
    """Fallback image extraction using python-docx, prioritizing heading images."""
    try:
        doc = docx.Document(docx_path)
        image_data_urls = []
        image_positions = []
        paragraph_index = 0

        # Count paragraphs and check inline shapes for images
        for para in doc.paragraphs:
            paragraph_index += 1
        for shape in doc.inline_shapes:
            try:
                if shape._inline.graphic.graphicData.uri == "http://schemas.openxmlformats.org/drawingml/2006/picture":
                    blip = shape._inline.graphic.graphicData.pic.blipFill.blip
                    rel_id = blip.embed or blip.link
                    if rel_id is None:
                        print(f"Fallback: Skipping image: No embed or link attribute found.")
                        continue
                    
                    image_rel = doc.part.rels.get(rel_id)
                    if image_rel is None:
                        print(f"Fallback: Skipping image: No relationship found for rel_id {rel_id}.")
                        continue
                    
                    image_part = image_rel.target_part
                    image_data = image_part.blob
                    content_type = image_part.content_type
                    ext = "png" if content_type == "image/png" else "jpg" if content_type == "image/jpeg" else "png"
                    base64_data = base64.b64encode(image_data).decode("utf-8")
                    data_url = f"data:image/{ext};base64,{base64_data}"
                    # Assume first image is at start (heading)
                    image_data_urls.append((data_url, ext))
                    image_positions.append(0)  # Force heading position
                    print(f"Fallback: Embedded image at paragraph index 0 (forced heading), content type: {content_type}")
            except Exception as e:
                print(f"Fallback: Failed to process image: {str(e)}")
                continue

        return image_data_urls, image_positions
    except Exception as e:
        print(f"Fallback: Failed to process DOCX: {str(e)}")
        return [], []

def convert_docx_to_html(docx_path, output_path):
    """Convert DOCX to HTML, embedding images as Base64 at the correct position."""
    image_counter = 0
    image_data_urls = []
    
    def handle_image(image):
        nonlocal image_counter, image_data_urls
        try:
            image_counter += 1
            ext = "png" if image.content_type == "image/png" else "jpg" if image.content_type == "image/jpeg" else "png"
            # Use get_stream() for newer mammoth versions
            try:
                image_data = image.get_stream().read()
            except AttributeError:
                # Fallback to get_reader() for older mammoth versions
                image_data = image.get_reader().read()
            base64_data = base64.b64encode(image_data).decode("utf-8")
            data_url = f"data:image/{ext};base64,{base64_data}"
            image_data_urls.append((data_url, ext))
            print(f"Mammoth: Embedded image {image_counter}, content type: {image.content_type}")
            return {"src": data_url, "alt": f"Course Image {image_counter}", "style": "max-width:100%;height:auto;"}
        except Exception as e:
            print(f"Mammoth: Failed to process image {image_counter}: {str(e)}")
            return {"src": ""}

    try:
        with open(docx_path, "rb") as f:
            result = mammoth.convert_to_html(f, convert_image=mammoth.images.img_element(handle_image))
            html_content = result.value
            messages = result.messages
    except Exception as e:
        print(f"Failed to convert DOCX to HTML: {str(e)}")
        return

    # Fallback to python-docx if no images extracted
    if not image_data_urls:
        print("No images extracted by mammoth. Trying python-docx fallback.")
        image_data_urls, image_positions = extract_images_fallback(docx_path)
        if image_data_urls:
            # Prepend first image as heading
            image_tag = f'<img src="{image_data_urls[0][0]}" alt="Course Image 1" style="max-width:100%;height:auto;">'
            html_content = f"{image_tag}\n{html_content}"
            print("Added image at the start as heading")
            # Append any additional images
            for i, (data_url, ext) in enumerate(image_data_urls[1:], 2):
                image_tag = f'<img src="{data_url}" alt="Course Image {i}" style="max-width:100%;height:auto;">'
                html_content += f"\n{image_tag}"
                print(f"Added image tag: {image_tag}")

    # Wrap HTML in a proper structure, escaping curly braces in CSS
    final_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Boating Course Enrollment Confirmation</title>
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; }}
        .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
        h1 {{ color: #004080; }}
        .section {{ margin-bottom: 20px; }}
        img {{ max-width: 100%; height: auto; display: block; margin: 0 auto; }}
    </style>
</head>
<body>
    <div class="container">
        {html_content}
    </div>
</body>
</html>"""

    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(final_html)
    except Exception as e:
        print(f"Failed to save HTML file: {str(e)}")
        return

    print(f"HTML saved to {output_path}")
    if not image_data_urls:
        print("No images extracted from the DOCX file.")
    if messages:
        print("Conversion warnings:", messages)

def main():
    convert_docx_to_html(DOCX_PATH, OUTPUT_HTML_PATH)

if __name__ == "__main__":
    main()