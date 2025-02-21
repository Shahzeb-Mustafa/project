from pptx import Presentation
from pptx.util import Inches

def split_text_into_chunks(text, max_chars=400):
    """Splits text into chunks of max_chars while keeping words intact."""
    words = text.split()
    chunks = []
    chunk = ""

    for word in words:
        if len(chunk) + len(word) + 1 <= max_chars:
            chunk += " " + word
        else:
            chunks.append(chunk.strip())
            chunk = word
    if chunk:
        chunks.append(chunk.strip())

    return chunks

def format_text(text_frame, text):
    """Formats the text inside the text box with bullet points and line breaks."""
    text_frame.clear()  # Clear any default formatting
    p = text_frame.add_paragraph()

    # Split text into sentences and add bullet points
    for line in text.split(". "):
        line = line.strip()
        if line:
            p = text_frame.add_paragraph()
            p.text = f"• {line}."
            p.space_after = Inches(0.1)  # Reduce spacing between bullet points

def create_presentation(topic, sections, slides_data, images_data):
    presentation = Presentation()

    # Define text box dimensions (6 inches wide, 8 inches high)
    text_left = Inches(0.5)
    text_top = Inches(1.5)
    text_width = Inches(6)   # Restored width to 6 inches
    text_height = Inches(8)  # Increased height to 8 inches

    # Define image dimensions and placement
    image_width = Inches(3)
    image_height = Inches(3)
    image_x = Inches(6.5)  # Top-right position
    image_y = Inches(3.0)  # Image placed at y = 3.0

    for section, (full_text, images) in zip(sections, zip(slides_data, images_data)):
        text_chunks = split_text_into_chunks(full_text)  # Split text into manageable parts

        for i, (text_chunk, img_path) in enumerate(zip(text_chunks, images)):
            slide = presentation.slides.add_slide(presentation.slide_layouts[5])  # Title Only layout

            # Set title (same title for all slides in the section)
            title_box = slide.shapes.title
            if title_box:
                title_box.text = section  # No "Part 1" or "Part 2"

            # Add text box (6 inches wide, 8 inches high)
            content_box = slide.shapes.add_textbox(text_left, text_top, text_width, text_height)
            content_frame = content_box.text_frame
            content_frame.word_wrap = True
            format_text(content_frame, text_chunk)  # Apply proper formatting

            # Set text box margins
            content_frame.margin_top = Inches(0.2)
            content_frame.margin_bottom = Inches(0.2)
            content_frame.margin_left = Inches(0.3)
            content_frame.margin_right = Inches(0.3)

            # Add image at (x=6.5, y=3.0)
            slide.shapes.add_picture(img_path, image_x, image_y, image_width, image_height)

    # Save the PowerPoint file
    ppt_path = f"{topic.replace(' ', '_')}.pptx"
    presentation.save(ppt_path)
    return ppt_path
