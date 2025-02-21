from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor


def add_text_to_frame(text_frame, text):
    """Formats text properly: bold for subheadings, italic for sub-subheadings, and avoids overflow."""
    text_frame.clear()
    text_frame.word_wrap = True
    text_frame.margin_top = Inches(0.1)
    text_frame.margin_bottom = Inches(0.1)
    text_frame.margin_left = Inches(0.1)
    text_frame.margin_right = Inches(0.1)

    # Splitting text into sections
    lines = text.split("\n")
    for line in lines:
        if line.strip():
            paragraph = text_frame.add_paragraph()
            run = paragraph.add_run()
            run.text = line.strip()
            run.font.size = Pt(18)

            # Formatting rules
            if line.startswith("## "):  # Subheading
                run.font.bold = True
                run.text = line.replace("## ", "")  # Remove ##
            elif line.startswith("### "):  # Sub-subheading
                run.font.italic = True
                run.text = line.replace("### ", "")  # Remove ###


def modify_image_path(image_path, part_index):
    """Replaces '1' with '2' in the image filename for continued slides."""
    if part_index == 0:
        return image_path
    return image_path.replace("1", "2", 1)  # Replace first occurrence of '1' with '2'


def create_presentation(topic, sections, slides_data, images_data, inverted_image_y_positions=None):
    """Creates a PowerPoint presentation with alternating layouts and custom Y-axis for inverted images."""
    presentation = Presentation()

    # Define dimensions
    text_top = Inches(2)
    text_width = Inches(5)
    text_height = Inches(6)

    # Define image size
    image_width = Inches(3.5)
    image_height = Inches(3)

    # Default to an empty dictionary if no custom Y-values are provided
    if inverted_image_y_positions is None:
        inverted_image_y_positions = {}

    for i, (section, (slide_text, images)) in enumerate(zip(sections, zip(slides_data, images_data))):
        # Split long text into multiple slides
        max_chars_per_slide = 350
        text_parts = [slide_text[i:i + max_chars_per_slide] for i in range(0, len(slide_text), max_chars_per_slide)]

        for part_index, text_part in enumerate(text_parts):
            slide = presentation.slides.add_slide(presentation.slide_layouts[5])

            # Set title (Only for the first slide of the section)
            title_box = slide.shapes.title
            if title_box:
                title_box.text = section if part_index == 0 else f"{section} (Cont.)"

            # Define positions based on layout
            if i % 2 == 0:
                # Original layout: text left, image right
                text_x = Inches(0.5)
                image_x = Inches(5.5)
                image_y = Inches(2.5)  # Default position
            else:
                # Inverted layout: text right, image left
                text_x = Inches(5)
                image_x = Inches(1)

                # Check for a manually defined Y-axis value
                image_y = inverted_image_y_positions.get(i, Inches(2.5))  # Default to 4.0 if not specified

            # Add text box
            content_box = slide.shapes.add_textbox(text_x, text_top, text_width, text_height)
            content_frame = content_box.text_frame
            add_text_to_frame(content_frame, text_part)

            # Add image on every slide of the section
            if images:
                image_filename = modify_image_path(images[0], part_index)  # Modify image path for continued slides
                slide.shapes.add_picture(image_filename, image_x, image_y, image_width, image_height)

    # Save the PowerPoint
    ppt_path = f"{topic.replace(' ', '_')}.pptx"
    presentation.save(ppt_path)
    return ppt_path
