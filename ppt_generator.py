from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import difflib
from spellchecker import SpellChecker


def correct_section_name(section, valid_sections):
    """Corrects the section name based on predefined sections and valid sections."""
    predefined_sections = {
        "INTRODUCTION": "Introduction", "HISTORY": "History", "APPLICATIONS": "Applications",
        "BENEFITS": "Benefits", "CHALLENGES": "Challenges", "INTRO": "Introduction", "STRY": "Story",
        "STORY": "Story", "CONCLUSION": "Conclusion", "SUMMARY": "Summary", "OUTLINE": "Outline",
        "METHODS": "Methods", "RESULTS": "Results", "DISCUSSION": "Discussion", "CONCLUSIONS": "Conclusions",
        "REVIEW": "Review"
    }

    spell = SpellChecker()
    section_corrected = spell.correction(section) or section  # Try to correct spelling
    section_upper = section_corrected.upper()

    # Check predefined sections first
    if section_upper in predefined_sections:
        return predefined_sections[section_upper].upper()

    # Check against valid sections
    valid_sections_upper = {s.upper(): s for s in valid_sections}
    closest_match = difflib.get_close_matches(section_upper, valid_sections_upper.keys(), n=1, cutoff=0.6)
    if closest_match:
        return valid_sections_upper[closest_match[0]].upper()

    return section_corrected.upper()


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
        # Correct section name
        corrected_section = correct_section_name(section, sections).upper()

        # Split long text into multiple slides
        max_chars_per_slide = 550
        text_parts = [slide_text[i:i + max_chars_per_slide] for i in range(0, len(slide_text), max_chars_per_slide)]

        for part_index, text_part in enumerate(text_parts):
            slide = presentation.slides.add_slide(presentation.slide_layouts[5])

            # Set title (Only for the first slide of the section)
            title_box = slide.shapes.title
            if title_box:
                title_box.text = corrected_section if part_index == 0 else f"{corrected_section} "

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
                image_y = inverted_image_y_positions.get(i, Inches(2.5))  # Default to 4.0 if not specified

            # Add text box
            content_box = slide.shapes.add_textbox(text_x, text_top, text_width, text_height)
            content_frame = content_box.text_frame
            add_text_to_frame(content_frame, text_part)

            # Add image on every slide of the section
            if images:
                image_index = part_index % len(images)  # Cycle through up to 5 images
                image_filename = images[image_index]  # Select the appropriate image
                slide.shapes.add_picture(image_filename, image_x, image_y, image_width, image_height)

    # Save the PowerPoint
    ppt_path = f"{topic.replace(' ', '_')}.pptx"
    presentation.save(ppt_path)
    return ppt_path
