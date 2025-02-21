from pptx import Presentation
from pptx.util import Inches

def create_presentation(topic, sections, slides_data, images_data):
    presentation = Presentation()

    for section, (slide_text, images) in zip(sections, zip(slides_data, images_data)):
        slide = presentation.slides.add_slide(presentation.slide_layouts[5])

        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
        title_box.text_frame.text = section

        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(6), Inches(5.5))
        content_box.text_frame.text = slide_text

        for idx, img_path in enumerate(images):
            slide.shapes.add_picture(img_path, Inches(6.5), Inches(1.5 + idx * 3), Inches(3))

    ppt_path = f"{topic.replace(' ', '_')}.pptx"
    presentation.save(ppt_path)
    return ppt_path
