import concurrent.futures
from slide_generator import generate_text as generate_slides
from script_generator import generate_text as generate_script
from image_scraper import search_and_download_images
from ppt_generator import create_presentation
from audio_generator import generate_audio_offline

def process_section(topic, section):
    slide_text = " ".join(generate_slides(topic, section))  # Combine tuple parts
    lecture_text = " ".join(generate_script(topic, section))  # Combine tuple parts
    images = search_and_download_images(f"{topic} {section}", num_images=5)
    return section, slide_text, lecture_text, images

def main():
    topic = input("Enter a topic: ")
    sections_input = input("Enter sections separated by commas (e.g., Introduction, History, Applications): ")
    sections = [section.strip() for section in sections_input.split(',')]

    print("\nStarting parallel processing for lecture script, slide text, and image downloading...\n")

    slides_data = {}
    lecture_script = f"Lecture Script for Topic: {topic}\n\n"
    images_data = {}

    # Run text generation and image downloading in parallel
    with concurrent.futures.ThreadPoolExecutor() as executor:
        future_to_section = {executor.submit(process_section, topic, section): section for section in sections}

        for future in concurrent.futures.as_completed(future_to_section):
            section, slide_text, lecture_text, images = future.result()
            slides_data[section] = slide_text
            lecture_script += f"### {section}\n{lecture_text}\n\n"
            images_data[section] = images

    print("\nCreating PowerPoint Presentation...\n")
    ppt_path = create_presentation(topic, sections, [slides_data[sec] for sec in sections],
                                   [images_data[sec] for sec in sections])
    print(f"Presentation saved as {ppt_path}")

    script_path = f"{topic.replace(' ', '_')}_lecture_script.txt"
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(lecture_script)
    print(f"Lecture script saved as {script_path}")

    print("\nGenerating lecture audio...\n")
    with concurrent.futures.ProcessPoolExecutor() as executor:
        audio_future = executor.submit(generate_audio_offline, lecture_script, topic)
        audio_path = audio_future.result()
    print(f"Lecture audio saved as {audio_path}")
    print("Slides Data:", slides_data)



if __name__ == "__main__":
    main()
