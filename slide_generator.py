from openai import OpenAI
import re

def trim_bullet_points(text, max_bullets=10):
    """Extracts up to `max_bullets` bullet points while keeping formatting intact."""
    bullet_points = re.split(r"\n[-•] ", text)  # Split on new lines with "-" or "•"
    trimmed_bullets = "\n- " + "\n- ".join(bullet_points[:max_bullets])  # Reconstruct with "-"
    return trimmed_bullets.strip() if trimmed_bullets != "\n- " else text  # Ensure we return valid text

def generate_text(topic, section):
    client = OpenAI(
        base_url="https://integrate.api.nvidia.com/v1",
        api_key="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"
    )

    prompt = (
        f"Create exactly **10 bullet points** for a PowerPoint slide on '{topic}', focusing on '{section}'.\n"
        "⚡ **Rules:**\n"
        "- Only list bullet points, NO extra text like 'Here is the response'.\n"
        "- Do NOT add introductions or explanations.\n"
        "- Each bullet point should be **short and impactful** (max 15 words each).\n"
        "- Do NOT exceed 650 characters in total.\n\n"
        "**Bullet Points:**"
    )

    completion = client.chat.completions.create(
        model="meta/llama-3.1-70b-instruct",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        top_p=0.7,
        max_tokens=4000,
        stream=True
    )

    section_details = ""
    for chunk in completion:
        if chunk.choices[0].delta.content is not None:
            section_details += chunk.choices[0].delta.content

    # Clean and trim response
    section_details = re.sub(r"\*", "", section_details)  # Remove unwanted symbols
    section_details = trim_bullet_points(section_details, max_bullets=10)  # Keep only 10 bullets

    slide_content = section_details, section_details  # Use trimmed content for slides

    return slide_content
