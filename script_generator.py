from openai import OpenAI
import re

def generate_text(topic, section):
    client = OpenAI(
        base_url="https://integrate.api.nvidia.com/v1",
        api_key="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"
    )

    prompt = (
        f"Provide a comprehensive, detailed explanation about {topic} focusing on the section: '{section}'. "
        "Include relevant information, examples, and context. "
        "and also a detailed lecture script.do not start with i have i done this start directly from the topic begining and slides data should be properly arranged"
        "donnot add slide 1 comprehensive type of heading just add some bullet point and a little details if required "
    )

    completion = client.chat.completions.create(
        model="meta/llama-3.1-70b-instruct",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        top_p=0.7,
        max_tokens=1024,
        stream=True
    )

    section_details = ""
    for chunk in completion:
        if chunk.choices[0].delta.content is not None:
            section_details += chunk.choices[0].delta.content

    section_details = re.sub(r"\*", "", section_details)
    lecture_content = section_details[:500], section_details[500:]

    return  lecture_content