import pyttsx3

def generate_audio_offline(text, topic):
    engine = pyttsx3.init()
    audio_path = f"{topic.replace(' ', '_')}_lecture.mp3"
    engine.save_to_file(text, audio_path)
    engine.runAndWait()
    print(f"Offline TTS saved as {audio_path}")
    return audio_path
