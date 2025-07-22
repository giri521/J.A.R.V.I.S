from gtts import gTTS
import os
import tempfile
import time
import platform

def speak_jarvis(text):
    try:
        tts = gTTS(text=text, lang='en', tld='co.in')  # Indian accent

        with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as fp:
            temp_path = fp.name
            tts.save(temp_path)

        # Play the MP3 file safely on Windows
        if platform.system() == "Windows":
            os.startfile(temp_path)
            # Wait enough time for audio to play before deleting
            time.sleep(len(text.split()) / 2)  # Approximate duration
        else:
            from playsound import playsound
            playsound(temp_path)

        os.remove(temp_path)

    except Exception as e:
        print("Error speaking:", e)

if __name__ == "__main__":
    speak_jarvis("Hello sir, I am your virtual assistant. All systems are online.")
