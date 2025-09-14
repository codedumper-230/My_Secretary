import speech_recognition as sr
import pyttsx3

# Initialize TTS engine
engine = pyttsx3.init()

def speak(text):
    engine.say(text)
    engine.runAndWait()

def listen():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("🎤 Listening...")
        audio = recognizer.listen(source)
        try:
            text = recognizer.recognize_google(audio)
            print("✅ You said:", text)
            return text
        except sr.UnknownValueError:
            print("❌ Sorry, I didn’t catch that.")
            return ""
        except sr.RequestError:
            print("❌ API unavailable.")
            return ""
