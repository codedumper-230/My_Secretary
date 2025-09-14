import os
import sys
import pvporcupine
import pyaudio
import struct
import speech_recognition as sr
import threading
import base64
from dotenv import load_dotenv


load_dotenv()

def get_porcupine_key():
    encoded = "SGErUDJRbllSYnppNHgwSTVDUWZvTGpCTU1kdTRpUFdVckxxZUVKS3FIV0JZZ3JIOUg1NkN3PT0="  # Base64-encoded valid AccessKey
    return base64.b64decode(encoded).decode()

access_key = get_porcupine_key()

def resource_path(relative_path):
    """ Get absolute path to resource (compatible with PyInstaller) """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class WakeWordAgent:
    def __init__(self, callback, app=None, access_key=None):
        self.callback = callback
        self.running = False
        self.app = app  # ✅ Reference to EmailAgentApp for UI updates
        self.access_key = access_key  # ✅ Fix: store the passed access key


    def start(self):
        self.running = True
        self._thread = threading.Thread(target=self._listen)
        self._thread.daemon = True
        self._thread.start()

    def stop(self):
        self.running = False

    def _listen(self):
        print("Listening frame...")

        # ✅ Use local keyword path with safe access
        self.keyword_path = resource_path("wake_words/Hey-Secretary_en_windows_v3_0_0.ppn")

        porcupine = pvporcupine.create(
            access_key=self.access_key,
            keyword_paths=[self.keyword_path]
        )
        pa = pyaudio.PyAudio()
        stream = pa.open(rate=porcupine.sample_rate, channels=1, format=pyaudio.paInt16,
                         input=True, frames_per_buffer=porcupine.frame_length)

        print("🕵️ Listening for wake word...")

        try:
            while self.running:
                pcm = stream.read(porcupine.frame_length, exception_on_overflow=False)
                pcm = struct.unpack_from("h" * porcupine.frame_length, pcm)
                result = porcupine.process(pcm)
                if result >= 0:
                    print("🎤 Wake word detected!")
                    self._handle_command()
        finally:
            stream.stop_stream()
            stream.close()
            pa.terminate()
            porcupine.delete()

    def _handle_command(self):
        print("⚡ Wake trigger activated!")

        if self.app:
            self.app.root.after(0, self.app.show_listening_orb)

        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            print("🎧 Listening for command...")
            audio = recognizer.listen(source)

        try:
            command = recognizer.recognize_google(audio)
            print("🧠 Heard:", command)
            if self.app:
                self.app.root.after(0, self.app.hide_listening_orb)
            self.callback(command)
        except Exception as e:
            print("❌ Could not understand:", e)
            if self.app:
                self.app.root.after(0, self.app.hide_listening_orb)    
