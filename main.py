import os
import time
import threading
import pythoncom
import win32com.client
import speech_recognition as sr
from transformers import pipeline
import openai
from google.generativeai import GenerativeModel
import difflib

# Config apinyaaa!!
openai.api_key = os.getenv("OPENAI_API_KEY")  # ini apinyaa!!
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")  # ini apinyaa!!

from google.generativeai import configure
configure(api_key=GOOGLE_API_KEY)
gemini_model = GenerativeModel('gemini-pro')

# ini buat transkrip audionya
asr_pipeline = pipeline("automatic-speech-recognition", model="openai/whisper-base")
recognizer = sr.Recognizer()

# ini kontrol powerpointnyaaa
class PowerPointController:
    def __init__(self):
        pythoncom.CoInitialize()
        self.ppt = win32com.client.Dispatch("PowerPoint.Application")
        self.presentation = self.ppt.Presentations(1)

    def get_current_slide_index(self):
        return self.presentation.SlideShowWindow.View.CurrentShowPosition

    def get_slide_text(self, index):
        if 1 <= index <= self.presentation.Slides.Count:
            slide = self.presentation.Slides(index)
            return " ".join([
                shape.TextFrame.TextRange.Text
                for shape in slide.Shapes if shape.HasTextFrame
            ])
        return ""

    def next_slide(self):
        self.presentation.SlideShowWindow.View.Next()

# in buat bandingin teks terakhir
def is_precise_slide_end(spoken_text, slide_text):
    slide_text = slide_text.strip().lower()
    spoken_text = spoken_text.strip().lower()

    # Bersihin tanda baca
    for char in ['.', ',', ';', ':', '!', '?', '\"', "'"]:
        slide_text = slide_text.replace(char, '')
        spoken_text = spoken_text.replace(char, '')

    # Ambil 2-3 kata terakhir dari spoken_text
    words = spoken_text.split()
    end_phrase = " ".join(words[-3:] if len(words) >= 3 else words)
 
    # kt bandingin sama slide sebelum selanjutnya
    ratio = difflib.SequenceMatcher(None, end_phrase, slide_text[-len(end_phrase):]).ratio()
    print(f"Match ratio (akhir 2-3 kata): {ratio:.2f}")
    return ratio > 0.85  # bisa disesuaikan

# cek slide baru pake gemini ai
def check_gemini_confirmation(first_words, slide_text):
    try:
        response = gemini_model.generate_content(
            f"Do these first words match the start of the new slide?\nWords: '{first_words}'\nSlide: '{slide_text}'\nReply 'yes' or 'no'"
        )
        answer = response.text.strip().lower()
        return 'yes' in answer
    except Exception as e:
        print(f"Gemini error: {e}")
        return False

# ini main processnya
def start_audio_control():
    ppt = PowerPointController()
    current_index = ppt.get_current_slide_index()
    current_slide_text = ppt.get_slide_text(current_index)

    print(f"Mulai di slide {current_index}:")
    print(current_slide_text)
    print("========================\n")

    while True:
        with sr.Microphone() as source:
            print("Mendengarkan...")
            audio = recognizer.listen(source, phrase_time_limit=10)

        try:
            print("Transkripsi...")
            audio_bytes = audio.get_wav_data()
            spoken_text = asr_pipeline(audio_bytes)["text"]
            print(f"Teks Didengar: {spoken_text}")

            if is_precise_slide_end(spoken_text, current_slide_text):
                print("Akhir slide terdeteksi. Ganti slide...")
                ppt.next_slide()

                #tunggu transisinya
                time.sleep(0.75)

                new_index = ppt.get_current_slide_index()
                new_slide_text = ppt.get_slide_text(new_index)
                first_words = " ".join(spoken_text.strip().split()[:5])

                if check_gemini_confirmation(first_words, new_slide_text):
                    print(f"Slide pindah ke {new_index} berhasil dikonfirmasi!")
                    current_index = new_index
                    current_slide_text = new_slide_text
                else:
                    print("Verifikasi Gemini gagal.")
            else:
                print("Belum akhir slide.")

        except Exception as e:
            print(f"Terjadi kesalahan: {e}")

# jalankaaaan!!
if __name__ == '__main__':
    print("Mulai sistem pengontrol slide otomatis dengan suara...")
    t = threading.Thread(target=start_audio_control)
    t.start()