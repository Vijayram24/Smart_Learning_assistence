import google.generativeai as genai
import speech_recognition as sr
import win32com.client as wc
import tkinter as tk
from tkinter import scrolledtext, filedialog
from threading import Thread
import time
import pythoncom
from fpdf import FPDF
from dotenv import load_dotenv
import os
load_dotenv(dotenv_path='vijay.env')


# 🔐 API Key and Model
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
MODEL_NAME = "gemini-1.5-flash"
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel(MODEL_NAME)

# 🎙️ Recognize speech
def recognize_speech():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        status_var.set("🎤 Listening...")
        root.update()
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
    try:
        status_var.set("🧠 Recognizing...")
        root.update()
        return recognizer.recognize_google(audio)
    except:
        return "Sorry, I couldn’t understand you."

# 🤖 Get Gemini response
def generate_response(prompt):
    try:
        status_var.set("⚡ Generating response...")
        root.update()
        response = model.generate_content(prompt)
        return response.text.strip()
    except:
        return "Error generating response."

# 🗣️ Speak and show each line in real-time
def speak_line_by_line(text, voice_index=1):
    try:
        speaker = wc.Dispatch("SAPI.SpVoice")
        voices = speaker.GetVoices()
        if voice_index < voices.Count:
            speaker.Voice = voices.Item(voice_index)
        speaker.Rate = -1
        speaker.Volume = 100

        lines = [line.strip() for line in text.splitlines() if line.strip()]
        response_box.delete('1.0', tk.END)
        for line in lines:
            response_box.insert(tk.END, line + '\n')
            response_box.see(tk.END)
            root.update()
            speaker.Speak(line)
            time.sleep(0.1)
    except Exception as e:
        response_box.insert(tk.END, f"\n❌ Speaking error: {e}\n")

# 🎬 Assistant thread
def run_assistant():
    pythoncom.CoInitialize()
    user_input = recognize_speech()
    user_input_var.set(user_input)
    if user_input.lower() in ["exit", "quit", "stop"]:
        root.quit()
        return
    response = generate_response(user_input)
    speak_line_by_line(response)
    status_var.set("✅ Done")

def start_thread():
    Thread(target=run_assistant).start()

def run_text_input():
    prompt = user_input_var.get()
    Thread(target=handle_text_prompt, args=(prompt,)).start()

def handle_text_prompt(prompt):
    pythoncom.CoInitialize()
    response = generate_response(prompt)
    speak_line_by_line(response)
    status_var.set("✅ Done")

def export_to_pdf():
    text = response_box.get("1.0", tk.END).strip()
    if not text:
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if file_path:
        try:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.set_font("Helvetica", size=12)  # Avoid deprecated Arial
            for line in text.splitlines():
                pdf.multi_cell(0, 10, line)
            pdf.output(file_path)
            status_var.set("✅ PDF Exported Successfully")
        except Exception as e:
            status_var.set(f"❌ PDF Export Failed: {e}")

# 🌟 GUI Setup
root = tk.Tk()
root.title("Vijay Prepration Assistant")
root.geometry("1000x700")
root.configure(bg="#f1f3f6")

font_label = ("Segoe UI", 12)
font_text = ("Segoe UI", 13)
color_fg = "#212121"
color_accent = "#2874f0"

# Title Label
title_label = tk.Label(root, text="🛍️ Vijay's  learning  Assistant", bg="#f1f3f6", fg=color_accent, font=("Segoe UI", 20, "bold"))
title_label.pack(pady=15)

# Input Frame
input_frame = tk.Frame(root, bg="#ffffff", padx=20, pady=15, bd=1, relief=tk.SOLID)
input_frame.pack(pady=10, padx=20, fill="x")

input_label = tk.Label(input_frame, text="📝 Type or  Speak:", font=font_label, fg=color_fg, bg="#ffffff")
input_label.pack(anchor="w")

user_input_var = tk.StringVar()
input_entry = tk.Entry(input_frame, textvariable=user_input_var, font=font_text, width=70, bg="#ffffff", fg=color_fg, insertbackground="#000000", bd=1, relief=tk.SOLID)
input_entry.pack(pady=5, fill="x")

# Buttons Frame
btn_frame = tk.Frame(input_frame, bg="#ffffff")
btn_frame.pack(pady=10)

def make_button(parent, text, command):
    return tk.Button(parent, text=text, command=command, bg=color_accent, fg="white", font=("Segoe UI", 10, "bold"), padx=12, pady=6, relief=tk.FLAT, activebackground="#0059c1")

make_button(btn_frame, "🎙️ Voice Input", start_thread).pack(side="left", padx=10)
make_button(btn_frame, "📨 Submit Text", run_text_input).pack(side="left", padx=10)
make_button(btn_frame, "📄 Export to PDF", export_to_pdf).pack(side="left", padx=10)

# Response Display Area
response_frame = tk.Frame(root, bg="#ffffff", padx=20, pady=15, bd=1, relief=tk.SOLID)
response_frame.pack(pady=10, padx=20, fill="both", expand=True)

response_label = tk.Label(response_frame, text="🤖 vijay this is ans:", font=font_label, fg=color_fg, bg="#ffffff")
response_label.pack(anchor="w")

response_box = scrolledtext.ScrolledText(response_frame, font=font_text, wrap=tk.WORD, bg="#ffffff", fg=color_fg, insertbackground="black")
response_box.pack(fill="both", expand=True, pady=5)

# Status Bar
status_var = tk.StringVar(value="🔵 Ready")
status_label = tk.Label(root, textvariable=status_var, fg="gray", bg="#f1f3f6", font=("Segoe UI", 10))
status_label.pack(pady=(5, 10))

# Start GUI
root.mainloop()
