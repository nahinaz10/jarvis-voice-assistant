# 🎙️ Jarvis — Voice Assistant for Disabled People

A Python-based voice assistant with a beautiful GUI, built to help **differently-abled individuals** control their computer hands-free using just their voice.

Built as a Python project to demonstrate real-world application of speech recognition and assistive technology.

---

## 💡 Purpose

Many people with physical disabilities struggle to use a keyboard or mouse. **Jarvis** allows them to control their computer, open apps, get information, and interact — all through voice commands. No hands needed.

---

## ✨ Features

- 🎤 **Voice Recognition** — Listens and understands natural speech
- 🔊 **Text to Speech** — Speaks responses back to the user
- 🖥️ **Opens Apps** — Calculator, Notepad, Camera, File Explorer, WhatsApp
- 🌐 **Opens Websites** — Google, YouTube, Weather
- 🕐 **Tells Time & Date** — Real-time clock information
- 😄 **Tells Jokes** — Keeps the user entertained
- ⚡ **Quick Command Buttons** — One tap for common commands
- 🌊 **Waveform Animation** — Visual feedback when listening
- 🎨 **Beautiful Warm UI** — Calm, accessible design for all users

---

## 🗣️ Voice Commands

| Say this... | Jarvis does... |
|---|---|
| "Hello" / "Hi" | Greets you |
| "What time is it?" | Tells current time |
| "What is today's date?" | Tells today's date |
| "Open Google" | Opens Google in browser |
| "Open YouTube" | Opens YouTube |
| "Open Calculator" | Launches Calculator app |
| "Open Notepad" | Launches Notepad |
| "Open Camera" | Opens Camera app |
| "Open WhatsApp" | Launches WhatsApp |
| "Weather" | Opens Google Weather |
| "Tell me a joke" | Tells a joke |
| "Who are you?" | Introduces itself |
| "Help" | Lists all commands |
| "Goodbye" / "Stop" | Shuts down Jarvis |

---

## 🛠️ How It Works

```
Microphone Input  →  [Voice Input Module]  →  Audio
Audio  →  [Speech Recognition Module]  →  Text
Text  →  [Text Processing Module]  →  Intent
Intent  →  [Command Execution Module]  →  Response
Response  →  [Output Module]  →  Voice + Screen
```

---

## 📂 Project Structure

```
jarvis-voice-assistant/
│
├── Jarvis_speech_assistant.py   # Main application
└── README.md                    # Project documentation
```

---

## ⚙️ Requirements

```bash
pip install customtkinter
pip install SpeechRecognition
pip install pywin32
pip install pyaudio
```

---

## ▶️ How to Run

```bash
python Jarvis_speech_assistant.py
```

> ⚠️ **Note:** This project runs on **Windows only** (uses Windows SAPI for text-to-speech and Windows apps like calc.exe, notepad.exe)

---

## 🧠 Concepts Used

| Concept | Usage |
|---|---|
| Speech Recognition | Google Speech API via `speech_recognition` |
| Text to Speech | Windows SAPI via `win32com` |
| Intent Detection | Keyword matching system |
| GUI Development | CustomTkinter framework |
| Multithreading | Non-blocking voice listening |
| Module Design | Separated into 5 clean modules |

---

## 🖥️ Tech Stack

| Technology | Purpose |
|---|---|
| Python | Core programming language |
| CustomTkinter | GUI framework |
| SpeechRecognition | Voice input |
| win32com (SAPI) | Text to speech output |
| Threading | Background listening |

---

## 👨‍💻 Author

**Nahidha Wasim I**
B.E. Computer Science and Engineering
K.Ramakrishnan College of Engineering, Samayapuram, Trichy

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue?logo=linkedin)](https://www.linkedin.com/in/nahidha-wasim-i-762177397)
[![GitHub](https://img.shields.io/badge/GitHub-Follow-black?logo=github)](https://github.com/nahinaz10)
