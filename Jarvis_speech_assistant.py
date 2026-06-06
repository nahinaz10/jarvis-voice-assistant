import customtkinter as ctk
import speech_recognition as sr
import win32com.client
import datetime
import webbrowser
import time
import os
import random
import subprocess
import threading

# ── Theme ──────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("green")

NUDE = {
    "bg":        "#EFE6DC",   # sandal background
    "card":      "#E6D7C7",   # darker sandal card
    "card2":     "#DDCBB9",   # secondary panel
    "border":    "#CBB6A2",

    "accent":    "#B08968",   # sandal brown button
    "accent_dk": "#8C6A50",   # darker hover brown

    "text":      "#3E2F25",   # rich brown text
    "text_mid":  "#6B5A4E",
    "text_lt":   "#9A8776",

    "green":     "#8F9779",   # muted olive highlight
    "white":     "#FAF6F1",
}

# ── Modules ────────────────────────────────────────────────────────────────────

class OutputModule:
    def speak(self, text, label_callback=None):
        if label_callback:
            label_callback(text)
        # Create a new COM instance on this thread to avoid threading issues
        spk = win32com.client.Dispatch("SAPI.SpVoice")
        spk.Speak(text)


class VoiceInputModule:
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self._calibrate()

    def _calibrate(self):
        with sr.Microphone() as source:
            self.recognizer.adjust_for_ambient_noise(source, duration=1)

    def capture_audio(self):
        with sr.Microphone() as source:
            try:
                self.recognizer.adjust_for_ambient_noise(source, duration=0.3)
                audio = self.recognizer.listen(source, timeout=10, phrase_time_limit=15)
                return audio
            except sr.WaitTimeoutError:
                return None


class SpeechRecognitionModule:
    def __init__(self):
        self.recognizer = sr.Recognizer()

    def recognize(self, audio):
        if audio is None:
            return None
        try:
            text = self.recognizer.recognize_google(audio)
            return text.lower()
        except (sr.UnknownValueError, sr.RequestError):
            return None


class TextProcessingModule:
    INTENTS = {
        "greeting":    ["hello", "hi", "hey", "good morning", "good evening", "good afternoon"],
        "time":        ["time", "what time", "current time", "tell me the time"],
        "date":        ["date", "what date", "today's date", "what day", "what is today"],
        "google":      ["open google", "google", "search google"],
        "youtube":     ["open youtube", "youtube", "play youtube"],
        "calculator":  ["open calculator", "calculator", "calc"],
        "notepad":     ["open notepad", "notepad", "text editor"],
        "camera":      ["open camera", "camera", "take photo"],
        "files":       ["open files", "file explorer", "my files"],
        "whatsapp":    ["open whatsapp", "whatsapp"],
        "weather":     ["weather", "temperature", "is it raining"],
        "joke":        ["joke", "tell me a joke", "make me laugh"],
        "name":        ["your name", "who are you", "what are you"],
        "how_are_you": ["how are you", "how are you doing"],
        "thanks":      ["thank you", "thanks"],
        "help":        ["help", "what can you do", "commands"],
        "stop":        ["stop", "exit", "quit", "bye", "goodbye", "shut down"],
    }

    def get_intent(self, text):
        if not text:
            return "unknown"
        for intent, keywords in self.INTENTS.items():
            for keyword in keywords:
                if keyword in text:
                    return intent
        return "unknown"


class CommandExecutionModule:
    JOKES = [
        "Why don't scientists trust atoms? Because they make up everything!",
        "Why did the computer go to the doctor? Because it had a virus!",
        "What do you call a fish with no eyes? A fsh!",
        "Why can't your nose be 12 inches long? Because then it would be a foot!",
        "Why did the scarecrow win an award? Because he was outstanding in his field!",
    ]
    GREETINGS = [
        "Hello! Great to hear from you! How can I assist you today?",
        "Hey there! Jarvis is here and ready to help!",
        "Hi! Always a pleasure. What do you need?",
    ]
    THANKS_RESPONSES = [
        "You are most welcome! Always happy to help!",
        "Anytime! That is what I am here for!",
        "My pleasure! Let me know if you need anything else!",
    ]
    _joke_index = 0

    def execute(self, intent):
        if intent == "greeting":
            return random.choice(self.GREETINGS)
        elif intent == "how_are_you":
            return "I am doing great, thank you for asking!"
        elif intent == "thanks":
            return random.choice(self.THANKS_RESPONSES)
        elif intent == "time":
            now = datetime.datetime.now()
            return f"The current time is {now.strftime('%I:%M %p')}."
        elif intent == "date":
            now = datetime.datetime.now()
            return f"Today is {now.strftime('%A, %B %d, %Y')}."
        elif intent == "google":
            webbrowser.open("https://www.google.com")
            return "Opening Google for you!"
        elif intent == "youtube":
            webbrowser.open("https://www.youtube.com")
            return "Opening YouTube. Enjoy watching!"
        elif intent == "calculator":
            subprocess.Popen("calc.exe", shell=True)
            return "Opening Calculator!"
        elif intent == "notepad":
            subprocess.Popen("notepad.exe", shell=True)
            return "Opening Notepad. Ready to take notes!"
        elif intent == "camera":
            subprocess.Popen("start microsoft.windows.camera:", shell=True)
            return "Opening Camera. Smile!"
        elif intent == "files":
            subprocess.Popen("explorer.exe", shell=True)
            return "Opening File Explorer!"
        elif intent == "whatsapp":
            subprocess.Popen("start whatsapp:", shell=True)
            return "Opening WhatsApp!"
        elif intent == "weather":
            webbrowser.open("https://www.google.com/search?q=weather+today")
            return "I have opened Google weather for you!"
        elif intent == "joke":
            joke = self.JOKES[self._joke_index % len(self.JOKES)]
            CommandExecutionModule._joke_index += 1
            return joke
        elif intent == "name":
            return "I am Jarvis, your personal voice assistant!"
        elif intent == "help":
            return "Say: Hello, Time, Date, Open Google, YouTube, Calculator, Notepad, Camera, WhatsApp, Files, Weather, Joke, or Goodbye!"
        elif intent == "stop":
            return "STOP"
        else:
            return "I didn't catch that. Say help to see what I can do!"


# ── GUI ────────────────────────────────────────────────────────────────────────
class JarvisApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("JARVIS — Voice Assistant")
        self.geometry("480x780")
        self.resizable(False, False)
        self.configure(fg_color=NUDE["bg"])

        self.output     = OutputModule()
        self.voice_in   = VoiceInputModule()
        self.recognizer = SpeechRecognitionModule()
        self.processor  = TextProcessingModule()
        self.executor   = CommandExecutionModule()

        self.running   = False
        self.cmd_count = 0

        self._build_ui()
        self._update_clock()
        # ── Startup greeting ──
        self.after(1000, lambda: self._speak(
            "Hello! I am Jarvis, your voice assistant. How can I help you?"))

    # ── UI Construction ────────────────────────────────────────────────────────
    def _build_ui(self):

        # ── Top bar ──
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=20, pady=(20, 0))

        ctk.CTkLabel(top, text="JARVIS", font=("Georgia", 22, "bold"),
                     text_color=NUDE["text"]).pack(side="left")
        ctk.CTkLabel(top, text="voice assistant", font=("Georgia", 11, "italic"),
                     text_color=NUDE["text_lt"]).pack(side="left", padx=(8, 0), pady=(6, 0))

        self.status_dot = ctk.CTkLabel(top, text="● Idle",
                                       font=("Helvetica", 12),
                                       text_color=NUDE["text_lt"])
        self.status_dot.pack(side="right", padx=4)

        # ── Divider ──
        ctk.CTkFrame(self, height=1, fg_color=NUDE["border"]).pack(
            fill="x", padx=20, pady=(12, 0))

        # ── Waveform area ──
        wave_card = ctk.CTkFrame(self, fg_color=NUDE["card"],
                                 corner_radius=16,
                                 border_width=1, border_color=NUDE["border"])
        wave_card.pack(fill="x", padx=20, pady=(16, 0))

        self.canvas = ctk.CTkCanvas(wave_card, height=70,
                                    bg=NUDE["card"], highlightthickness=0)
        self.canvas.pack(fill="x", padx=20, pady=(16, 8))
        self._draw_idle_wave()

        # Mic button
        self.mic_btn = ctk.CTkButton(
            wave_card,
            text="🎤  Start Listening",
            font=("Helvetica", 14, "bold"),
            fg_color=NUDE["accent"],
            hover_color=NUDE["accent_dk"],
            text_color=NUDE["white"],
            corner_radius=30,
            height=48,
            command=self._toggle_listen,
        )
        self.mic_btn.pack(pady=(4, 18), padx=60)

        # ── You said ──
        ctk.CTkLabel(self, text="YOU SAID",
                     font=("Helvetica", 10),
                     text_color=NUDE["text_lt"]).pack(anchor="w", padx=24, pady=(14, 2))

        self.you_said = ctk.CTkLabel(
            self,
            text="Tap the mic button to begin…",
            font=("Georgia", 14, "italic"),
            text_color=NUDE["text_mid"],
            wraplength=420, justify="left",
            fg_color=NUDE["white"],
            corner_radius=10,
            anchor="w",
        )
        self.you_said.pack(fill="x", padx=20, ipady=10, ipadx=14)

        # ── Jarvis response ──
        ctk.CTkLabel(self, text="JARVIS",
                     font=("Helvetica", 10),
                     text_color=NUDE["text_lt"]).pack(anchor="w", padx=24, pady=(12, 2))

        resp_frame = ctk.CTkFrame(self, fg_color=NUDE["card"],
                                  corner_radius=10,
                                  border_width=1, border_color=NUDE["border"])
        resp_frame.pack(fill="x", padx=20)

        ctk.CTkLabel(resp_frame, text="J",
                     font=("Georgia", 14, "bold"),
                     text_color=NUDE["white"],
                     fg_color=NUDE["accent"],
                     width=34, height=34,
                     corner_radius=17).pack(side="left", padx=(12, 0), pady=12)

        self.resp_label = ctk.CTkLabel(
            resp_frame,
            text="Hello! I am Jarvis. How can I help you?",
            font=("Georgia", 13),
            text_color=NUDE["text"],
            wraplength=360, justify="left",
            anchor="w",
        )
        self.resp_label.pack(side="left", padx=12, pady=12)

        # ── Quick commands ──
        ctk.CTkLabel(self, text="QUICK COMMANDS",
                     font=("Helvetica", 10),
                     text_color=NUDE["text_lt"]).pack(anchor="w", padx=24, pady=(16, 6))

        cmds = [
            ("🕐 Time",       "time"),
            ("📅 Date",       "date"),
            ("🔍 Google",     "google"),
            ("📺 YouTube",    "youtube"),
            ("🧮 Calculator", "calculator"),
            ("📝 Notepad",    "notepad"),
            ("😄 Joke",       "joke"),
            ("🌤 Weather",    "weather"),
        ]

        grid = ctk.CTkFrame(self, fg_color="transparent")
        grid.pack(fill="x", padx=20)

        for i, (label, intent) in enumerate(cmds):
            btn = ctk.CTkButton(
                grid,
                text=label,
                font=("Helvetica", 12),
                fg_color=NUDE["card2"],
                hover_color=NUDE["card"],
                text_color=NUDE["text"],
                border_width=1,
                border_color=NUDE["border"],
                corner_radius=10,
                height=40,
                command=lambda iv=intent: self._quick_cmd(iv),
            )
            btn.grid(row=i // 4, column=i % 4,
                     padx=4, pady=4, sticky="ew")

        for c in range(4):
            grid.grid_columnconfigure(c, weight=1)

        # ── Bottom stats ──
        stats = ctk.CTkFrame(self, fg_color="transparent")
        stats.pack(fill="x", padx=20, pady=(16, 20))

        count_frame = self._stat_tile(stats, "Commands", "0")
        count_frame.grid(row=0, column=0, padx=(0, 6), sticky="ew")
        self.count_lbl = count_frame._value_label

        self._stat_tile(stats, "Status", "Active",
                        val_color=NUDE["green"]).grid(
            row=0, column=1, padx=6, sticky="ew")

        self.time_tile_val = ctk.StringVar(value="--:--")
        tf = self._stat_tile(stats, "Time", "", var=self.time_tile_val)
        tf.grid(row=0, column=2, padx=(6, 0), sticky="ew")

        for c in range(3):
            stats.grid_columnconfigure(c, weight=1)

    # ── Stat tile ─────────────────────────────────────────────────────────────
    def _stat_tile(self, parent, label, value, val_color=None, var=None):
        frame = ctk.CTkFrame(parent, fg_color=NUDE["card"],
                             corner_radius=10,
                             border_width=1, border_color=NUDE["border"])
        ctk.CTkLabel(frame, text=label,
                     font=("Helvetica", 10),
                     text_color=NUDE["text_lt"]).pack(pady=(10, 2))
        if var:
            lbl = ctk.CTkLabel(frame, textvariable=var,
                               font=("Georgia", 15, "bold"),
                               text_color=val_color or NUDE["text"])
        else:
            lbl = ctk.CTkLabel(frame, text=value,
                               font=("Georgia", 15, "bold"),
                               text_color=val_color or NUDE["text"])
        lbl.pack(pady=(0, 10))
        frame._value_label = lbl
        return frame

    # ── Waveform drawing ──────────────────────────────────────────────────────
    def _draw_idle_wave(self):
        self.canvas.delete("all")
        w = self.canvas.winfo_width() or 440
        heights = [10, 18, 12, 26, 16, 32, 14, 28, 12, 20, 10, 24]
        bar_w, gap, x = 6, 8, (w - len(heights) * (6 + 8)) // 2
        for h in heights:
            cy = 35
            self.canvas.create_rectangle(
                x, cy - h, x + bar_w, cy + h,
                fill=NUDE["border"], outline="", width=0)
            x += bar_w + gap

    def _animate_wave(self):
        if not self.running:
            return
        import math
        self.canvas.delete("all")
        w = self.canvas.winfo_width() or 440
        n = 12
        bar_w, gap = 6, 8
        x = (w - n * (bar_w + gap)) // 2
        t = time.time() * 4
        for i in range(n):
            h = int(18 + 16 * abs(math.sin(t + i * 0.6)))
            cy = 35
            self.canvas.create_rectangle(
                x, cy - h, x + bar_w, cy + h,
                fill=NUDE["accent"], outline="", width=0)
            x += bar_w + gap
        self.after(60, self._animate_wave)

    # ── Clock ─────────────────────────────────────────────────────────────────
    def _update_clock(self):
        self.time_tile_val.set(
            datetime.datetime.now().strftime("%I:%M %p"))
        self.after(30000, self._update_clock)

    # ── Listen toggle ─────────────────────────────────────────────────────────
    def _toggle_listen(self):
        if self.running:
            self.running = False
            self.mic_btn.configure(text="🎤  Start Listening",
                                   fg_color=NUDE["accent"])
            self.status_dot.configure(text="● Idle",
                                      text_color=NUDE["text_lt"])
            self._draw_idle_wave()
        else:
            self.running = True
            self.mic_btn.configure(text="⏹  Stop Listening",
                                   fg_color="#A05030")
            self.status_dot.configure(text="● Listening",
                                      text_color=NUDE["green"])
            self._animate_wave()
            threading.Thread(target=self._listen_loop, daemon=True).start()

    # ── Core listen loop ──────────────────────────────────────────────────────
    def _listen_loop(self):
        self._speak("Listening! How can I help you?")
        while self.running:
            audio = self.voice_in.capture_audio()
            text  = self.recognizer.recognize(audio)

            if text is None:
                self._set_you_said("Didn't catch that, please try again…")
                continue

            self._set_you_said(f'"{text}"')
            intent   = self.processor.get_intent(text)
            response = self.executor.execute(intent)

            if response == "STOP":
                self._speak("Goodbye! Take care!")
                self.after(0, self._toggle_listen)
                break
            else:
                self.cmd_count += 1
                self.after(0, lambda: self.count_lbl.configure(
                    text=str(self.cmd_count)))
                self._speak(response)

    def _quick_cmd(self, intent):
        response = self.executor.execute(intent)
        self.cmd_count += 1
        self.after(0, lambda: self.count_lbl.configure(
            text=str(self.cmd_count)))
        self._set_you_said(f"[Quick: {intent}]")
        self._speak(response)

    # ── FIX: Create a fresh COM instance on each thread + wait for completion ──
    def _speak(self, text):
        self.after(0, lambda: self.resp_label.configure(text=text))

        def speak_in_thread():
            try:
                import pythoncom
                pythoncom.CoInitialize()          # initialise COM for this thread
                spk = win32com.client.Dispatch("SAPI.SpVoice")
                spk.Speak(text)
                spk.WaitUntilDone(-1)             # wait until speech fully finishes
                pythoncom.CoUninitialize()        # clean up COM
            except Exception as e:
                print(f"[Speech Error] {e}")

        threading.Thread(target=speak_in_thread, daemon=True).start()

    def _set_you_said(self, text):
        self.after(0, lambda: self.you_said.configure(text=text))


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = JarvisApp()
    app.mainloop()
