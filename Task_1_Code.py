import speech_recognition as sr
import pyttsx3
import webbrowser
import secrets
import math
import operator
import subprocess
import sys
import platform
import time
import ast

# Optional clipboard support
try:
    import pyperclip
    CLIPBOARD_AVAILABLE = True
except Exception:
    CLIPBOARD_AVAILABLE = False

# ---------------- Text-to-speech setup ----------------
engine = pyttsx3.init()
engine.setProperty("rate", 160)
voices = engine.getProperty("voices")
if voices:
    preferred = None
    for v in voices:
        name = v.name.lower()
        if "female" in name or "zira" in name or "susan" in name or "karen" in name:
            preferred = v.id
            break
    if not preferred:
        preferred = voices[0].id
    engine.setProperty("voice", preferred)

def respond(text, speak=True):
    print("VoiceBuddy:", text)
    if speak:
        engine.say(text)
        engine.runAndWait()

# -------------- Recognizer setup (FIXED) --------------
recognizer = sr.Recognizer()
recognizer.non_speaking_duration = 0.2   # seconds of non-speaking used to consider phrase finished
recognizer.pause_threshold = 0.6         # seconds of silence before phrase is considered complete
recognizer.dynamic_energy_threshold = True
recognizer.energy_threshold = 300        # initial energy threshold
recognizer.phrase_threshold = 0.1

PHRASE_TIME_LIMIT = 8

def calibrate_microphone(source, duration=1.5):
    """
    Calibrate ambient noise. This function is called once at startup and on demand.
    Wrapped in try/except to avoid crashes if microphone isn't available.
    """
    try:
        # Ensure safe timing settings before calling adjust_for_ambient_noise
        if recognizer.pause_threshold < recognizer.non_speaking_duration:
            # fix to satisfy assertion
            recognizer.non_speaking_duration = max(0.0, recognizer.pause_threshold * 0.4)

        print("Calibrating for ambient noise (keep quiet)...")
        # adjust_for_ambient_noise will set recognizer.energy_threshold automatically
        recognizer.adjust_for_ambient_noise(source, duration=duration)
        # after calibration, ensure invariant again (some versions may change thresholds)
        if recognizer.pause_threshold < recognizer.non_speaking_duration:
            recognizer.non_speaking_duration = min(recognizer.non_speaking_duration, recognizer.pause_threshold)
        print(f"Calibration complete. energy_threshold = {recognizer.energy_threshold:.1f}")
        respond("Calibration complete.", speak=False)
        return True
    except AssertionError as ae:
        print("Calibration assertion error:", ae)
        recognizer.non_speaking_duration = 0.15
        recognizer.pause_threshold = max(0.4, recognizer.non_speaking_duration + 0.2)
        recognizer.energy_threshold = 400
        print("Applied conservative fallback audio settings.")
        return False
    except Exception as e:
        print("Calibration failed:", e)
        return False

# -------------- Safe math evaluator --------------
class _MathEvaluator(ast.NodeVisitor):
    def visit(self, node):
        return super().visit(node)

    def visit_Expression(self, node):
        return self.visit(node.body)

    def visit_BinOp(self, node):
        left = self.visit(node.left)
        right = self.visit(node.right)
        op_type = type(node.op).__name__
        if op_type == "Add":
            return operator.add(left, right)
        if op_type == "Sub":
            return operator.sub(left, right)
        if op_type == "Mult":
            return operator.mul(left, right)
        if op_type == "Div":
            return operator.truediv(left, right)
        if op_type == "Pow":
            return operator.pow(left, right)
        raise ValueError(f"Unsupported operator {op_type}")

    def visit_UnaryOp(self, node):
        op_type = type(node.op).__name__
        if op_type == "USub":
            return - self.visit(node.operand)
        raise ValueError(f"Unsupported unary op {op_type}")

    def visit_Num(self, node):
        return node.n

    def visit_Constant(self, node):
        if isinstance(node.value, (int, float)):
            return node.value
        raise ValueError("Only int/float constants are allowed")

    def visit_Call(self, node):
        raise ValueError("Function calls are not allowed")

    def generic_visit(self, node):
        raise ValueError(f"Unsupported expression element: {type(node).__name__}")

def safe_eval_math(expr: str):
    try:
        parsed = ast.parse(expr, mode="eval")
        evaluator = _MathEvaluator()
        val = evaluator.visit(parsed)
        return val
    except Exception as e:
        raise ValueError(f"Invalid expression: {e}")

# -------------- Utilities & Commands --------------
JOKES = [
    "Why don't scientists trust atoms? Because they make up everything.",
    "I told my computer I needed a break. It said 'no problem — I'll go to sleep.'",
    "Why did the scarecrow win an award? He was outstanding in his field.",
    "Why did the math book look sad? Because it had too many problems.",
]

def tell_joke():
    joke = secrets.choice(JOKES)
    respond(joke)

def open_calculator():
    system = platform.system()
    try:
        if system == "Windows":
            subprocess.Popen(["calc.exe"])
        elif system == "Darwin":
            subprocess.Popen(["open", "-a", "Calculator"])
        else:
            for cmd in (["gnome-calculator"], ["kcalc"], ["galculator"], ["xcalc"]):
                try:
                    subprocess.Popen(cmd)
                    return True
                except Exception:
                    continue
            respond("Could not open calculator automatically on this system.")
            return False
        respond("Opening calculator.")
        return True
    except Exception as e:
        respond(f"Failed to open calculator: {e}")
        return False

_last_generated_text = None
def copy_last_to_clipboard():
    global _last_generated_text
    if not CLIPBOARD_AVAILABLE:
        respond("Clipboard support not installed.")
        return False
    if not _last_generated_text:
        respond("Nothing to copy yet.")
        return False
    pyperclip.copy(_last_generated_text)
    respond("Copied to clipboard.")
    return True

# -------------- Command processing --------------
def process_command(command_text: str):
    global _last_generated_text
    text = command_text.lower().strip()
    print("Heard (processed):", text)

    if any(text.startswith(w) for w in ("hello", "hi", "hey buddy", "hey")):
        respond(secrets.choice(["Hi there!", "Hello! Ready when you are.", "Hey! How can I help?"]))
        return False

    if any(k in text for k in ("goodbye", "exit", "quit", "stop")):
        respond("Goodbye! Shutting down.")
        return True

    if "recalibrate" in text or "calibrate" in text or "adjust for noise" in text:
        respond("Recalibrating ambient noise. Please be quiet for a moment.")
        try:
            with sr.Microphone() as source:
                calibrate_microphone(source, duration=1.5)
        except Exception as e:
            respond(f"Recalibration failed: {e}")
        return False

    if "time" in text:
        t = time.localtime()
        respond(f"It's {time.strftime('%I:%M %p', t)}")
        return False
    if "date" in text:
        respond(f"Today is {time.strftime('%A, %B %d, %Y', time.localtime())}")
        return False

    if text.startswith("search for ") or text.startswith("look up ") or text.startswith("search "):
        query = text
        for prefix in ("search for ", "look up ", "search "):
            if query.startswith(prefix):
                query = query[len(prefix):]
        query = query.strip()
        if not query:
            respond("What should I search for?")
            return False
        respond(f"Searching for {query}.")
        url = "https://www.google.com/search?q=" + webbrowser.quote(query)
        webbrowser.open(url)
        _last_generated_text = f"Search: {query} -> {url}"
        return False

    if "joke" in text or "tell me a joke" in text:
        tell_joke()
        _last_generated_text = "joke"
        return False

    if text.startswith("repeat after me "):
        phrase = command_text[len("repeat after me "):].strip()
        if not phrase:
            respond("Say what should I repeat?")
            return False
        respond(phrase)
        _last_generated_text = phrase
        return False

    if "copy that" in text or "copy this" in text:
        copy_last_to_clipboard()
        return False

    if "open calculator" in text or "calculator" == text.strip():
        open_calculator()
        return False

    if text.startswith("calculate ") or text.startswith("what is ") or text.startswith("what's "):
        expr = command_text
        for p in ("calculate ", "what is ", "what's "):
            if expr.lower().startswith(p):
                expr = expr[len(p):]
                break
        expr = expr.strip()
        replacements = {
            " plus ": " + ",
            " minus ": " - ",
            " times ": " * ",
            " multiplied by ": " * ",
            " divided by ": " / ",
            " over ": " / ",
            " power of ": " ** ",
            " to the power of ": " ** ",
        }
        low = expr.lower()
        for k, v in replacements.items():
            low = low.replace(k, v)
        low = low.replace(" equals", "").replace(" equal", "")
        try:
            val = safe_eval_math(low)
            respond(f"The answer is {val}")
            _last_generated_text = str(val)
        except Exception:
            respond("Sorry, I couldn't compute that. Try a simpler expression.")
        return False

    if len(text) < 60:
        if "how are you" in text:
            respond("I'm a program, but thanks for asking — I'm ready to help.")
            return False

    respond("Sorry, I didn't understand that. You can say 'tell me a joke', 'calculate 3 plus 4', 'search for cats', or 'open calculator'. Say 'recalibrate' if I'm not hearing you well.")
    return False

# -------------- Listening loop (FIXED) --------------
def listen_for_command(timeout=6, phrase_time_limit=PHRASE_TIME_LIMIT):
    """
    Listen from microphone and return recognized text (str) or None.
    NOTE: We DO NOT run adjust_for_ambient_noise here (calibration only at startup or on demand).
    """
    try:
        with sr.Microphone() as source:
            print("Listening... (speak now)")
            audio = None
            try:
                audio = recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
            except sr.WaitTimeoutError:
                print("No speech detected (timeout).")
                return None
    except Exception as e:
        print("Microphone access error:", e)
        return None

    try:
        recognized = recognizer.recognize_google(audio)
        print("Recognized (raw):", recognized)
        return recognized
    except sr.UnknownValueError:
        print("Could not understand audio.")
        return None
    except sr.RequestError as e:
        print("Speech recognition request failed:", e)
        try:
            recognized = recognizer.recognize_sphinx(audio)
            return recognized
        except Exception:
            return None

# -------------- Main loop ----------------
def main_loop():
    respond("VoiceBuddy activated. Say 'Hey Buddy' or 'Hello' to wake me.")
    try:
        with sr.Microphone() as source:
            ok = calibrate_microphone(source, duration=1.5)
            if not ok:
                respond("Calibration couldn't complete; using fallback audio settings.", speak=False)
    except Exception as e:
        print("Microphone calibration failed at startup:", e)
        respond("Microphone calibration failed. You can try saying 'recalibrate' to attempt again.", speak=False)

    running = True
    while running:
        print("Waiting for wake word ('hey buddy' or 'hello')...")
        text = listen_for_command(timeout=8, phrase_time_limit=5)
        if not text:
            continue
        txt = text.lower()
        if "hey buddy" in txt or txt.strip().startswith("hello") or txt.strip().startswith("hi"):
            respond("Yes? I'm listening.")
            cmd_text = listen_for_command(timeout=10, phrase_time_limit=PHRASE_TIME_LIMIT)
            if not cmd_text:
                respond("I didn't catch that. Could you repeat?")
                cmd_text = listen_for_command(timeout=6, phrase_time_limit=PHRASE_TIME_LIMIT)
                if not cmd_text:
                    respond("Still didn't catch you. Say 'recalibrate' if I'm having trouble hearing.")
                    continue
            should_exit = process_command(cmd_text)
            if should_exit:
                running = False
                break
        else:
            if len(txt.split()) <= 6 and any(k in txt for k in ("time", "date", "search", "joke", "calculate", "calculator")):
                should_exit = process_command(text)
                if should_exit:
                    running = False
                    break
            else:
                print("No wake word detected; ignoring.")
                continue
    respond("VoiceBuddy has shut down. Bye!")
    return

if __name__ == "__main__":
    try:
        main_loop()
    except KeyboardInterrupt:
        respond("Interrupted — stopping.")
        sys.exit(0)
