import re
import time
import cv2
import mediapipe as mp
import speech_recognition as sr
from pynput.mouse import Controller, Button
import subprocess
import pyttsx3
import pywhatkit
import pyautogui
import socket
import random
import psutil
import datetime
import os

mouse = Controller()
PIXELS_PER_DEGREE = 5
import threading

engine = pyttsx3.init()
current_mode = "terminal"  # default

def speak(text: str):
    print(f"🔊 {text}")
    def run_speech():
        engine.say(text)
        engine.runAndWait()
    threading.Thread(target=run_speech, daemon=True).start()

def say_and_print(msg):
    print(msg)
    speak(msg)

def get_voice_input(prompt="Speak now..."):
    speak(prompt)
    recognizer = sr.Recognizer()
    mic = sr.Microphone()
    with mic as source:
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
    try:
        return recognizer.recognize_google(audio)
    except:
        return ""

def play_youtube_song(song_name=None):
    if not song_name:
        song_name = input("🎵 Enter song name: ")
    say_and_print(f"Playing {song_name} on YouTube...")
    try:
        pywhatkit.playonyt(song_name)
    except:
        say_and_print("❌ Failed to play song. Check your internet connection.")

# ------------------ Mouse Functions ------------------
def move_mouse(direction, degree):
    dx, dy = 0, 0
    pixels = degree * PIXELS_PER_DEGREE
    if direction == "up": dy = -pixels
    elif direction == "down": dy = pixels
    elif direction == "left": dx = -pixels
    elif direction == "right": dx = pixels
    mouse.move(dx, dy)
    say_and_print(f"Moved {direction} by {pixels} pixels")

def click_mouse(type="left", times=1):
    button = {"left": Button.left, "right": Button.right, "middle": Button.middle}.get(type, Button.left)
    mouse.click(button, times)
    say_and_print(f"{type.capitalize()} click {times} times done")

def scroll_mouse(direction, amount):
    mouse.scroll(0, amount if direction == "up" else -amount)
    say_and_print(f"Scrolled {direction} by {amount}")

def report_position():
    say_and_print(f"Current Mouse Position: {mouse.position}")

def set_position(x, y):
    mouse.position = (x, y)
    say_and_print(f"Set Mouse to Position X={x}, Y={y}")

def drag_mouse(direction, degree):
    dx, dy = 0, 0
    pixels = degree * PIXELS_PER_DEGREE
    if direction == "up": dy = -pixels
    elif direction == "down": dy = pixels
    elif direction == "left": dx = -pixels
    elif direction == "right": dx = pixels
    mouse.press(Button.left)
    mouse.move(dx, dy)
    mouse.release(Button.left)
    say_and_print(f"Dragged {direction} by {pixels}px")

def open_application(app_name: str):
    app_map = {
        "notepad": "notepad.exe",
        "browser": r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        "edge": r"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
        "calculator": "calc.exe",
        "paint": "mspaint.exe",
        "word": "winword.exe",
        "excel": "excel.exe",
        "powerpoint": "powerpnt.exe",
    }
    executable = app_map.get(app_name.lower(), app_name)
    try:
        subprocess.Popen(executable)
        say_and_print(f"Opened application: {executable}")
    except Exception as e:
        say_and_print(f"Failed to open application '{executable}': {e}")

def help_command():
    commands = """
Available Commands:

Mouse:
- up/down/left/right mouse {degree} degree
- drag mouse up/down/left/right {degree}
- one click mouse, double click mouse
- right click mouse, middle click mouse
- scroll up/down {amount}
- mouse position
- set mouse {x} {y}

Apps:
- open notepad, calculator, paint, word, excel, browser
- open site [url]
- open explorer
- play song on youtube

System:
- lock, shutdown, restart
- volume up, volume down
- screenshot
- battery status
- date and time
- ip address
- system info

Text/Voice:
- type [text]
- search [query]
- note [your note]
- read note

Windows:
- minimize all, maximize window, minimize window, close window

Fun:
- joke, quote

General:
- help
- exit
"""
    say_and_print(commands)

# ------------------ Command Handler ------------------
import io
import sys

def get_command_response(cmd):
    cmd = cmd.lower().strip()
    output = io.StringIO()
    sys_stdout = sys.stdout
    sys.stdout = output
    try:
        if match := re.match(r"(up|down|left|right)\s+mouse\s+(\d+)\s*degree", cmd):
            move_mouse(match.group(1), int(match.group(2)))
        elif cmd == "one click mouse":
            click_mouse("left", 1)
        elif cmd == "double click mouse":
            click_mouse("left", 2)
        elif cmd == "right click mouse":
            click_mouse("right", 1)
        elif cmd == "middle click mouse":
            click_mouse("middle", 1)
        elif match := re.match(r"scroll\s+(up|down)\s+(\d+)", cmd):
            scroll_mouse(match.group(1), int(match.group(2)))
        elif cmd == "mouse position":
            report_position()
        elif match := re.match(r"set mouse (\d+)\s+(\d+)", cmd):
            set_position(int(match.group(1)), int(match.group(2)))
        elif match := re.match(r"drag mouse (up|down|left|right) (\d+)", cmd):
            drag_mouse(match.group(1), int(match.group(2)))
        elif match := re.match(r"open\s+(\w+)", cmd):
            open_application(match.group(1))
        elif match := re.match(r"open site (.+)", cmd):
            webbrowser.open(f"https://{match.group(1)}")
            say_and_print(f"Opening website: {match.group(1)}")
        elif match := re.match(r"search\s+(.+)", cmd):
            url = f"https://www.google.com/search?q={match.group(1).replace(' ', '+')}"
            webbrowser.open(url)
            say_and_print(f"Opened Google search for: {match.group(1)}")
        elif match := re.match(r"type\s+(.+)", cmd):
            mouse.click(Button.left, 1)
            pyautogui.typewrite(match.group(1))
            say_and_print(f"Typed text: {match.group(1)}")
        elif cmd == "play song on youtube":
            if current_mode == "voice":
                song = get_voice_input("Which song should I play?")
            else:
                song = input("🎵 Song name: ")
            if song:
                play_youtube_song(song)
            else:
                say_and_print("Sorry, I didn't catch the song name.")
        elif cmd == "volume up" or cmd == "volume down":
            try:
                subprocess.run(["sndvol.exe"], check=True)
                say_and_print("Volume control not supported directly. Opened volume mixer.")
            except Exception as e:
                say_and_print(f"Failed to open volume mixer: {e}")
        elif cmd == "lock":
            subprocess.run(["rundll32.exe", "user32.dll,LockWorkStation"])
            say_and_print("System locked.")
        elif cmd == "shutdown":
            subprocess.run(["shutdown", "/s", "/t", "0"])
            say_and_print("System shutting down.")
        elif cmd == "restart":
            subprocess.run(["shutdown", "/r", "/t", "0"])
            say_and_print("System restarting.")
        elif cmd == "help":
            help_command()
        elif cmd == "exit":
            say_and_print("Exiting program...")
            exit()
        else:
            say_and_print("Unknown command. Please try again.")
    finally:
        sys.stdout = sys_stdout
    return output.getvalue()

# ------------------ Modes ------------------
def terminal_mode():
    global current_mode
    current_mode = "terminal"
    speak("Terminal mode activated. Type your command.")
    while True:
        cmd = input(">> ")
        if cmd.lower() in ['exit', 'quit']:
            speak("Exiting terminal mode.")
            break
        handle_command(cmd)

def voice_mode():
    global current_mode
    current_mode = "voice"
    recognizer = sr.Recognizer()
    mic = sr.Microphone()
    speak("Voice mode activated. Please speak your command.")
    with mic as source:
        recognizer.adjust_for_ambient_noise(source, duration=1)
    while True:
        try:
            with mic as source:
                print("🎙 Speak your command:")
                audio = recognizer.listen(source)
                cmd = recognizer.recognize_google(audio)
                print(f">> You said: {cmd}")
                if "exit" in cmd.lower():
                    speak("Exiting voice mode.")
                    break
                handle_command(cmd)
        except sr.UnknownValueError:
            speak("Could not understand. Please speak clearly.")
        except sr.RequestError:
            speak("Google service error. Check your internet connection.")
        except sr.WaitTimeoutError:
            speak("Timeout. Please try again.")

# ------------------ Main Menu ------------------
def main_menu():
    speak("Welcome sir! Please select a control mode.")
    print("""
🔘 Select Control Mode:
1. Terminal-Based Control
2. Voice-Based Control
""")
    choice = input(">> ").strip()
    if choice == '1':
        terminal_mode()
    elif choice == '2':
        voice_mode()
    else:
        speak("Invalid choice. Please restart the program.")

if __name__ == "__main__":
    main_menu()
