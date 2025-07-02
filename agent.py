import re
import time
import cv2
import mediapipe as mp
from pynput.mouse import Controller, Button
import subprocess

mouse = Controller()
PIXELS_PER_DEGREE = 5

# ----------------------------------------
# Common Mouse Functions
# ----------------------------------------
def move_mouse(direction, degree):
    dx, dy = 0, 0
    pixels = degree * PIXELS_PER_DEGREE

    if direction == "up": dy = -pixels
    elif direction == "down": dy = pixels
    elif direction == "left": dx = -pixels
    elif direction == "right": dx = pixels

    mouse.move(dx, dy)
    print(f"✅ Moved {direction} by {pixels}px")

def click_mouse(type="left", times=1):
    button = {"left": Button.left, "right": Button.right, "middle": Button.middle}.get(type, Button.left)
    mouse.click(button, times)
    print(f"✅ {type.capitalize()} click x{times} done.")

def scroll_mouse(direction, amount):
    mouse.scroll(0, amount if direction == "up" else -amount)
    print(f"✅ Scrolled {direction} by {amount}")

def report_position():
    print(f"📍 Current Mouse Position: {mouse.position}")

def set_position(x, y):
    mouse.position = (x, y)
    print(f"✅ Set Mouse to Position X={x}, Y={y}")

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
    print(f"✅ Dragged {direction} by {pixels}px")

def open_application(app_name: str):
    # Map common app names to executable commands
    app_map = {
        "notepad": "notepad.exe",
        "browser": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
       "edge": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "calculator": "calc.exe",
        "paint": "mspaint.exe",
        "word": "winword.exe",
        "excel": "excel.exe",
        "powerpoint": "powerpnt.exe",
        # Add more mappings as needed
    }
    executable = app_map.get(app_name.lower(), app_name)
    try:
        subprocess.Popen(executable)
        print(f"✅ Opened application: {executable}")
    except Exception as e:
        print(f"❌ Failed to open application '{executable}': {e}")

def handle_command(cmd):
    import pyautogui
    cmd = cmd.lower().strip()

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
    elif match := re.match(r"search\s+(.+)", cmd):
        search_query = match.group(1)
        import webbrowser
        url = f"https://www.google.com/search?q={search_query.replace(' ', '+')}"
        webbrowser.open(url)
        print(f"✅ Opened Google search for: {search_query}")
    elif match := re.match(r"type\s+(.+)", cmd):
        text_to_type = match.group(1)
        # Perform left click at current mouse position
        mouse.click(Button.left, 1)
        import pyautogui
        pyautogui.typewrite(text_to_type)
        print(f"✅ Typed text: {text_to_type}")
    elif cmd == "volume up":
        try:
            subprocess.run(["powershell", "-Command", "(Get-Volume -DeviceId (Get-AudioDevice -Playback).Id).SetMasterVolume((Get-Volume -DeviceId (Get-AudioDevice -Playback).Id).MasterVolume + 0.1)"], check=True)
            print("✅ Volume increased")
        except Exception as e:
            print(f"❌ Failed to increase volume: {e}")
    elif cmd == "volume down":
        try:
            subprocess.run(["powershell", "-Command", "(Get-Volume -DeviceId (Get-AudioDevice -Playback).Id).SetMasterVolume((Get-Volume -DeviceId (Get-AudioDevice -Playback).Id).MasterVolume - 0.1)"], check=True)
            print("✅ Volume decreased")
        except Exception as e:
            print(f"❌ Failed to decrease volume: {e}")
    else:
        print("❌ Unknown command.")

# ----------------------------------------
# Mode 1 - Terminal
# ----------------------------------------
def terminal_mode():
    print("\n🧠 Terminal Mode Active (type 'exit' to quit)")
    while True:
        cmd = input(">> ")
        if cmd.lower() in ['exit', 'quit']:
            print("👋 Exiting Terminal Mode...")
            break
        handle_command(cmd)

# ----------------------------------------
# Main Menu
# ----------------------------------------
def main_menu():
    print("""
🔘 Select Control Mode:
1. Terminal-Based Control
    """)
    choice = input(">> ").strip()
    if choice == '1':
        terminal_mode()
    else:
        print("❌ Invalid choice.")

# ----------------------------------------
# Main
# ----------------------------------------
if __name__ == "__main__":
    main_menu()
