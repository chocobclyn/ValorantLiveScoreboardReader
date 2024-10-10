import tkinter as tk
import subprocess
import threading
import time
import os
import sys

VAL_SCOREBOARD_SCRIPT = "ValScoreboard.py"
VAL_WS_SCRIPT = "ValWS.py"

running = False

def check_python_installed():
    try:
        subprocess.check_call(['python', '--version'])
        return True
    except subprocess.CalledProcessError:
        return False

def install_dependencies():
    dependencies = ["paddleocr", "opencv-python", "mss", "openpyxl", "numpy"]
    for dep in dependencies:
        subprocess.check_call([sys.executable, "-m", "pip", "install", dep])

def install_paddlepaddle():
    subprocess.call([sys.executable, "-m", "pip", "install", "paddlepaddle"])

def install_visual_cpp():
    subprocess.call('start https://aka.ms/vs/17/release/vc_redist.x64.exe', shell=True)

def run_scripts(interval):
    global running
    while running:
        subprocess.run([sys.executable, VAL_SCOREBOARD_SCRIPT])
        subprocess.run([sys.executable, VAL_WS_SCRIPT])

        time.sleep(interval)

def start_scripts():
    global running
    if not running:
        running = True
        interval = int(interval_var.get())  # Get interval from dropdown
        threading.Thread(target=run_scripts, args=(interval,)).start()

def stop_scripts():
    global running
    running = False

window = tk.Tk()
window.title("Valorant Scoreboard + WS Runner")

interval_var = tk.StringVar(window)
interval_var.set("1")

tk.Label(window, text="Select Interval (seconds):").pack(pady=10)
interval_dropdown = tk.OptionMenu(window, interval_var, "1", "2", "3", "5", "10")
interval_dropdown.pack(pady=10)

start_button = tk.Button(window, text="Start", command=start_scripts, width=20, bg="green", fg="white")
start_button.pack(pady=10)

stop_button = tk.Button(window, text="Stop", command=stop_scripts, width=20, bg="red", fg="white")
stop_button.pack(pady=10)

install_button = tk.Button(window, text="Install Dependencies", command=install_dependencies, width=20)
install_button.pack(pady=10)

paddle_button = tk.Button(window, text="Install PaddlePaddle", command=install_paddlepaddle, width=20)
paddle_button.pack(pady=10)

visual_cpp_button = tk.Button(window, text="Install Visual C++", command=install_visual_cpp, width=20)
visual_cpp_button.pack(pady=10)

window.mainloop()
