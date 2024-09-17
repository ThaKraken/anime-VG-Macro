import subprocess
import time
import ctypes
import ctypes.wintypes
import json
import pydirectinput
import win32con
import win32gui
import cv2
import numpy as np
from PIL import ImageGrab

class Macro:
    SIZE_FILE_PATH = 'window_size.json'

    def __init__(self, window_name):
        self.window_name = window_name
        self.hwnd = self._get_window_handle()
        self.stored_width, self.stored_height = self._load_window_size()
        self.first_time_open = True
        self._focus_window()

    def _capture_window(self):
        left, top, right, bottom = win32gui.GetWindowRect(self.hwnd)
        screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
        screenshot = np.array(screenshot)
        return cv2.cvtColor(screenshot, cv2.COLOR_RGB2GRAY)

    def _match_template(self, screenshot, template, threshold=0.6):
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        if max_val >= threshold:
            print(f"Match found at {max_loc} with confidence {max_val}")
            return max_loc
        print("No match found")
        return None

    def _get_window_handle(self):
        hwnd = win32gui.FindWindow(None, self.window_name)
        if not hwnd:
            raise RuntimeError(f"Window with name '{self.window_name}' not found.")
        return hwnd

    def _load_window_size(self):
        try:
            with open(self.SIZE_FILE_PATH, 'r') as f:
                size_info = json.load(f)
            return size_info['width'], size_info['height']
        except (FileNotFoundError, KeyError):
            return None, None

    def _resize_window(self, width, height):
        win32gui.SetWindowPos(self.hwnd, win32con.HWND_TOP, 0, 0, width, height,
                              win32con.SWP_NOMOVE | win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER)

    def ensure_window_size(self):
        current_width, current_height = win32gui.GetWindowRect(self.hwnd)[2:4]
        if (current_width, current_height) != (self.stored_width, self.stored_height):
            self._resize_window(self.stored_width, self.stored_height)

    def _focus_window(self):
        """Bring the Roblox window to the foreground and focus it."""
        win32gui.ShowWindow(self.hwnd, win32con.SW_RESTORE)  # Restore window if minimized
        win32gui.SetForegroundWindow(self.hwnd)  # Bring window to the front

    def _relative_to_absolute(self, rel_x, rel_y):
        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.wintypes.LONG), ("y", ctypes.wintypes.LONG)]
        point = POINT(rel_x, rel_y)
        ctypes.windll.user32.ClientToScreen(self.hwnd, ctypes.byref(point))
        return point.x, point.y

    def _move_click(self, rel_x, rel_y, double_click=False, crash_check=True):
        if crash_check:
            self._check_and_handle_crash()
        abs_x, abs_y = self._relative_to_absolute(rel_x, rel_y)
        pydirectinput.moveTo(abs_x, abs_y)
        time.sleep(0.1)
        pydirectinput.move(1, 1)
        time.sleep(0.1)
        pydirectinput.click()
        if double_click:
            time.sleep(0.1)
            pydirectinput.click()

    def _scroll(self, delta):
        self._check_and_handle_crash()
        time.sleep(0.1)
        ctypes.windll.user32.mouse_event(0x0800, 0, 0, delta, 0)
        self._check_and_handle_crash()

    def _handle_crash(self):
        if not self._check_internet():
            print("No internet connection. Trying to reconnect...")
            self._connect_to_network()
        print("Handling crash...")
        self._move_click(550, 400, crash_check=False)
        time.sleep(15)
        self._move_click(110, 290)
        time.sleep(2)
        pydirectinput.keyDown('w')
        time.sleep(1)
        pydirectinput.keyUp('w')
        pydirectinput.keyDown('d')
        time.sleep(1)
        pydirectinput.keyUp('d')
        pydirectinput.keyDown('w')
        time.sleep(3.5)
        pydirectinput.keyUp('w')
        pydirectinput.keyDown('d')
        time.sleep(1)
        pydirectinput.keyUp('d')
        self._move_click(450, 450)
        self._move_click(575, 470)
        self.first_time_open = True
        time.sleep(10)
        self.setup()
        self.main_loop()

    def _check_and_handle_crash(self):
        screenshot = self._capture_window()
        crash_image = cv2.imread('crash_panel.png', cv2.IMREAD_GRAYSCALE)
        result = self._match_template(screenshot, crash_image)
        if result is not None:
            x, y = result
            self._move_click(x, y, crash_check=False)
            self._handle_crash()
            return True
        return False

    def _check_internet(self):
        try:
            subprocess.check_call(["ping", "-n", "1", "8.8.8.8"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            return True
        except subprocess.CalledProcessError:
            return False

    def _connect_to_network(self):
        networks = subprocess.check_output(["netsh", "wlan", "show", "profiles"]).decode('utf-8')
        networks = [line for line in networks.split('\n') if "All User Profile" in line]
        networks = [line.split(":")[1].strip() for line in networks]
        
        if not networks:
            print("No saved networks found.")
            while not self._check_internet():
                print("Waiting for internet connection...")
                time.sleep(10)
            return
        
        for network in networks:
            print(f"Trying to connect to {network}...")
            subprocess.call(["netsh", "wlan", "connect", "name=" + network])
            time.sleep(10)
            if self._check_internet():
                print(f"Connected to {network}")
                return
        
        print("Failed to connect to any saved network. Waiting for connection...")
        while not self._check_internet():
            print("Waiting for internet connection...")
            time.sleep(10)

    def open_settings_and_scroll(self):
        self._move_click(25, 620, double_click=False)
        time.sleep(0.1)
        self._move_click(440, 400)
        self._scroll(-100)
        self._scroll(-20)
        self.first_time_open = False

    def click_start(self):
        self._move_click(425, 130, double_click=True)

    def click_teleport(self):
        self._move_click(600, 470, double_click=True)

    def close_settings(self):
        self._move_click(645, 135, double_click=True)

    def select_unit(self):
        self._move_click(375, 580)

    def place_unit(self, offset_x=0):
        self._move_click(375 + offset_x, 475, double_click=True)

    def place_unit_with_delay(self, unit_number, spacing_x):
        self.select_unit()
        self.place_unit(unit_number * spacing_x)
        time.sleep(20)

    def place_multiple_units(self, num_units, spacing_x):
        for i in range(num_units):
            self.place_unit_with_delay(i, spacing_x)

    def upgrade_units(self, num_units, spacing_x):
        self._check_and_handle_crash()
        time.sleep(20)
        for i in range(num_units):
            self._move_click(375 + i * spacing_x, 475)
            self._move_click(225, 380)
            self._move_click(250, 250)

    def get_rewards(self):
        for _ in range(3):
            pydirectinput.click()
            self._check_and_handle_crash()
            time.sleep(1.5)

    def click_replay(self):
        self._move_click(520, 460, double_click=True)

    def setup(self):
        if self.first_time_open:
            self.open_settings_and_scroll()
            self.click_teleport()
            self.close_settings()

    def main_loop(self):
        while True:
            self.click_start()
            time.sleep(20)
            self.place_multiple_units(3, 80)
            for _ in range(9):
                self.upgrade_units(3, 80)
            time.sleep(15)
            self.get_rewards()
            self.click_replay()

if __name__ == "__main__":
    time.sleep(1)
    macro = Macro("Roblox")
    macro.ensure_window_size()
    macro.setup()
    while True:
        macro.main_loop()
