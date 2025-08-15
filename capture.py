from PIL import Image, ImageGrab
import time
import pyautogui
import os
import random
import threading
import queue

def capture(file_name: str, bbox=None):
    screenshot = ImageGrab.grab(bbox=bbox)
    screenshot.save(file_name, format="JPEG", subsampling=0, quality=95)
    screenshot.close()

def capture_macro(
    capture_directory: str,
    start_page_no: int,
    end_page_no: int,
    capture_interval: float,
    capture_area,
    cancel_event: threading.Event,
    progress_queue: queue.Queue,
) -> None:
    try:
        for i in range(5, 0, -1):
            if cancel_event.is_set():
                progress_queue.put("cancelled")
                return
            progress_queue.put(f"{i}초 후에 캡처가 시작됩니다...")
            time.sleep(1)

        for page_no in range(start_page_no, end_page_no + 1):
            if cancel_event.is_set():
                progress_queue.put("cancelled")
                return

            progress_queue.put(f"{page_no} / {end_page_no} 캡처 중...")
            formatted_page_no = f"{page_no:04}"
            capture(
                file_name=os.path.join(capture_directory, f"{formatted_page_no}.jpeg"),
                bbox=capture_area,
            )

            if page_no < end_page_no:
                pyautogui.keyDown("right")
                time.sleep(0.1)
                pyautogui.keyUp("right")
                time.sleep(capture_interval + random.uniform(0, 2))

        progress_queue.put("done")
    except Exception as e:
        progress_queue.put(f"Error: {e}")
