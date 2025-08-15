import os

def get_current_directory() -> str:
    return os.path.dirname(os.path.realpath(__file__))

def get_capture_directory(current_directory: str, capture_dname: str = "output") -> str:
    return os.path.join(current_directory, capture_dname)

def create_capture_directory(capture_directory_path: str) -> None:
    os.makedirs(capture_directory_path, exist_ok=True)
