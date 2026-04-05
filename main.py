# main.py
import platform
import tkinter as tk
from ui import QuantitativeUI


def _maximize_window(root: tk.Tk) -> None:
    """플랫폼에 맞게 창을 최대화합니다."""
    system = platform.system()
    if system == "Windows":
        root.state("zoomed")
    elif system == "Darwin":  # macOS
        root.attributes("-zoomed", True)
    else:  # Linux
        root.attributes("-zoomed", True)


if __name__ == "__main__":
    root = tk.Tk()
    _maximize_window(root)
    app = QuantitativeUI(root)
    root.mainloop()
