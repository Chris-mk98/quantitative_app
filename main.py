# main.py
import tkinter as tk
from ui import QuantitativeUI

if __name__ == "__main__":
    root = tk.Tk()
    root.state('zoomed')
    app = QuantitativeUI(root)
    root.mainloop()
