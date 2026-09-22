from auth import show_login_window
from Cal_Report_UI import creat_dialog  
import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

def main():
    root = ttk.Window(themename="cosmo")
    root.withdraw()
    show_login_window(root, lambda: creat_dialog(root))
    root.mainloop()

if __name__ == "__main__":
    main()
