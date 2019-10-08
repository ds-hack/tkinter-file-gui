import tkinter as tk
from tkinter_gui import MainFrame

if __name__ == "__main__":
    APP_TITLE = "DS Hack GUI Application"

    root = tk.Tk()
    root.geometry('680x220+300+200')
    root.title(APP_TITLE)
    # メインウィンドウ内全体に白背景のフレームを生成
    MainFrame(root, app_title=APP_TITLE, bg="white").pack(side="top", 
                                                          fill="both",
                                                          expand=True)
    root.mainloop()