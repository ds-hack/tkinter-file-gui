import os
import threading
import tkinter as tk
import tkinter.filedialog
from tkinter import ttk
from logic import execute_logic

class MainFrame(tk.Frame):

    def __init__(self, parent, app_title, *args, **kwargs):
        """
        メインウィンドウ内のフレームを生成するコンストラクタ
        下記の処理を実行する
        ・フレームの設定
        ・フレーム内に含まれるヴィジェット（ラベル・ボタンなど）の配置

        parameters
        ----------
        parent : tk.Tk
            フレームの親となるメインウィンドウ
        app_title : app_title
            アプリケーションのタイトル（GUIに表示される）
        *args : variable arguments
            Frameオブジェクトの初期化引数
        **kwargs : variable arguments
            Frameオブジェクトの初期化引数
        """
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent

        # titleラベル
        self.gui_title_label = tk.Label(self,
                                        text=app_title,
                                        font=('メイリオ', 32),
                                        bg='white',
                                        fg='limegreen')
        # titleラベルの設置（今回はplaceメソッドを使う）
        self.gui_title_label.place(x=80, y=30)

        # file entry (ファイルパスの表示部分)
        self.file_path = tk.StringVar()
        self.file_entry = tk.Entry(self,
                                   textvariable=self.file_path,
                                   font=('メイリオ', 10),
                                   width=55,
                                   bd=3)
        self.file_entry.place(x=80, y=110)

        # file選択ボタン
        self.file_button_img = tk.PhotoImage(file='./icons/folder.png')
        self.file_button = tk.Button(self,
                                     image=self.file_button_img,
                                     cursor='hand2')
        # ボタンがクリックされた際のコールバック関数をバインド（割り当て）
        self.file_button.bind('<Button-1>', self.file_button_clicked)
        self.file_button.place(x=540, y=107)
        # ツールチップの生成
        self.file_menu_ttp = CreateToolTip(self.file_button, "入力ファイルを選択")

        # 設定ボタン
        # 今回設定ウィンドウは実装しないので、コールバック関数は設定しない。
        self.setting_button_img = tk.PhotoImage(file='./icons/settings.png')
        self.setting_button = tk.Button(self,
                                        image=self.setting_button_img,
                                        cursor='hand2')
        self.setting_button.place(x=585, y=107)
        self.file_menu_ttp = CreateToolTip(self.setting_button, "設定画面を開く")

        # アプリケーション実行ボタン
        self.execute_button = tk.Button(self,
                                        text='アプリケーション実行',
                                        font=('メイリオ', 10),
                                        relief='raised',
                                        cursor='hand2',
                                        fg='limegreen',
                                        bg='snow',
                                        highlightbackground='limegreen',
                                        width=20)
        # マウスホバーで色を変化させる
        self.execute_button.bind('<Enter>', self.hover_enter)
        self.execute_button.bind('<Leave>', self.hover_leave)
        self.execute_button.bind('<Button-1>', self.execute_button_clicked)
        self.execute_button.place(x=250, y=160)


    def file_button_clicked(self, event):
        # xlsまたはxlsx形式のファイルを選択させる
        filetypes = [("Excelブック(.xlsx)","*.xlsx"),
                     ("Excelブック(.xls)",".xls")]
        # ファイル選択ダイアログの初期フォルダ
        initialdir = os.path.abspath(os.path.dirname(__file__))
        path = tk.filedialog.askopenfilename(filetypes=filetypes,
                                             initialdir=initialdir)
        self.file_path.set(path)

    def hover_enter(self, event):
        # (注)mac, linuxではbuttonのbackgroundオプションが働かない
        self.execute_button.configure(fg='snow', bg='limegreen')

    def hover_leave(self, event):
        self.execute_button.configure(fg='limegreen', bg='snow')

    def execute_button_clicked(self, event):
        # スレッドの開始(GUIの描画更新のため)
        file_path = self.file_path.get()
        # アプリケーションの進捗確認（プログラスバー）ウィンドウの生成
        progress_window = tk.Toplevel(master=self.parent)
        progress_window.geometry('640x200+400+300')
        progress_window.title("アプリ実行進捗状況")
        # MainFrame同様、フレーム内部のヴィジェットとコールバック関数をクラス化
        self.progress_frame = ProgressFrame(progress_window,
                                            file_path=file_path,
                                            bg="white")
        self.progress_frame.pack(side="top", fill="both", expand=True)
        # メインロジックの実行
        self.progress_frame.execute_logic()
        progress_window.mainloop()


class ProgressFrame(tk.Frame):

    def __init__(self, parent, file_path, *args, **kwargs):
        """
        プログレスウィンドウ内のフレームを生成するコンストラクタ
        下記の処理を実行する
        ・フレームの設定
        ・フレーム内に含まれるヴィジェット（ラベル・ボタンなど）の配置

        parameters
        ----------
        parent : tk.Tk
            フレームの親となるメインウィンドウ
        file_path : str
            メインウィンドウで選択されたファイルのパス
        *args : variable arguments
            Frameオブジェクトの初期化引数
        **kwargs : variable arguments
            Frameオブジェクトの初期化引数
        """
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.file_path = file_path

        # 進捗状況ラベル
        self.progress_message = tk.StringVar()
        self.progress_message.set('アプリケーション実行中...')
        self.progress_label = tk.Label(self,
                                       font=('メイリオ', 12),
                                       textvariable=self.progress_message,
                                       bg='white',
                                       justify='left')
        self.progress_label.place(x=80, y=30)

        # プログレスバー
        # プログレスバーのスタイル設定
        s = ttk.Style()
        s.theme_use('clam')
        s.configure("green.Horizontal.TProgressbar",
                    foreground='limegreen',
                    background='limegreen')
        # プログレスバーの配置
        self.progress_bar = ttk.Progressbar(self,
                                            style="green.Horizontal.TProgressbar",
                                            orient='horizontal',
                                            length=400,
                                            mode='determinate')
        self.progress_bar.configure(maximum=100)
        self.progress_bar.place(x=80, y=120)


    def execute_logic(self):
        """
        メインロジックの実行（描画更新のため、別スレッドで実行）
        選択したファイルのパスと、プログレスバーのオブジェクトを渡す
        """
        th = threading.Thread(target=execute_logic,
                              args=(self.file_path,
                                    self.progress_bar,
                                    self.after_complete_process))
        th.start()

    def after_complete_process(self):
        """
        メインロジックの終了時に実行する処理
        ここでは、完了メッセージと完了（OK）ボタンの表示を行う。
        """
        # メッセージ
        self.progress_message.set('アプリケーションの実行が完了しました')
        # OKボタン
        self.ok_button = tk.Button(self,
                                   text='OK',
                                   font=('メイリオ', 10),
                                   relief='raised',
                                   cursor='hand2',
                                   fg='limegreen',
                                   bg='snow',
                                   highlightbackground='limegreen',
                                   width=10)
        self.ok_button.bind('<Enter>', self.hover_enter)
        self.ok_button.bind('<Leave>', self.hover_leave)
        self.ok_button.bind('<Button-1>', self.ok_button_clicked)
        self.ok_button.place(x=500, y=30)

    def hover_enter(self, event):
        # (注)mac, linuxではbuttonのbackgroundオプションが働かない
        self.ok_button.configure(fg='snow', bg='limegreen')

    def hover_leave(self, event):
        self.ok_button.configure(fg='limegreen', bg='snow')

    def ok_button_clicked(self, event):
        # progress windowの削除
        self.parent.destroy()

class CreateToolTip(object):

    def __init__(self, widget, text='widget info'):
        """
        引数として与えられたヴィジェットに対して、指定のテキストを
        適切な位置にツールチップとして配置する。
        また、マウスホバー時とボタン押下時のコールバック関数をバインドし、
        ツールチップとしての機能を実現する。

        parameters
        ----------
        widget :  
            ツールチップを配置するヴィジェット
        text : str
            ツールチップとして表示するテキスト
        """
        self.waittime = 500   # 単位は[ms]
        self.wraplength = 180 # pixels
        self.widget = widget
        self.text = text
        self.widget.bind('<Enter>', self.enter)
        self.widget.bind('<Leave>', self.leave)
        self.widget.bind('<ButtonPress>', self.leave)
        self.id = None
        self.tw = None

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.waittime, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self, event=None):
        x = y = 0
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        # トップレベルウィンドウの生成
        self.tw = tk.Toplevel(self.widget)
        # ラベルのみを残し、アプリケーションウィンドウを削除する
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(self.tw,
                         text=self.text,
                         justify='left',
                         background='#ffffff',
                         relief='solid',
                         borderwidth=1,
                         wraplength=self.wraplength)
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tw
        self.tw = None
        if tw:
            tw.destroy()
