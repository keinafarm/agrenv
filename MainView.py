# -*- coding: utf-8 -*-
from Debug import Debug
import tkinter as tk
import tkinter.font as font

class   MainView():
    ENTRY_FRAME_HEIGHT = 40
    BUTTON_FRAME_HEIGHT = 60
    ENTRY_AREA_WIDTH = 60
    BASE_MARGIN = 10

    def __init__(self):
        self.root = tk.Tk()

#        root.geometry("640x480")
        self.root.title('AgrEnv - 環境保全型農業直接支払交付金')
        self.mainFrame = tk.Frame(self.root)

        frame1 = tk.Frame(self.mainFrame, height=20)
        label1 = tk.Label(frame1,
                           text='\n環境保全型農業直接支払交付金 支援ツール\n',
                           font=("Yu Gothic UI", 18, "bold"))
        label1.pack(fill = tk.X, side="top")
        frame1.pack(side="top")

        # ファイル名入力領域
        self.impleFile = self.makeEntry('実施計画書ファイル名')
        self.planFile = self.makeEntry('事業計画書ファイル名')
        self.appliFile = self.makeEntry('交付申請書ファイル名')

        # ボタン領域
        frame5 = tk.Frame(self.mainFrame, height=MainView.BUTTON_FRAME_HEIGHT)
        btnStart = tk.Button(frame5, text='開始', width=14, command=self.process)
        btnStart.pack(fill = 'x', padx=20, side = 'left')


        btnCancel = tk.Button(frame5, text='キャンセル', width=14, command=self.exit)
        btnCancel.pack(fill = 'x', padx=20, side = 'right')
        frame5.propagate(False)
        frame5.pack(side="top", fill='x')

        # デバッグ出力領域
        frame6 = tk.Frame(self.mainFrame)
        # Text
        self.debugText = tk.Text(frame6, wrap=tk.NONE)
        self.debugText.configure()
        self.debugText.insert(1.0, "")
        self.debugText.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.S, tk.E))

        # Scrollbar
        scrollbarV = tk.Scrollbar(
            frame6,
            orient=tk.VERTICAL,
            command=self.debugText.yview)
        self.debugText['yscrollcommand'] = scrollbarV.set
        scrollbarV.grid(row=0, column=1, sticky=(tk.N, tk.S))
        frame6.pack(side="top")

        scrollbarH = tk.Scrollbar(
            frame6,
            orient=tk.HORIZONTAL,
            command=self.debugText.xview)
        self.debugText['xscrollcommand'] = scrollbarH.set
        scrollbarH.grid(row=1, column=0, sticky=(tk.E, tk.W))
        self.debugText.config(
            xscrollcommand=scrollbarH.set,
            yscrollcommand=scrollbarV.set)
        frame6.pack(side="top")

        self.mainFrame.pack(expand = 0, fill = tk.X)

    def setDefault(self, impleFile, planFile, appliFile ):
        """
        入力のデフォルト値をセット
        :param impleFile: 実施計画書ファイル名
        :param planFile: 事業計画書ファイル名
        :param appliFile: 交付申請書ファイル名
        :return:
        """
        self.impleFile.insert(tk.END,impleFile)
        self.planFile.insert(tk.END,planFile)
        self.appliFile.insert(tk.END,appliFile)


    def makeEntry(self, text ):
        """
        テキスト入力欄を作成する
        :param text: 入力初期値
        :return: 作成したEntryオブジェクト
        """
        frame = tk.Frame(self.mainFrame, height=MainView.ENTRY_FRAME_HEIGHT)
        label = tk.Label(frame,
                           text=text,
                           font=("Yu Gothic UI", 14))
        label.pack(side="left", padx=MainView.BASE_MARGIN)
        entry = tk.Entry(frame, width=MainView.ENTRY_AREA_WIDTH)
        entry.pack(side="left")
        frame.propagate(False)
        frame.pack(side="top", fill='x')
        return entry

    def start(self):
        """
        Windowsの表示を開始する
        :return:
        """
        self.root.mainloop()

    def exit(self):
        """
        中断終了の処理
        :return:
        """
        self.root.quit()
        exit()

    def process(self):
        """
        開始ボタンの処理
        :return:
        """
        if self.proc == None:
            return

        self.proc()

    def setProcess(self, proc):
        self.proc = proc

    def print(self, text):
        self.debugText.insert(tk.END,text)


###########################
#   テスト
###########################
terget = None
def testProc():
    terget.debugText.insert(tk.END, "デバッグ用出力")

if __name__ == '__main__':
    terget = MainView()
    terget.setDefault("実施計画書(元データ).xlsx", "交付申請書（添付資料）.xlsx", "事業計画書.xlsx")
    terget.setProcess(testProc)
    terget.start()


