# -*- coding: utf-8 -*-
from Debug import Debug
import tkinter as tk
import tkinter.font as font

class   MainView():
    def __init__(self):
        root = tk.Tk()

#        root.geometry("640x480")
        root.title('AgrEnv - 環境保全型農業直接支払交付金')
        mainFrame = tk.Frame(root,
                             bg="green")

        frame1 = tk.Frame(mainFrame)
        label1 = tk.Label(frame1,
                           text='\n環境保全型農業直接支払交付金 支援ツール\n',
                           bg="lime green", fg="blue",
                           font=("@System", 18, "bold"))
        label1.pack(fill = tk.X, side="top")
        frame1.pack(side="top")
        frame2 = tk.Frame(mainFrame)
        label2 = tk.Label(frame2,
                           text='実施計画書ファイル名',
                           bg="lime green", fg="blue",
                           font=("@System", 14))
        label2.pack(side="left")
        impleFile = tk.Entry(frame2, width=70)
        impleFile.pack(side="left")
        frame2.pack(side="top")

        frame3 = tk.Frame(mainFrame)
        label3 = tk.Label(frame3,
                           text='事業計画書ファイル名',
                           bg="lime green", fg="blue",
                           font=("@System", 14))
        label3.pack(side="left")
        planFile = tk.Entry(frame3, width=70)
        planFile.pack(side="left")
        frame3.pack(side="top")

        frame4 = tk.Frame(mainFrame)
        label4 = tk.Label(frame4,
                           text='交付申請書ファイル名',
                           bg="lime green", fg="blue",
                           font=("@System", 14))
        label4.pack(side="left")
        appliFile = tk.Entry(frame4, width=70)
        appliFile.pack(side="left")
        frame4.pack(side="top")

        frame5 = tk.Frame(mainFrame)
        btnStart = tk.Button(frame5, text='開始', width=14)
        btnStart.pack(side="left")
        btnCancel = tk.Button(frame5, text='終了', width=14)
        btnCancel.pack(side="left")
        frame5.pack(side="top")

        frame6 = tk.Frame(mainFrame)
        # Text
        txt = tk.Text(frame6, wrap=tk.NONE)
        txt.configure()
        txt.insert(1.0, "Hello!")
        txt.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.S, tk.E))

        # Scrollbar
        scrollbarV = tk.Scrollbar(
            frame6,
            orient=tk.VERTICAL,
            command=txt.yview)
        txt['yscrollcommand'] = scrollbarV.set
        scrollbarV.grid(row=0, column=1, sticky=(tk.N, tk.S))
        frame6.pack(side="top")

        scrollbarH = tk.Scrollbar(
            frame6,
            orient=tk.HORIZONTAL,
            command=txt.xview)
        txt['xscrollcommand'] = scrollbarH.set
        scrollbarH.grid(row=1, column=0, sticky=(tk.E, tk.W))
        txt.config(
            xscrollcommand=scrollbarH.set,
            yscrollcommand=scrollbarV.set)
        frame6.pack(side="top")

        mainFrame.pack(expand = 0, fill = tk.X)


        root.mainloop()

###########################
#   テスト
###########################
if __name__ == '__main__':
    terget = MainView()
