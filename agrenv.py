# -*- coding: utf-8 -*-
from ArgMan import ArgMan
from AgrenvModel import AgrenvModel
from MainView import MainView
from Debug import Debug


#################################################################
#
#   環境保全型農業直接支払い交付金の為のデータ作成
#
#################################################################
#
class agrenv:

    outFile = "out.xlsx"
    dir = "."
    file = '実施計画書(元データ).xlsx'

    def __init__(self):

        self.__argMan = ArgMan("環境保全型農業直接支払い交付金の実施計画書からデータを抜き出す",
                               [
                                   ["--o", None, self.outFile, None, '出力ファイル名', '出力ファイル指定(指定しない場合はout.xlsx)'],
                                   ["--d", None, self.dir, None, 'ディレクトリ名', 'ディレクトリ指定'],
                                   ["--f", None, self.file, None, 'ファイル名', 'ファイル指定'],
                               ])

        self.view = MainView()
        Debug.setView(self.view)
        self.view.setDefault("実施計画書(元データ).xlsx", "交付申請書（添付資料）.xlsx", "事業計画書.xlsx")
        self.view.setProcess(self.processStart)

        self.view.start()

    def processStart(self):

        self.model = AgrenvModel(self.view.impleFile.get())
        self.model.output(self.view.appliFile.get(), self.view.planFile.get())

if __name__ == '__main__':
    obj = agrenv()
