# -*- coding: utf-8 -*-
import argparse                 # コマンドパーサー

class ArgMan:
    """
    コマンドパラメータを取得するクラス

    parser.add_argumentへのパラメータをListに列挙する事によって、パラメータ形式を指定する

    add_argumentの使い方は、
    https://docs.python.org/ja/3/library/argparse.html
    https://qiita.com/kzkadc/items/e4fc7bc9c003de1eb6d0


    コマンド実行時、取得したパラメータがListに格納される

    """
    __parser = None
    __isFirst = True
    args = None

    def __init__(self, usage, paramList):
        """
        コンストラクタ

        :param usage: 使用法を記載した文字列
        :param paramList: パラメータ形式を指定するList
        [
            [ name, type, default, nargs,  metavar, help ]
            [ name, type, default, nargs,  metavar, help ]
                    :
        ]
        の形式で指定する

        """
        self. __parser = argparse.ArgumentParser(usage=usage)       # パーサーを作成

        for parameter in paramList:                                # ListからパラメータListを取り出す
            name, type, default, nargs, metavar, help = parameter   # 取り出したパラメータList内のアイテムを取り出す
            self.__parser.add_argument(name, type=type, nargs=nargs, default=default, metavar=metavar,help=help)
                                                                    # add_argumentに、パラメータの書式を渡す

    def getArgs(self):
        """
        パラメータを解析して返す
        :return: 解析したパラメータ
        """
        if(self.__isFirst):                         # 初回かどうかをチェック
            self.__isFirst = False                 # 初回フラグをオフ
            self.args = self.__parser.parse_args()  # パラメータを解析
        return self.args

###########################
#
#   テスト用
#
#       > python ArgMan.py
#           dir=.
#           file=
#           *.*
#       > python ArgMan.py -h
#           usage: Usage:Test用
#
#           optional arguments:
#             -h, --help            show this help message and exit
#             --dir ディレクトリ名         ディレクトリ指定
#            --file [ファイル名 [ファイル名 ...]]
#                                  ファイル指定
#       > python ArgMan.py --dir=testDir
#            dir=testDir
#            file=
#            *.*
#
#       > python ArgMan.py --dir=testDir --file FileA FileB File.C A,FileD
#            dir=testDir
#            file=
#            ['FileA', 'FileB', 'File.C', 'A,FileD']
#
###########################

if __name__ == '__main__':

    argMan = ArgMan( "Usage:Test用",
        [
            [ "--dir", None, '.', None, 'ディレクトリ名','ディレクトリ指定'  ],
            [ "--file", None, '*.*', '*', 'ファイル名','ファイル指定'  ],
        ])
    args = argMan.getArgs()
    print("dir=" +args.dir )
    print("file=")
    print(args.file )
