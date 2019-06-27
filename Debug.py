# -*- coding: utf-8 -*-

import sys

class   Debug():
    """
    デバッグ用のクラス
    """
    view = None
    @classmethod
    def print(cls, message ):
        """
        デバッグ出力する
        :param message: 出力する文字列
        :return:
        """
        frame = sys._getframe(1)

        text = frame.f_code.co_filename + ":" + str(frame.f_lineno) + " " + str(message) + '\n'
        if Debug.view :
            Debug.view.print(text)
        else:
            print( text)

    @classmethod
    def error(cls, message):
        """
        クリティカルなメッセージを出力する
        :param message: 出力する文字列
        :return:
        """
        frame = sys._getframe(1)
        text = "!!!"+frame.f_code.co_filename + ":" + str(frame.f_lineno) + " " + str(message) + '\n'
        if Debug.view :
            Debug.view.print(text)
        else:
            print( text)

    @classmethod
    def setView(cls, view):
        Debug.view = view

if __name__ == '__main__':
    Debug.print("DEBUG用の出力です")
    Debug.error("DEBUG用の出力ですってばー")
