# -*- coding: utf-8 -*-

import sys

class   Debug():
    """
    デバッグ用のクラス
    """
    @classmethod
    def print(cls, message ):
        """
        デバッグ出力する
        :param message: 出力する文字列
        :return:
        """
        frame = sys._getframe(1)

        print( frame.f_code.co_filename + ":" + str(frame.f_lineno) + " " + message )

    @classmethod
    def error(cls, message):
        """
        クリティカルなメッセージを出力する
        :param message: 出力する文字列
        :return:
        """
        frame = sys._getframe(1)

        print( "!!!"+frame.f_code.co_filename + ":" + str(frame.f_lineno) + " " + message)


if __name__ == '__main__':
    Debug.print( "DEBUG用の出力です")
    Debug.error( "DEBUG用の出力ですってばー")
