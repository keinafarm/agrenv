# -*- coding: utf-8 -*-

############################################################
#
#   元号ベースの日付を扱う処理
#
#   30より大きい年は平成と判断する、それ以外は令和
#   内部表現では令和は平成の30年分に+令和を加えて管理する
#
#
############################################################

class   EraProc():
    def __init__(self, year, month, day ):
        self._month = month
        self._day = day
        if year >= 30:
            self._year = year
        else:
            self._year = year+30

    def year(self):
        if self._year > 31:
            return self._year-30            # 31年より大きい時は令和２年以降
        elif self._year == 30:
            return 30                       # 30年は平成30年
        elif self._month >= 5:
            return 1                        # 5/1以降は令和元年
        else:
            return 31                       # それ以前は平成31年

    def isGreater(self, anotherDay ):
        if ( self._year > anotherDay._year): # こっちの年が大きければ文句なく大きい
            return True
        elif self._year < anotherDay._year:  # こっちの年が小さければ文句なく小さい
            return False
        elif self._month > anotherDay._month:    # こっちの月が大きければ文句なく大きい
            return True
        elif self._month < anotherDay._month:  # こっちの月が小さければ文句なく小さい
            return False
        elif self._day > anotherDay._day:    # こっちの日が大きければ文句なく大きい
            return True
        else:
            return False                    # 日付が小さいか、同じの時は大きくない

    def print(self):
        return   str(self.year()) + "/" + str(self._month) + "/" + str(self._day)

    def month(self):
        return  str(self._month)

if __name__ == '__main__':
    date301201 = EraProc(30,12,1)
    date310430 = EraProc(31,4,30)
    date310501 = EraProc(31,5,1)
    date320101 = EraProc(32,1,1)
    date020201 = EraProc(2,2,1)

    print( date301201.print() )
    print(date310430.print())
    print(date310501.print())
    print(date320101.print())
    print(date301201.print())
    print(date020201.print())

    print( date301201.isGreater(date310430) )
    print( date310430.isGreater(date310501) )
    print( date310501.isGreater(date310430) )
    print( date301201.isGreater(date310430) )
