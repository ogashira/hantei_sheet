import sys


class UserInterface:

    @staticmethod
    def select_sz_yt_date()-> str:
        '''
        21000101の時はテスト用として実行される
        '''
        sz_yt_date: str = ''
        while True:
            print('製造予定日を入力して下さい (例 : 20250930) / 空Returnで中止')
            print('                            *テスト時は 21000101 を入力*')
            sz_yt_date = input('製造予定日 : ')
            if not sz_yt_date:
                print('ﾌﾟﾛｸﾞﾗﾑを中止します')
                sys.exit()
                
            if (
                len(sz_yt_date) == 8 and
                2020 <= int(sz_yt_date[:4]) <= 2100 and
                1 <= int(sz_yt_date[4:6]) <= 12 and
                1 <= int(sz_yt_date[6:]) <= 31
                ):
                return f"{sz_yt_date[:4]}/{sz_yt_date[4:6]}/{sz_yt_date[6:]}"

            print('正しい年月日を入力してください')


