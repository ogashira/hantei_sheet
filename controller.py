import sys
from user_interface import UserInterface
from create_kensa_hinban_instance import CreateKensaHinbanInstance
from select_data import *
from excel import Excel

class Controller:

    def start(self)-> None:
        sz_yt_date: str = UserInterface.select_sz_yt_date()
        select_data: ISelectData = SelectV3(sz_yt_date)
        # 入力値が21000101の時はテスト用 
        if sz_yt_date == '2100/01/01': # テスト用モック
            df_yotei: pd.DataFrame = select_data.fetch_data_test()
        else:
            df_yotei: pd.DataFrame = select_data.fetch_data()



        # 詰め替えではなく製造、SではなくR or SRに絞る
        df_yotei = df_yotei.loc[(df_yotei['isTumekae'] == '01') & \
                                                   (df_yotei['SorR'] != '01'),:]

        if df_yotei.empty:
            print("この製造予定日での検査製品はありません")
            sys.exit()

        DL_data: ISelectData = DLHanteiSheet()
        df_hantei: pd.DataFrame = DL_data.fetch_data()
        df_hantei = df_hantei.fillna('-')

        kensaHinbans: List[KensaHinban] = []
        kensaHinbans = CreateKensaHinbanInstance.create(df_yotei, df_hantei)



        excel:Excel = Excel(kensaHinbans, sz_yt_date)


