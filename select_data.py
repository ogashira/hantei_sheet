import warnings
import sys
import pandas as pd
import platform
from typing import List
from abc import ABC, abstractmethod

warnings.filterwarnings('ignore', category=UserWarning)

class ISelectData(ABC):

    def __init__(self)->None:
        '''
        サーバーにあるsql_server.pyをモジュールとして使う
        importするためにsys.path.appendでpathを認識させて
        importと生成を行う
        '''
        shared_folder_path:str = r'./'
        if platform.system() == 'Linux':
            shared_folder_path = r'/mnt/public/技術課ﾌｫﾙﾀﾞ/200. effit_data/ﾏｽﾀ/sql_python_module'
        elif platform.system() == 'Windows':
            shared_folder_path = r'//192.168.1.247/共有/技術課ﾌｫﾙﾀﾞ/200. effit_data/ﾏｽﾀ/sql_python_module'
        else:
            pass

        sys.path.append(shared_folder_path)
        from sql_server_tss import SqlServer
        self.sql_server:SqlServer = SqlServer()


    @abstractmethod
    def fetch_data(self)-> pd.DataFrame:
        pass



class SelectV3(ISelectData):

    def __init__(self, sz_yt_date: str) -> None:
        super().__init__()
        self.__sz_yt_date:str = sz_yt_date


    def fetch_data(self)-> pd.DataFrame:
        cnxn = self.sql_server.get_cnxn()

        # cursor = cnxn.cursor()

        sqlQuery = ("SELECT SZ_YT_DT AS 'Date', SZ_KBN AS 'isTumekae'," 
                    " KT_KBN AS 'SorR', ITEM_ID AS 'hinban', LOT AS 'Lot'," 
                    " IREME_QTY AS 'ireme', SZ_QTY AS 'cans'"  
                    " From dbo.TF_SZ_YT"
                    " WHERE SZ_YT_DT = ? AND DEL_FLG <> ?" 
                    " ORDER BY ITEM_ID")
        df: pd.DataFrame = pd.read_sql(sqlQuery, cnxn, 
                                       params=[self.__sz_yt_date, '1'])
        self.sql_server.close()

        return df

    def fetch_data_test(self)-> pd.DataFrame:
        # 入力値が21000101の時呼ばれる
        d = self.__sz_yt_date
        df = pd.DataFrame({'Date':[d, d, d, d, d],
                           'isTumekae':['01', '01', '01', '01', '01'],
                           'SorR':['03', '03', '03', '03', '03'],
                           'hinban':['S6-UV361-U', 
                                     'S6-SV3800L-U', 
                                     'S9-U330-TH', 
                                     'S7-A-M', 
                                     'S4-BS421B-U'],
                           'Lot':['20250101T', 
                                  '20250102T', 
                                  '20250103T', 
                                  '20250104T', 
                                  '20250105T'],
                           'ireme':[15, 15 , 14 , 15 , 15 ],
                           'cans':[25, 220 , 28 , 20 , 100 ]
                           })

        return df


class DLHanteiSheet(ISelectData):

    def __init__(self) -> None:
        pass


    def fetch_data(self)-> pd.DataFrame:
        path = './'
        if platform.system() == 'Windows':
            path = r'//192.168.1.247/共有/技術課ﾌｫﾙﾀﾞ/200. effit_data/ﾏｽﾀ/hantei_sheet.csv'
        if platform.system() == 'Darwin':
            path = r'/Volumes/共有/技術課ﾌｫﾙﾀﾞ/200. effit_data/ﾏｽﾀ/hantei_sheet.csv'
        if platform.system() == 'Linux':
            path = r'/mnt/public/技術課ﾌｫﾙﾀﾞ/200. effit_data/ﾏｽﾀ/hantei_sheet.csv'

        df = pd.read_csv(path, encoding='cp932')

        return df


