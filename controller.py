import sys
import yaml
from typing import Dict
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

        # hantei_sheet.yamlを得る。メタルのクリヤに対するmixの対応表
        # 対応表に載ってるクリヤがあったら下表にmixを追加する
        yaml_path:str = r'//192.168.1.247/共有/技術課ﾌｫﾙﾀﾞ/' \
                        r'200. effit_data/ﾏｽﾀ/hantei_sheet.yaml'
        try:
            with open(yaml_path, 'r', encoding='utf-8') as file:
                hantei_yaml = yaml.safe_load(file)
        except FileNotFoundError:
            print(f'エラー："{yaml_path}"が見つかりません')
            sys.exit()
        except yaml.YAMLError as exc:
            print(f'yamlファイルのパース中にエラー出ました:{exc}')
            sys.exit()
        except Exception as e:
            print(f'yamlファイルのパース中にエラー出ました:{e}')
            sys.exit()

        add_mix_metal:Dict = hantei_yaml['add_mix_metal']

        kensaHinbans: List[KensaHinban] = []
        kensaHinbans = CreateKensaHinbanInstance.create(df_yotei, 
                                               df_hantei, add_mix_metal)


        excel:Excel = Excel(kensaHinbans, sz_yt_date)


