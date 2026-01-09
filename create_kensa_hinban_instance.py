from typing import List, Dict 
from kensa_hinban import KensaHinban

class CreateKensaHinbanInstance:

    @staticmethod
    def create(df_yotei, df_hantei) -> List[KensaHinban]:

        kensa_hinbans: List[KensaHinban] = []
        for i in range(len(df_yotei)):
            yotei_info: Dict[str, object] = {}
            hinban: str = df_yotei.iloc[i,:]['hinban']
            yotei_info['date'] = df_yotei.iloc[i,:]['Date']
            yotei_info['sz_kbn'] = df_yotei.iloc[i,:]['isTumekae']
            yotei_info['kt_kbn'] = df_yotei.iloc[i,:]['SorR']
            yotei_info['hinban'] = hinban
            yotei_info['lot'] =  df_yotei.iloc[i,:]['Lot']
            yotei_info['net'] = df_yotei.iloc[i,:]['ireme']
            yotei_info['cans'] = df_yotei.iloc[i,:]['cans']

            # 大事
            if len(df_hantei[df_hantei['入力名'] == hinban])==0:
                continue

            spec_info: Dict[str, object] = {} 
            spec_info['actual_vis_min'] = \
                  df_hantei[df_hantei['入力名'] == hinban].iloc[0,:]['粘度最小']
            spec_info['actual_vis_max'] = \
                  df_hantei[df_hantei['入力名'] == hinban].iloc[0,:]['粘度最大']
            spec_info['actual_sg_min'] = \
                 df_hantei[df_hantei['入力名'] == hinban].iloc[0,:]['比重最小']
            spec_info['actual_sg_max'] = \
                 df_hantei[df_hantei['入力名'] == hinban].iloc[0,:]['比重最大']
            spec_info['actual_nv_min'] = \
                 df_hantei[df_hantei['入力名'] == hinban].iloc[0,:]['加残最小']
            spec_info['actual_nv_max'] = \
                 df_hantei[df_hantei['入力名'] == hinban].iloc[0,:]['加残最大']

            spec_info['spec_vis_min'] = \
                 df_hantei[df_hantei['入力名'] == hinban].iloc[0,:]['vis_min']
            spec_info['spec_vis_max'] = \
                 df_hantei[df_hantei['入力名'] == hinban].iloc[0,:]['vis_max']
            spec_info['spec_sg_min'] = \
                 df_hantei[df_hantei['入力名'] == hinban].iloc[0,:]['sg_min']
            spec_info['spec_sg_max'] = \
                 df_hantei[df_hantei['入力名'] == hinban].iloc[0,:]['sg_max']
            spec_info['spec_nv_min'] = \
                 df_hantei[df_hantei['入力名'] == hinban].iloc[0,:]['nv_min']
            spec_info['spec_nv_max'] = \
                 df_hantei[df_hantei['入力名'] == hinban].iloc[0,:]['nv_max']

            spec_info['nendo_cup'] = \
                 df_hantei[df_hantei['入力名'] == hinban].iloc[0,:]['粘度カップ']
            spec_info['bikou_hs'] = \
                 df_hantei[df_hantei['入力名'] == hinban].iloc[0,:]['備考_hs']

            # インスタンス生成
            kensa_hinbans.append(KensaHinban(i, yotei_info, spec_info))

        return kensa_hinbans
