from re import I
from typing import Dict
class KensaHinban:

    def __init__(self, index_no: int, yotei_info: Dict[str, object], 
                                    spec_info: Dict[str, str])-> None:

        self.__index_no = index_no
        self.__date = yotei_info['date']
        self.__hinban = yotei_info['hinban']
        self.__lot = yotei_info['lot']
        self.__net = yotei_info['net']
        self.__cans = yotei_info['cans']

        # 桁を揃える 数値に変換出来ない場合は''
        try:
            actual_vis_min = f"{float(spec_info['actual_vis_min']):.1f}"
        except ValueError:
            actual_vis_min = ' - '
        try:
            actual_vis_max = f"{float(spec_info['actual_vis_max']):.1f}"
        except ValueError:
            actual_vis_max = ' - '
        try:
            actual_sg_min = f"{float(spec_info['actual_sg_min']):.3f}"
        except ValueError:
            actual_sg_min = ' - '
        try:
            actual_sg_max = f"{float(spec_info['actual_sg_max']):.3f}"
        except ValueError:
            actual_sg_max = ' - '
        try:
            actual_nv_min = f"{float(spec_info['actual_nv_min']):.1f}"
        except ValueError:
            actual_nv_min = ' - '
        try:
            actual_nv_max = f"{float(spec_info['actual_nv_max']):.1f}"
        except ValueError:
            actual_nv_max = ' - '
        try:
            spec_vis_min = f"{float(spec_info['spec_vis_min']):.1f}"
        except ValueError:
            spec_vis_min = ' - '
        try:
            spec_vis_max = f"{float(spec_info['spec_vis_max']):.1f}"
        except ValueError:
            spec_vis_max = ' - '
        try:
            spec_sg_min = f"{float(spec_info['spec_sg_min']):.3f}"
        except ValueError:
            spec_sg_min = ' - '
        try:
            spec_sg_max = f"{float(spec_info['spec_sg_max']):.3f}"
        except ValueError:
            spec_sg_max = ' - '
        try:
            spec_nv_min = f"{float(spec_info['spec_nv_min']):.1f}"
        except ValueError:
            spec_nv_min = ' - '
        try:
            spec_nv_max = f"{float(spec_info['spec_nv_max']):.1f}"
        except ValueError:
            spec_nv_max = ' - '

        self.__actual_vis = f'{actual_vis_min} / {actual_vis_max}'
        self.__actual_sg = f'{actual_sg_min} / {actual_sg_max}'
        self.__actual_nv = f'{actual_nv_min} / {actual_nv_max}'
        self.__spec_vis = f'{spec_vis_min} / {spec_vis_max}'
        self.__spec_sg = f'{spec_sg_min} / {spec_sg_max}'
        self.__spec_nv = f'{spec_nv_min} / {spec_nv_max}'

        self.__nendo_cup = spec_info['nendo_cup']
        if spec_info['nendo_cup'] == '-':
            self.__nendo_cup = ''
        self.__bikou_hs = spec_info['bikou_hs']
        if spec_info['bikou_hs'] == '-':
            self.__bikou_hs = ''


    def filling_data(self, ws, standerd_row)-> None:

        ws.cell(row=standerd_row,     column=2).value = self.__hinban
        ws.cell(row=standerd_row + 1, column=2).value = self.__lot
        net_cans:str = f'{self.__net} X {self.__cans}'
        ws.cell(row=standerd_row + 1, column=4).value = net_cans

        ws.cell(row=standerd_row + 8, column=6).value = self.__actual_nv
        ws.cell(row=standerd_row + 8, column=7).value = self.__actual_vis
        ws.cell(row=standerd_row + 8, column=8).value = self.__actual_sg
        ws.cell(row=standerd_row + 9, column=6).value = self.__spec_nv
        ws.cell(row=standerd_row + 9, column=7).value = self.__spec_vis
        ws.cell(row=standerd_row + 9, column=8).value = self.__spec_sg

        ws.cell(row=standerd_row + 7, column=2).value = self.__nendo_cup
        ws.cell(row=standerd_row + 8, column=2).value = self.__bikou_hs

        # 加熱残分が無い場合
        if self.__spec_nv == ' -  /  - ':
            ws.cell(row=standerd_row + 4, column=2).value = '---'
            ws.cell(row=standerd_row + 5, column=2).value = '---'
            ws.cell(row=standerd_row + 6, column=2).value = '---'
            ws.cell(row=standerd_row + 4, column=3).value = '---'
            ws.cell(row=standerd_row + 5, column=3).value = '---'
            ws.cell(row=standerd_row + 6, column=3).value = '---'
            ws.cell(row=standerd_row + 4, column=4).value = '---'
            ws.cell(row=standerd_row + 5, column=4).value = '---'
            ws.cell(row=standerd_row + 6, column=4).value = '---'
            ws.cell(row=standerd_row + 4, column=5).value = '---'
            ws.cell(row=standerd_row + 5, column=5).value = '---'
            ws.cell(row=standerd_row + 6, column=5).value = '---'
            ws.cell(row=standerd_row + 4, column=6).value = '---'
            ws.cell(row=standerd_row + 5, column=6).value = '---'
            ws.cell(row=standerd_row + 6, column=6).value = '---'
            ws.cell(row=standerd_row + 7, column=6).value = '---'

        # 粘度が無い場合
        if self.__spec_vis == ' -  /  - ':
            ws.cell(row=standerd_row + 4, column=7).value = '---'
            ws.cell(row=standerd_row + 5, column=7).value = '---'
            ws.cell(row=standerd_row + 6, column=7).value = '---'
            ws.cell(row=standerd_row + 7, column=7).value = '---'

        # 比重の表示 通常は2/3は '---' N3の時は ''
        ws.cell(row=standerd_row + 5, column=8).value = '---'
        ws.cell(row=standerd_row + 6, column=8).value = '---'
        if self.__spec_sg == ' -  /  - ':
            ws.cell(row=standerd_row + 4, column=8).value = '---'
            ws.cell(row=standerd_row + 7, column=8).value = '---'
        if self.__bikou_hs == 'N3' or self.__bikou_hs == 'n3':
            ws.cell(row=standerd_row + 4, column=8).value = ''
            ws.cell(row=standerd_row + 5, column=8).value = ''
            ws.cell(row=standerd_row + 6, column=8).value = ''




