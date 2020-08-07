# -*- coding: utf-8 -*-
import xlrd
from numpy import sqrt, abs, sin, pi, array, arcsin

import sys
sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Method, Fitting
from GeneralMethod.Report import Report


class Photoelectic:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = []

    PREVIEW_FILENAME = "Preview.pdf"  # 预习报告模板文件的名称
    DATA_SHEET_FILENAME = "data.xlsx"  # 数据填写表格的名称
    REPORT_TEMPLATE_FILENAME = "Photoelectric_empty.docx"  # 实验报告模板（未填数据）的名称
    REPORT_OUTPUT_FILENAME = "../../Report/Experiment2/2061Report.docx"  # 最后生成实验报告的相对路径

    def __init__(self):
        self.data = {}  # 存放实验中的各个物理量
        self.uncertainty = {}  # 存放物理量的不确定度
        self.report_data = {}  # 存放需要填入实验报告的
        print("2061 晶体的光电效应实验\n1. 实验预习\n2. 数据处理")
        while True:
            try:
                oper = input("请选择: ").strip()
            except EOFError:
                sys.exit(0)
            if oper != '1' and oper != '2':
                print("输入内容非法！请输入一个数字1或2")
            else:
                break
        if oper == '1':
            print("现在开始实验预习")
            print("正在打开预习报告......")
            Method.start_file(self.PREVIEW_FILENAME)
        elif oper == '2':
            print("现在开始数据处理")
            print("即将打开数据输入文件......")
            # 打开数据输入文件
            Method.start_file(self.DATA_SHEET_FILENAME)
            input("输入数据完成后请保存并关闭excel文件，然后按回车键继续")
            # 从excel中读取数据
            self.input_data("./"+self.DATA_SHEET_FILENAME) # './' is necessary when running this file, but should be removed if run main.py
            print("数据读入完毕，处理中......")
            # 计算物理量
            self.calc_data()
            # 计算不确定度
            self.calc_uncertainty()
            print("正在生成实验报告......")
            # 生成实验报告
            self.fill_report()
            print("实验报告生成完毕，正在打开......")
            Method.start_file(self.REPORT_OUTPUT_FILENAME)
            print("Done!")

    '''
    从excel表格中读取数据，并创建填写key的list
    @param filename: 输入excel的文件名
    @return none
    '''
    def input_data(self, filename):
        ws_plate = xlrd.open_workbook(filename).sheet_by_name('Sheet1')
        ws_electro_optic = xlrd.open_workbook(filename).sheet_by_name('Sheet2')
        # 从excel中读取数据
        list_data = []  # 实验数据：0是实验三，1是实验四
        add_item = {'theta': ws_plate.row_values(1, 1, 6), 'V1': ws_plate.row_values(2, 1, 6),
                    'V2': ws_plate.row_values(3, 1, 6), 'd': float(ws_plate.cell_value(1, 9)),
                    'l': float(ws_plate.cell_value(2, 9)), 'n0': float(ws_plate.cell_value(3, 9)),
                    'lambda': float(ws_plate.cell_value(4, 9)), 'V_pi_theo': float(ws_plate.cell_value(5, 9))}
        list_data.append(add_item)
        add_item = {'theta': ws_electro_optic.row_values(1, 1, 11), 'I': ws_electro_optic.row_values(2, 1, 11),
                    'I0': float(ws_electro_optic.cell_value(1, 14)),
                    'delta_theo': float(ws_electro_optic.cell_value(2, 14))}
        list_data.append(add_item)
        self.data['list_data'] = list_data  # 存储从表格中读入的数据

    '''
    数据处理总函数，调用三个实验的数据处理函数
    '''
    def calc_data(self):
        result_list = []
        add_result = self.calc_data1(self.data['list_data'][0])
        result_list.append(add_result)
        add_result = self.calc_data2(self.data['list_data'][1])
        result_list.append(add_result)
        self.data['result_data'] = result_list

    '''
    实验三的数据处理
    注意：若需计算的物理量较多，建议对计算过程复杂的物理量单独封装函数.
    对于实验中重要的数据，采用dict对象self.data存储，方便其他函数共用数据
    '''
    @staticmethod
    def calc_data1(data_list):
        V_pi_meas = list(array(data_list['V2']) - array(data_list['V1']))
        aver_V_pi_meas = Method.average(V_pi_meas)
        E_V_pi = abs(aver_V_pi_meas - data_list['V_pi_theo']) / data_list['V_pi_theo'] * 100
        r22 = data_list['lambda'] * data_list['d'] / \
              (2 * (data_list['n0'] ** 3) * aver_V_pi_meas * data_list['l'] * (10 ** 9))
        r22_theo = data_list['lambda'] * data_list['d'] / \
                   (2 * (data_list['n0'] ** 3) * data_list['V_pi_theo'] * data_list['l'] * (10 ** 9))
        E_r22 = abs(r22 - r22_theo) / r22_theo * 100
        result_data = {'V_pi_meas': V_pi_meas, 'aver_V_pi': aver_V_pi_meas, 'E_V_pi': E_V_pi, 'r22': r22,
                       'r22_theo': r22_theo, 'E_r22': E_r22}
        return result_data

    '''
        实验四的数据处理
        注意：若需计算的物理量较多，建议对计算过程复杂的物理量单独封装函数.
        对于实验中重要的数据，采用dict对象self.data存储，方便其他函数共用数据
        '''
    @staticmethod
    def calc_data2(data_list):
        x, y = [], []
        for i in range(len(data_list['theta'])):
            x.append(sin(data_list['theta'][i] / 360 * 4 * pi) ** 2 / 2)
            y.append(data_list['I'][i] / data_list['I0'])
        [k, b, r] = Fitting.linear(x, y, show_plot=False)
        delta = arcsin(sqrt(k)) / pi * 360
        E_delta = abs(delta - data_list['delta_theo']) / data_list['delta_theo'] * 100
        result_data = {'x': x, 'y': y, 'k': k, 'r': r, 'delta': delta, 'E_delta': E_delta}
        return result_data

    '''
    计算所有实验的不确定度
    '''
    def calc_uncertainty(self):
        uncertain_list = []
        add_uncertain = self.calc_uncertainty1(self.data['list_data'][0], self.data['result_data'][0])
        uncertain_list.append(add_uncertain)
        add_uncertain = self.calc_uncertainty2(self.data['list_data'][1], self.data['result_data'][1])
        uncertain_list.append(add_uncertain)
        self.data['uncertain_data'] = uncertain_list

    '''
    计算实验三的不确定度
    '''
    # 对于数据处理简单的实验，可以根据此格式，先计算数据再算不确定度，若数据处理复杂也可每计算一个物理量就算一次不确定度
    @staticmethod
    def calc_uncertainty1(data_list, result_list):
        ua_V_pi = Method.a_uncertainty(result_list['V_pi_meas'])
        V_pi_final = Method.final_expression(result_list['aver_V_pi'], ua_V_pi)
        u_r22_r22 = abs(ua_V_pi / result_list['aver_V_pi'])
        u_r22 = u_r22_r22 * result_list['r22']
        r22_final = Method.final_expression(result_list['r22'], u_r22)
        uncertain_data = {'ua_V_pi': ua_V_pi, 'V_pi_final': V_pi_final, 'u_r22_r22': u_r22_r22, 'u_r22': u_r22,
                          'r22_final': r22_final}
        return uncertain_data

    '''
        计算实验三的不确定度
        '''

    # 对于数据处理简单的实验，可以根据此格式，先计算数据再算不确定度，若数据处理复杂也可每计算一个物理量就算一次不确定度
    @staticmethod
    def calc_uncertainty2(data_list, result_list):
        k = result_list['k']
        u_k = k * sqrt((1 / result_list['r'] - 1) / (len(data_list['theta']) - 2))
        u_delta = abs(u_k) / sqrt((1 - k) * k) / 2 / pi * 360
        delta_final = Method.final_expression(result_list['delta'], u_delta)
        uncertain_data = {'u_k': u_k, 'u_delta': u_delta, 'delta_final': delta_final}
        return uncertain_data

    '''
    填充实验报告
    调用ReportWriter类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self):
        # 实验三
        for i, theta in enumerate(self.data['list_data'][0]['theta']):
            self.report_data['311' + str(i+1)] = "%d" % theta
        for i, V1 in enumerate(self.data['list_data'][0]['V1']):
            self.report_data['312' + str(i+1)] = "%d" % V1
        for i, V2 in enumerate(self.data['list_data'][0]['V2']):
            self.report_data['313' + str(i+1)] = "%d" % V2
        for i, V_pi_meas in enumerate(self.data['result_data'][0]['V_pi_meas']):
            self.report_data['314' + str(i+1)] = "%d" % V_pi_meas
        self.report_data['32V_pi'] = "%.3f" % self.data['result_data'][0]['aver_V_pi']
        self.report_data['32uaV_pi'] = "%.3f" % self.data['uncertain_data'][0]['ua_V_pi']
        self.report_data['32V_pi_final'] = self.data['uncertain_data'][0]['V_pi_final']
        self.report_data['32EV_pi'] = "%.2f" % self.data['result_data'][0]['E_V_pi']
        self.report_data['32d'] = "%d" % self.data['list_data'][0]['d']
        self.report_data['32l'] = "%d" % self.data['list_data'][0]['d']
        self.report_data['32n0'] = "%.3f" % self.data['list_data'][0]['d']
        self.report_data['32lambda'] = "%.1f" % self.data['list_data'][0]['lambda']
        self.report_data['32r22'] = "%.3e" % self.data['result_data'][0]['r22']
        self.report_data['32u_r22_r22'] = "%.6f" % self.data['uncertain_data'][0]['u_r22_r22']
        self.report_data['32u_r22'] = "%.3e" % self.data['uncertain_data'][0]['u_r22']
        self.report_data['32r22_final'] = self.data['uncertain_data'][0]['r22_final']
        self.report_data['32Er22'] = "%.2f" % self.data['result_data'][0]['E_r22']
        # 实验四
        for i, theta in enumerate(self.data['list_data'][1]['theta']):
            self.report_data['411' + str(i+1)] = "%d" % theta
        for i, I in enumerate(self.data['list_data'][1]['I']):
            self.report_data['412' + str(i+1)] = "%.3f" % I
        for i, x in enumerate(self.data['result_data'][1]['x']):
            self.report_data['413' + str(i+1)] = "%.5f" % x
        for i, y in enumerate(self.data['result_data'][1]['y']):
            self.report_data['414' + str(i+1)] = "%.5f" % y
        self.report_data['42k'] = "%.5f" % self.data['result_data'][1]['k']
        self.report_data['42r'] = "%.5f" % self.data['result_data'][1]['r']
        self.report_data['42delta'] = "%.2f" % self.data['result_data'][1]['delta']
        self.report_data['42ua_k'] = "%.4f" % self.data['uncertain_data'][1]['u_k']
        self.report_data['42u_delta'] = "%.5f" % self.data['uncertain_data'][1]['u_delta']
        self.report_data['42delta_final'] = self.data['uncertain_data'][1]['delta_final']
        self.report_data['42E_delta'] = "%.2f" % self.data['result_data'][1]['E_delta']
        # 调用ReportWriter类
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)


if __name__ == '__main__':
    mc = Photoelectic()
