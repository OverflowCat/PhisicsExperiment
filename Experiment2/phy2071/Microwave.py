# -*- coding: utf-8 -*-
import xlrd
from numpy import sqrt, abs

import sys
sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Method, Fitting
from GeneralMethod.Report import Report


class Microwave:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = []

    PREVIEW_FILENAME = "Preview.pdf"  # 预习报告模板文件的名称
    DATA_SHEET_FILENAME = "data.xlsx"  # 数据填写表格的名称
    REPORT_TEMPLATE_FILENAME = "Microwave_empty.docx"  # 实验报告模板（未填数据）的名称
    REPORT_OUTPUT_FILENAME = "../../Report/Experiment2/2071Report.docx"  # 最后生成实验报告的相对路径

    def __init__(self):
        self.data = {}  # 存放实验中的各个物理量
        self.uncertainty = {}  # 存放物理量的不确定度
        self.report_data = {}  # 存放需要填入实验报告的
        print("2071 微波实验与布拉格衍射实验\n1. 实验预习\n2. 数据处理")
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
        ws_Michelson = xlrd.open_workbook(filename).sheet_by_name('Sheet1')
        # 从excel中读取数据
        list_data = []  # 实验数据：0是实验三
        add_item = {'xn': ws_Michelson.row_values(2, 1, 5), 'xn1': ws_Michelson.row_values(4, 1, 5)}
        list_data.append(add_item)
        self.data['list_data'] = list_data  # 存储从表格中读入的数据

    '''
    数据处理总函数，调用三个实验的数据处理函数
    '''
    def calc_data(self):
        result_list = []
        add_result = self.calc_data1(self.data['list_data'][0])
        result_list.append(add_result)
        self.data['result_data'] = result_list

    '''
    实验三的数据处理
    注意：若需计算的物理量较多，建议对计算过程复杂的物理量单独封装函数.
    对于实验中重要的数据，采用dict对象self.data存储，方便其他函数共用数据
    '''
    @staticmethod
    def calc_data1(data_list):
        x = []
        for i in range(4):
            x.append((data_list['xn'][i] + data_list['xn1'][i]) / 2)
        [k, b, r] = Fitting.linear(list(range(1, 5)), x, False)
        res_lambda = k * 2
        result_data = {'x': x, 'k': k, 'r': r, 'lambda': res_lambda}
        return result_data

    '''
    计算所有实验的不确定度
    '''
    def calc_uncertainty(self):
        uncertain_list = []
        add_uncertain = self.calc_uncertainty1(self.data['list_data'][0], self.data['result_data'][0])
        uncertain_list.append(add_uncertain)
        self.data['uncertain_data'] = uncertain_list

    '''
    计算实验三的不确定度
    '''
    # 对于数据处理简单的实验，可以根据此格式，先计算数据再算不确定度，若数据处理复杂也可每计算一个物理量就算一次不确定度
    @staticmethod
    def calc_uncertainty1(data_list, result_list):
        k = result_list['k']
        u_k = k * sqrt((1 / (result_list['r'] ** 2) - 1) / (len(result_list['x']) - 2))
        u_lambda = u_k * 2
        lambda_final = Method.final_expression(result_list['lambda'], u_lambda)
        uncertain_data = {'u_k': u_k, 'u_lambda': u_lambda, 'lambda_final': lambda_final}
        return uncertain_data

    '''
    填充实验报告
    调用ReportWriter类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self):
        # 实验三
        for i, xn in enumerate(self.data['list_data'][0]['xn']):
            self.report_data['311' + str(i+1)] = "%.3f" % xn
        for i, xn1 in enumerate(self.data['list_data'][0]['xn1']):
            self.report_data['312' + str(i+1)] = "%.3f" % xn1
        for i, x in enumerate(self.data['result_data'][0]['x']):
            self.report_data['32' + str(i+1)] = "%.3f" % x
        self.report_data['32k'] = "%.4f" % self.data['result_data'][0]['k']
        self.report_data['32r'] = "%.5f" % self.data['result_data'][0]['r']
        self.report_data['32lambda'] = "%.4f" % self.data['result_data'][0]['lambda']
        self.report_data['32u_k'] = "%.4f" % self.data['uncertain_data'][0]['u_k']
        self.report_data['32u_lambda'] = "%.4f" % self.data['uncertain_data'][0]['u_lambda']
        self.report_data['32lambda_final'] = self.data['uncertain_data'][0]['lambda_final']
        # 调用ReportWriter类
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)


if __name__ == '__main__':
    mc = Microwave()
