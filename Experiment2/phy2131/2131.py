import xlrd
# from xlutils.copy import copy as xlscopy
import shutil
import os
import math
from numpy import sqrt, abs
import pandas
import sys
import numpy
import matplotlib
import scipy
from scipy import interpolate
# sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
sys.path.append('./.')
from GeneralMethod.PyCalcLib import Method
from GeneralMethod.PyCalcLib import Fitting
from GeneralMethod.Report import Report

class Report2131:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = []

    PREVIEW_FILENAME = "E:\基物实验程序\PhysicsExperiment\Experiment2\phy2131\Preview.pdf" 
    DATA_SHEET_FILENAME = "E:\基物实验程序\PhysicsExperiment\Experiment2\phy2131\data.xlsx"
    REPORT_TEMPLATE_FILENAME = "E:\基物实验程序\PhysicsExperiment\Experiment2\phy2131\2131_empty.docx"
    #"E:\基物实验程序\PhysicsExperiment\Experiment2\phy2131\2131_empty.docx"
    # "E:\基物实验程序\PhysicsExperiment\Experiment1\phy1022\heatConductivity_empty.docx"
    REPORT_OUTPUT_FILENAME = "E:\基物实验程序\PhysicsExperiment\Experiment2\phy2131\2131_out.docx"
    #"E:\基物实验程序\PhysicsExperiment\Experiment2\phy2131\2131_out.docx"
    #"E:\基物实验程序\PhysicsExperiment\Experiment1\phy1022\heatConductivity_out.docx"


    Method.start_file(DATA_SHEET_FILENAME)
    def __init__(self):
        self.data = {} # 存放实验中的各个物理量
        self.uncertainty = {} # 存放物理量的不确定度
        self.report_data = {} # 存放需要填入实验报告的
        print("2131 多光束干涉和法布里伯罗干涉\n1. 实验预习\n2. 数据处理")
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
            os.startfile(self.PREVIEW_FILENAME)
        elif oper == '2':
            print("现在开始数据处理")
            print("即将打开数据输入文件......")
            # 打开数据输入文件
            os.startfile(self.DATA_SHEET_FILENAME)
            input("输入数据完成后请保存并关闭excel文件，然后按回车键继续")
            # 从excel中读取数据
            self.input_data(self.DATA_SHEET_FILENAME) # './' is necessary when running this file, but should be removed if run main.py
            print("数据读入完毕，处理中......")
            # 计算物理量
            self.calc_data()
            # 计算不确定度
            self.calc_uncertainty()
            print("正在生成实验报告......")
            # 生成实验报告
            self.fill_report()
            print("实验报告生成完毕，正在打开......")
            os.startfile(self.REPORT_OUTPUT_FILENAME)
            print("Done!")


    '''
    从excel表格中读取数据
    @param filename: 输入excel的文件名
    @return none
    '''
    def input_data(self, filename):
       # print("initialyes")
        ws = xlrd.open_workbook(filename).sheet_by_name('Sheet1')
       # print("xlsx after yes")
        # 从excel中读取数据
        list_x = [1,2,3,4,5,6,7,8,9,10]
        list_i = [1,2,3,4,5,6,7,8]
        list_d = []
        list_dl = []
        list_dr = []
        list_di = []
    
        self.data['list_x'] = list_x
        self.data['list_i'] = list_i

        for row in [3]:
            for col in range(1, 11):
                list_d.append(float(ws.cell_value(row, col)))
        self.data['list_d'] = list_d

        for row in [7]:
            for col in range(1, 9):
                list_dl.append(float(ws.cell_value(row, col)))
        self.data['list_dl'] = list_dl

        for row in [8]:
            for col in range(1, 9):
                list_dr.append(float(ws.cell_value(row, col)))
        self.data['list_dr'] = list_dr

        for row in [9]:
            for col in range(1, 9):
                list_di.append(float(ws.cell_value(row, col)))
        self.data['list_di'] = list_di

        
    '''
    进行数据处理
    '''

    def calc_data(self):
        # 第一部分数据处理
        self.data['x'] = 5.5
        y = Method.average(self.data['list_d'])
        self.data['y'] = y
        self.data['x2'] = 38.5
        sum1 = 0
        for i in self.data['list_d']:
            sum1 = sum1 + i*i
        y2 = sum1 / (len(self.data['list_d']))
        self.data['y2'] = y2
        sum1 = 0
        count = 1
        for i in self.data['list_d']:
            sum1 = sum1 + i*count
            count = count + 1
        xy = sum1 / (len(self.data['list_d']))
        self.data['xy'] = xy
        self.data['x_2'] = 30.25
        self.data['y_2'] = self.data['y']**2
        self.data['lambda'] = 589.3
        (a, self.data['b'], self.data['r']) = Fitting.linear(self.data['list_x'], self.data['list_d'], show_plot=False)
        self.data['b'] = 10 ** (-5) * self.data['b']
        self.data['delta_lambda'] = 10**(-8) * self.data['lambda']**2 / (2 * self.data['b'])
        print("b = %f" % self.data['b'])
        print("r = %f" % self.data['r'])


        # 第二部分数据处理
        list_di2 = []
        i = 0
        for x in self.data['list_di']:
            list_di2.append(x**2)
        self.data['list_di2'] = list_di2
        (a, self.data['b2'], self.data['r2']) = Fitting.linear(self.data['list_i'], self.data['list_di2'], show_plot=False)
        self.data['b2'] = 10 ** (-5) * self.data['b2']
        self.data['abs_r'] = math.fabs(self.data['r2'])
        d = 4 * 632.8 * 0.15**2 * 10**(-4) / math.fabs(self.data['b2'])
        self.data['d'] = d
        print("b2 = %f" % self.data['b2'])
        print("r2 = %f" % self.data['r2'])
        print("d = %f" % self.data['d'])


    def calc_uncertainty(self):
        u_b = self.data['b'] * math.sqrt(1/8 * (1/self.data['r']**2 - 1)) * 10**(7)
        self.data['u_b'] = u_b
        u_delta = 10**(-4) * self.data['delta_lambda'] * self.data['u_b'] / self.data['b']
        self.data['u_delta'] = u_delta
        error = math.fabs(0.6 - self.data['delta_lambda'] * 10**(-1)) / 0.6
        self.data['error'] = error
        self.data['newdelta_lambda'] = self.data['delta_lambda'] * 10**(-3)
        self.data['newu_delta'] = 10**(-3) * self.data['u_delta']
        print("u_b = %f" % self.data['u_b'])
        print("delta_lambda = %f" % self.data['delta_lambda'])
        print("u_delta = %f" % self.data['u_delta'])
        print("error = %f" % self.data['error'])

    
    '''
    填充实验报告
    调用ReportWriter类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self):
        i = 0
        for i, x in enumerate(self.data['list_d']):
            self.report_data[str(i + 1)] = "%.5f" % x
        self.report_data['y'] = "%f" % self.data['y']
        self.report_data['x2'] = "%f" % self.data['x2']
        self.report_data['xy'] = "%f" % self.data['xy']
        self.report_data['y2'] = "%f" % self.data['y2']
        self.report_data['x_2'] = "%f" % self.data['x_2']
        self.report_data['y_2'] = "%f" % self.data['y_2']
        self.report_data['b'] = "%f" % self.data['b']
        self.report_data['delta_lambda'] = "%f" % self.data['delta_lambda']
        self.report_data['r'] = "%f" % self.data['r']
        self.report_data['u_b'] = "%f" % self.data['u_b']
        self.report_data['u_delta'] = "%f" % self.data['u_delta']
        self.report_data['error'] = "%f" % self.data['error']
        self.report_data['newdelta_lambda'] = "%.3f" % self.data['newdelta_lambda']
        self.report_data['newu_delta'] = "%.3f" % self.data['newu_delta']

        for i, x in enumerate(self.data['list_dl']):
            self.report_data[str(i + 101)] = "%.5f" % x
        for i, x in enumerate(self.data['list_dr']):
            self.report_data[str(i + 201)] = "%.5f" % x
        for i, x in enumerate(self.data['list_di']):
            self.report_data[str(i + 301)] = "%.5f" % x
        self.report_data['b2'] = "%f" % self.data['b2']
        self.report_data['r2'] = "%f" % self.data['r2']
        self.report_data['d'] = "%f" % self.data['d']
        self.report_data['abs_r'] = "%f" % self.data['abs_r']
        print("report")

        # 调用Report类
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

if __name__ == '__main__':
    hc = Report2131()