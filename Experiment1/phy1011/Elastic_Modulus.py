import xlrd
import os, sys, shutil
from numpy import sqrt, sys

sys.path.append("../..")
from GeneralMethod.PyCalcLib import Fitting, Method
from GeneralMethod.Report import Report


class Elastic_Modulus:
    PREVIEW_FILENAME = "Preview.pdf"
    DATA_SHEET_FILENAME = "data.xlsx"
    REPORT_TEMPLATE_FILENAME = "Elastic_Modulus_empty.docx"  # 实验报告模板（未填数据）的名称
    REPORT_OUTPUT_FILENAME = "../../Report/Experiment1/1011Report.docx"  # 最后生成实验报告的相对路径

    def __init__(self):
        self.data = {}
        self.uncertainty = {}
        self.report_data = {}

        print("1011 弹性模量实验\n1. 实验预习\n2. 数据处理")
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
        pass
    
    def input_data(self, filename):
        pass

    def calc_data(self):

        pass
    
    def calc_uncertainty(self):
        pass

    def fill_report(self):
        pass

    pass

if __name__ == '__main__':
    EM = Elastic_Modulus()
    pass