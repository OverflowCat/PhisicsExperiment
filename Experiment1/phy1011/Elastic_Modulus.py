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
        ws = xlrd.open_workbook(filename).sheet_by_index(0)
        # 从excel中读取数据
        # 钢丝长度
        L1 = float(ws.cell_value(3, 1))
        L2 = float(ws.cell_value(3, 2))
        L_ave = float(ws.cell_value(3, 3))
        L_arr = [L1, L2]
        # 镜子和尺的距离
        H1 = float(ws.cell_value(5, 1))
        H2 = float(ws.cell_value(5, 2))
        H_ave = float(ws.cell_value(5, 3))
        H_arr = [H1, H2]
        # 光杠杆间距，重力加速度，螺旋测微计-零点误差
        b = float(ws.cell_value(3, 6))
        g = float(ws.cell_value(5, 6))
        D0 = float(ws.cell_value(6, 4))
        # 存储数据
        self.data.update({"list_L":L_arr, "num_L":L_ave, "list_H":H_arr, "num_H":H_ave})
        self.data.update({"num_b":b, "num_g":g, "num_D0":D0})
        
        # 读取D表格
        D_raw_arr = []
        D_arr = []
        for cidx in range(1, 7):
            D_raw_arr.append(float(ws.cell_value(9, cidx)))
            D_arr.append(float(ws.cell_value(10, cidx)))
        D_raw_ave = float(ws.cell_value(9, 7))
        D_ave = float(ws.cell_value(10, 7))
        # 存储
        self.data.update({"list_D_raw":D_raw_arr, "list_D": D_arr})
        self.data.update({"num_D_raw": D_raw_ave, "num_D": D_ave})
        # 读取c表格
        m_arr = []; r_p_arr = []; r_n_arr = []; r_ave_arr = []
        for cidx in range(1, 9):
            m_arr.append(float(ws.cell_value(14, cidx)))
            r_p_arr.append(float(ws.cell_value(15, cidx)))
            r_n_arr.append(float(ws.cell_value(16, cidx)))
            r_ave_arr.append(float(ws.cell_value(17, cidx)))
        # 存储
        self.data.update({"list_m":m_arr, "list_r_p":r_p_arr, "list_r_n": r_n_arr, "list_r": r_ave_arr})
        # 读取最后四个数据
        self.data['num_m'] = float(ws.cell_value(19, 1))
        self.uncertainty['D_L'] = float(ws.cell_value(19, 4))
        self.uncertainty['D_H'] = float(ws.cell_value(19, 6))
        self.uncertainty['D_b'] = float(ws.cell_value(19, 8))
        # Finish!

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