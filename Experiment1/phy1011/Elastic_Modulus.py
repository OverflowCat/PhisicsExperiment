import xlrd
import os, sys, shutil
from numpy import sqrt


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
        # Finished!
    
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
        list_r = self.data['list_r']
        list_dif_r, ave_c = Method.successive_diff(list_r)
        self.data.update({"list_c": list_dif_r, "num_c": ave_c})
        # 计算弹性模量E
        m, g, L, H = self.data['num_m'], self.data['num_g'], self.data['num_L'], self.data['num_H']
        D, b, ave_c = self.data['num_D'], self.data['num_b'], self.data['num_c']
        PI = 3.1416
        E = (16 * 4 * m * g * L * H) / (PI * (D**2) * b * ave_c)
        E *= 1e6 # 单位换算
        self.data['num_E'] = E
        # Finish!
    
    def calc_uncertainty(self):
        # 几个只有b类的不确定度
        u_L = self.uncertainty['u_L'] = self.uncertainty['D_L'] / sqrt(3)
        u_H = self.uncertainty['u_H'] = self.uncertainty['D_H'] / sqrt(3)
        u_b = self.uncertainty['u_b'] = self.uncertainty['D_b'] / sqrt(3)
        L, H, D, b, c = self.data['num_L'], self.data['num_H'], self.data['num_D'], self.data['num_b'], self.data['num_c']

        # D的不确定度
        list_D = self.data['list_D']
        ua_D = self.uncertainty['ua_D'] = Method.a_uncertainty(list_D)
        ub_D = self.uncertainty['ub_D'] = 0.005 / sqrt(3)
        u_D = self.uncertainty['u_D'] = sqrt(ua_D ** 2 + ub_D ** 2)
        # c的不确定度
        list_c = self.data['list_c']
        ua_c = self.uncertainty['ua_c'] = Method.a_uncertainty(list_c)
        ub_c = self.uncertainty['ub_c'] = 0.05 / sqrt(3)
        u_c = self.uncertainty['u_c'] = sqrt(ua_c ** 2 + ub_c ** 2)
        # E的不确定度合成
        E = self.data['num_E']
        u_E_E = self.uncertainty['u_E_E'] = sqrt((u_L/L)**2 + (u_H/H)**2 + 4 * (u_D/D)**2 + (u_b/b)**2 + (u_c/c)**2)
        u_E = self.uncertainty['u_E'] = E * u_E_E
        # 最终结果表述
        self.data['final'] = Method.final_expression(E, u_E)
        

    def fill_report(self):
        # 钢丝长度
        self.report_data['L1'], self.report_data['L2'] = ["%.2f" % _i for _i in self.data['list_L']]
        self.report_data['L_ave'] = "%.2f" % self.data['num_L']
        # 平面镜距标尺的距离
        self.report_data['H1'], self.report_data['H2'] = ["%.2f" % _i for _i in self.data['list_H']]
        self.report_data['H_ave'] = "%.2f" % self.data['num_H']
        # 光杠杆前后足间距
        self.report_data['base_b'] = "%.2f" % self.data['num_b']
        self.report_data['unc0_b'] = "%.2f" % self.uncertainty['D_b']
        # 螺旋测微计的零点误差
        self.report_data['D0'] = "%.3f" % self.data['num_D0']
        # 钢丝直径 - 表格
        for i, Draw_i in enumerate(self.data['list_D_raw'], start=1):
            key = "Dc_%d" % i
            self.report_data[key] = "%.3f" % Draw_i
        self.report_data['Dc_ave'] = "%.3f" % self.data['num_D_raw']
        for i, D_i in enumerate(self.data['list_D'], start=1):
            key = "D_%d" % i
            self.report_data[key] = "%.3f" % D_i
        self.report_data['D_ave'] = "%.3f" % self.data['num_D']
        # 加力 - 标尺读数
        for i, m_i in enumerate(self.data['list_m'], start=1):
            key = "m_%d" % i
            self.report_data[key] = "%.3f" % m_i
        for i, r_p_i in enumerate(self.data['list_r_p'], start=1):
            key = "r_p%d" % i
            self.report_data[key] = "%.2f" % r_p_i
        for i, r_n_i in enumerate(self.data['list_r_n'], start=1):
            key = "r_n%d" % i
            self.report_data[key] = "%.2f" % r_n_i
        for i, r_i in enumerate(self.data['list_r'], start=1):
            key = "r_%d" % i
            self.report_data[key] = "%.3f" % r_i
        # 逐差法弹性模量
        for i, c_i in enumerate(self.data['list_c'], start=1):
            key = "c_%d" % i
            self.report_data[key] = "%.3f" % c_i
        self.report_data['c_ave'] = "%.3f" % self.data['num_c']
        # 弹性模量计算结果
        E_bse, E_pwr = Method.scientific_notation(self.data['num_E'])
        self.report_data['E'] = "%.3fe%d" % (E_bse, E_pwr)
        # 不确定度
        self.report_data['u_L'] = "%.3f" % self.uncertainty['u_L']
        self.report_data['u_H'] = "%.3f" % self.uncertainty['u_H']
        self.report_data['u_b'] = "%.4f" % self.uncertainty['u_b']
        # D的不确定度
        self.report_data['ua_D'] = "%e" % self.uncertainty['ua_D']
        self.report_data['ub_D'] = "%e" % self.uncertainty['ub_D']
        self.report_data['u_D'] = "%e" % self.uncertainty['u_D']
        # c的不确定度
        self.report_data['ua_c'] = "%e" % self.uncertainty['ua_c']
        self.report_data['ub_c'] = "%e" % self.uncertainty['ub_c']
        self.report_data['u_c'] = "%e" % self.uncertainty['u_c']
        # E的不确定度合成
        self.report_data['u_E_E'], self.report_data['u_E'] = "%e" % self.uncertainty['u_E_E'], "%e" % self.uncertainty['u_E']
        # 输出最终结果
        self.report_data['final'] = self.data['final']

        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

if __name__ == '__main__':
    EM = Elastic_Modulus()
    pass