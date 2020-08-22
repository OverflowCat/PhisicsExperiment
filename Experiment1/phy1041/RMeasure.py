import xlrd
import shutil
import os
from numpy import sqrt, abs

import sys
#sys.path.append('./.')  # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.Report import Report
from GeneralMethod.PyCalcLib import Fitting, InstrumentError, Method

# 求电阻箱每一级的示数
def calc_r(num_R):
    list_r = []
    num_R = int(num_R * 10)
    cnt = 0
    count = 0
    while count < 6:
        list_r.append(num_R % 10 * (10 ** (cnt - 1)))
        num_R = num_R // 10
        cnt = cnt + 1
        count = count + 1
    return list_r

class RMeasure:
    report_data_keys = []

    def __init__(self, cwd=""): # 初始化实验类时给一个路径参数
        self.PREVIEW_FILENAME = cwd + "Preview.pdf"
        self.DATA_SHEET_FILENAME = cwd + "data.xlsx"
        self.REPORT_TEMPLATE_FILENAME = cwd + "RMeasure_empty.docx"
        self.REPORT_OUTPUT_FILENAME = cwd + "../../Report/Experiment1/RMeasure_out.docx"

        self.data = {}  # 存放实验中的各个物理量
        self.uncertainty = {}  # 存放物理量的不确定度
        self.report_data = {}  # 存放需要填入实验报告的
        print("1041 电阻的测量\n1. 实验预习\n2. 数据处理")
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
            # './' is necessary when running this file, but should be removed if run main.py
            self.input_data("./"+self.DATA_SHEET_FILENAME)
            print("数据读入完毕，处理中......")
            # 计算物理量
            self.calc_data()
            # 计算不确定度
            # self.calc_uncertainty()
            print("正在生成实验报告......")
            # 生成实验报告
            self.fill_report()
            print("实验报告生成完毕，正在打开......")
            Method.start_file(self.REPORT_OUTPUT_FILENAME)
            print("Done!")

    def input_data(self, filename):
        ws = xlrd.open_workbook(filename).sheet_by_name('RMeasure')
        # 从excel中读取数据
        list_x = []
        list_y = []
        list_2 = []
        list_3 = []


        for col in range(1, 9):
            list_y.append(float(ws.cell_value(3, col)))
            list_x.append(float(ws.cell_value(4, col)))
        for col in range(1, 6):
            list_2.append(float(ws.cell_value(8, col)))
            list_3.append(float(ws.cell_value(12, col)))

        self.data['list_x'] = list_x
        self.data['list_y'] = list_y
        self.data['list_2'] = list_2
        self.data['list_3'] = list_3
        # print(list_x)

    def calc_data(self):
        # 实验一
        # TODO ua_rx和ub_rx的计算方式有点问题 （不知道那个k是个什么物理量啊）
        num_b, num_a, num_r = Fitting.linear(
            self.data['list_x'], self.data['list_y'], False)
        self.data['num_a'] = num_a
        self.data['num_b'] = num_b
        self.data['num_r'] = num_r
        num_rx = num_b
        list_k = []
        for i in range(0, 8):
            list_k.append(self.data['list_x'][i]/self.data['list_y'][i])
        # num_ua_b = Method.a_uncertainty(list_k)
        num_k = 8
        num_ave_x = Method.average(self.data['list_x'])
        num_ave_y = Method.average(self.data['list_y'])
        num_ua_b = num_b * (1 / (num_k-2) * ((1/num_r) ** 2 - 1)) ** (1/2)
        num_u_u = 0.00433
        num_u_i = 0.0000433
        num_ub_rx_rx = sqrt((num_u_u/num_ave_y) ** 2 +
                            (num_u_i/num_ave_x) ** 2)
        num_ub_rx = num_ub_rx_rx * num_rx
        num_u_rx = (num_ub_rx ** 2 + num_ua_b ** 2) ** (1/2)
        num_u_rx_1bit, pwr = Method.scientific_notation(num_u_rx)
        num_fin_rx = int(num_rx * (10 ** pwr)) / (10 ** pwr)
        self.data['num_ave_y'] = num_ave_y
        self.data['num_ave_x'] = num_ave_x
        self.data['num_ub_rx'] = num_ub_rx
        self.data['num_ub_rx_rx'] = num_ub_rx_rx
        self.data['num_u_u'] = num_u_u
        self.data['num_u_i'] = num_u_i
        self.data['num_rx'] = num_rx
        self.data['num_ua_b'] = num_ua_b
        self.data['num_u_rx'] = num_u_rx
        self.data['str_fin_rx'] = "%.0f±%.0f" % (num_fin_rx, num_u_rx_1bit)

        # 实验二
        num_2r1 = self.data['list_2'][0]
        num_2v = self.data['list_2'][1]
        num_2r0 = self.data['list_2'][2]
        num_2d = self.data['list_2'][3]
        num_2r2 = self.data['list_2'][4]
        num_2rg = num_2r2

        self.data['num_2r0'] = num_2r0
        self.data['num_2r1'] = num_2r1
        self.data['num_2r2'] = num_2r2
        self.data['num_2v'] = num_2v
        self.data['num_2d'] = num_2d

        num_2ki = num_2r1*num_2v / ((num_2r0 + num_2r1) * num_2rg * num_2d)
        list_a = [5, 0.5, 0.2, 0.1, 0.1, 0.1]
        list_r0 = calc_r(self.data['num_2r0'])
        list_r1 = calc_r(self.data['num_2r1'])
        list_r2 = calc_r(self.data['num_2r2'])

        num_dt_r0 = InstrumentError.resistance_box(list_a, list_r0, 0.02)
        num_dt_r1 = InstrumentError.resistance_box(list_a, list_r1, 0.02)
        num_dt_r2 = InstrumentError.resistance_box(list_a, list_r2, 0.02)
        
        u_r0 = num_dt_r0 / sqrt(3)
        u_r1 = num_dt_r1 / sqrt(3)
        u_r2 = num_dt_r2 / sqrt(3)
        u_v = 0.00866
        u_d = 1 / sqrt(3)
        u_mix = sqrt(u_r0**2 + u_r1**2)
        u_ki = num_2ki * (sqrt(
            (u_r1/num_2r1)**2 + (u_mix/(num_2r0+num_2r1))**2 + (u_v/num_2v)**2 + (u_r2/num_2rg)**2 + (u_d/num_2d)**2
        ))

        self.data['u_r0'] = u_r0
        self.data['u_r1'] = u_r1
        self.data['u_r2'] = u_r2
        self.data['u_v'] = u_v
        self.data['u_d'] = u_d
        self.data['u_mix'] = u_mix
        self.data['u_ki'] = u_ki
        
        self.data['num_dt_2r0'] = num_dt_r0
        self.data['num_dt_2r1'] = num_dt_r1
        self.data['num_dt_2r2'] = num_dt_r2
        self.data['num_2r1'] = num_2r1
        self.data['num_2r2'] = num_2r2
        self.data['num_2r0'] = num_2r0
        self.data['num_2d'] = num_2d
        self.data['num_2v'] = num_2v
        self.data['num_2ki'] = num_2ki

        # 实验三
        num_3rs = self.data['list_3'][0]
        num_3v = self.data['list_3'][1]
        num_3d = self.data['list_3'][2]
        num_3rg = self.data['list_3'][3]
        num_3ki = self.data['list_3'][4]
        num_u3_ki = 2.678*10**(-11)
        num_3rxh = num_3rs/((num_3rg+num_3rs)*num_3ki)*num_3v/num_3d
        
        self.data['num_3d'] = num_3d
        self.data['num_3rs'] = num_3rs
        self.data['num_3v'] = num_3v
        self.data['num_3rg'] = num_3rg

        num_u3_d = 1 / sqrt(3)
        list_rs = calc_r(self.data['num_3rs'])
        num_dt_rs = InstrumentError.resistance_box(list_a, list_rs, 0.02)
        num_u3_rs = num_dt_rs / sqrt(3)
        num_u3_v = 0.00866
        list_rg2 = calc_r(self.data['num_3rg'])
        num_dt_rg2 = InstrumentError.resistance_box(list_a, list_rg2, 0.02)
        num_u3_rg = num_dt_rg2 / sqrt(3)

        num_u3_r = (num_u3_rg**2+num_u3_rs**2) ** (1/2)
        num_u3_rxh_rxh = ((num_u3_rs/num_3rs) ** 2+(num_u3_r/num_3rs) ** 2 +
                          (num_u3_v/num_3v)**2+(num_u3_ki/num_3ki)**2+(num_u3_d/num_3d)**2) ** (1/2)
        num_u3_rxh = num_u3_rxh_rxh * num_3rxh

        self.data['num_3rxh'] = num_3rxh
        self.data['num_u3_rs'] = num_u3_rs
        self.data['num_u3_v'] = num_u3_v
        self.data['num_u3_r'] = num_u3_r
        self.data['num_u3_rxh_rxh'] = num_u3_rxh_rxh
        self.data['num_u3_rxh'] = num_u3_rxh
        self.data['num_3ki'] = num_3ki
        self.data['num_u3_d'] = num_u3_d
        self.data['num_u3_rg'] = num_u3_rg
        self.data['num_u3_ki'] = num_u3_ki

    def fill_report(self):
        for i, x_i in enumerate(self.data['list_x']):
            self.report_data["x"+str(i + 1)] = "%.5f" % (x_i)
            self.report_data["x2"+str(i + 1)] = "%.5f" % (x_i ** 2)
        for i, y_i in enumerate(self.data['list_y']):
            self.report_data["y"+str(i + 1)] = "%.5f" % (y_i)
            self.report_data["y2"+str(i + 1)] = "%.5f" % (y_i ** 2)
        for i in range(0, 8):
            self.report_data["xy"+str(i+1)] = (self.data['list_x']
                                               [i])*(self.data['list_y'][i])
        self.report_data['ave_y'] = "%.4f" % self.data['num_ave_y']
        self.report_data['ave_x'] = "%.4f" % self.data['num_ave_x']
        self.report_data['b'] = "%.4f" % self.data['num_b']
        self.report_data['a'] = "%.4f" % self.data['num_a']
        self.report_data['r'] = "%.4f" % self.data['num_r']
        self.report_data['Rx'] = "%.4f" % self.data['num_rx']
        self.report_data['ua_b'] = "%.4f" % self.data['num_ua_b']
        self.report_data['ub_rx_rx'] = "%.4f" % self.data['num_ub_rx_rx']
        self.report_data['ub_rx'] = "%.4f" % self.data['num_ub_rx']
        self.report_data['u_rx'] = "%.4f" % self.data['num_u_rx']
        self.report_data['fin_Rx'] = self.data['str_fin_rx']
        self.report_data['dt_2r0'] = self.data['num_dt_2r0']
        self.report_data['dt_2r1'] = self.data['num_dt_2r1']
        self.report_data['dt_2r2'] = self.data['num_dt_2r2']
        self.report_data['2v'] = self.data['num_2v']
        self.report_data['2d'] = self.data['num_2d']
        self.report_data['2r2'] = self.data['num_2r2']
        self.report_data['2r1'] = self.data['num_2r1']
        self.report_data['2r0'] = self.data['num_2r0']
        self.report_data['2ki'] = self.data['num_2ki']
        self.report_data['u_r0'] = self.data['u_r0']
        self.report_data['u_r1'] = self.data['u_r1']
        self.report_data['u_r2'] = self.data['u_r2']
        self.report_data['u_v'] = self.data['u_v']
        self.report_data['u_d'] = self.data['u_d']
        self.report_data['u_mix'] = self.data['u_mix']
        self.report_data['u_ki'] = self.data['u_ki']

        self.report_data['3rs'] = self.data['num_3rs']
        self.report_data['3v'] = self.data['num_3v']
        self.report_data['3d'] = self.data['num_3d']
        self.report_data['3rg'] = self.data['num_3rg']
        self.report_data['3ki'] = self.data['num_3ki']
        self.report_data['u3_rg'] = self.data['num_u3_rg']
        self.report_data['u3_ki'] = self.data['num_u3_ki']
        self.report_data['3rxh'] = self.data['num_3rxh']
        self.report_data['u3_rs'] = self.data['num_u3_rs']
        self.report_data['u3_v'] = self.data['num_u3_v']
        self.report_data['u3_r'] = self.data['num_u3_r']
        self.report_data['u3_d'] = self.data['num_u3_d']
        self.report_data['fin_u_3rxh_3rxh'] = self.data['num_u3_rxh_rxh']
        self.report_data['fin_u_3rxh'] = self.data['num_u3_rxh']
        

        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME,self.REPORT_OUTPUT_FILENAME)


if __name__ == '__main__':
    rms = RMeasure()
    pass
