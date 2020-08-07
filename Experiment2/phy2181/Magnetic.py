import xlrd
import os, shutil, sys
from numpy import sqrt, asarray, linspace, arange, cos, pi, polyfit
from matplotlib import pyplot as plt
import matplotlib
matplotlib.use("Agg")
from scipy import interpolate
from scipy.optimize import curve_fit

sys.path.append("../..")
from GeneralMethod.PyCalcLib import Method, Fitting
from GeneralMethod.Report import Report

class Magnetic:
    def __init__(self, cwd=""): # 初始化实验类时给一个路径参数
        
        self.cwd = cwd
        self.PREVIEW_FILENAME = cwd + "Preview.pdf"
        self.DATA_SHEET_FILENAME = cwd + "data.xlsx"
        self.REPORT_TEMPLATE_FILENAME = cwd + "Magnetic_empty.docx"  # 实验报告模板（未填数据）的名称
        self.REPORT_OUTPUT_FILENAME = cwd + "../../Report/Experiment1/2181Report.docx"  # 最后生成实验报告的相对路径

        self.data = {}
        self.report_data = {}
        self.report_pics = {} # 本实验报告里有图片

        print("2181 磁阻传感器和磁场测量\n1. 实验预习\n2. 数据处理")
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
            # Method.start_file(self.DATA_SHEET_FILENAME)
            input("输入数据完成后请保存并关闭excel文件，然后按回车键继续")
            # 从excel中读取数据
            self.input_data("./"+self.DATA_SHEET_FILENAME) # './' is necessary when running this file, but should be removed if run main.py
            print("数据读入完毕，处理中......")
            # 计算物理量
            self.calc_data()
            # 画图
            self.draw_graph()
            print("正在生成实验报告......")
            # 生成实验报告
            self.fill_report()
            print("实验报告生成完毕，正在打开......")
            Method.start_file(self.REPORT_OUTPUT_FILENAME)
            print("Done!")
        # Finished!
    def input_data(self, filename):
        ws = xlrd.open_workbook(self.DATA_SHEET_FILENAME).sheet_by_index(0)
        # 第一个表格
        list_1_1_I = []
        list_1_1_B = []
        list_1_1_V = []
        for i in range(1, 14):
            list_1_1_I.append(float(ws.cell_value(3, i)))
            list_1_1_B.append(float(ws.cell_value(4, i)))
            list_1_1_V.append(float(ws.cell_value(5, i)))
        self.data.update({"list_1_1_I":list_1_1_I, "list_1_1_B": list_1_1_B, "list_1_1_V": list_1_1_V})
        # 第二个表格
        list_1_2_angle = []
        list_1_2_V = []
        for i in range(1, 11):
            list_1_2_angle.append(float(ws.cell_value(7, i)))
            list_1_2_V.append(float(ws.cell_value(8, i)))
        self.data.update({"list_1_2_angle": list_1_2_angle, "list_1_2_V": list_1_2_V})
        # 第三个表格
        list_2_1_BxB0 = []
        list_2_1_BxV = []
        list_2_1_BxGauss = []
        for i in range(1, 12):
            list_2_1_BxB0.append(float(ws.cell_value(12, i)))
            list_2_1_BxV.append(float(ws.cell_value(13, i)))
            list_2_1_BxGauss.append(float(ws.cell_value(14, i)))
        self.data.update({"list_2_1_BxB0":list_2_1_BxB0, "list_2_1_BxV":list_2_1_BxV, "list_2_1_BxGauss":list_2_1_BxGauss})
        # 第四个表格(这是个二维数组)
        list_2_2 = []
        for r in range(17, 24):
            aline = []
            for c in range(1, 8):
                aline.append(float(ws.cell_value(r, c)))
            list_2_2.append(aline)
        self.data['list_2_2'] = asarray(list_2_2)
        # 第五个表格
        list_3 = []
        for i in range(1, 5):
            list_3.append(float(ws.cell_value(27, i)))
        self.data['list_3'] = list_3
        # Finished!
    def calc_data(self):
        # 这应该就算一个之前的斜率
        list_1_1_B = self.data['list_1_1_B']
        list_1_1_V = self.data['list_1_1_V']
        a, b, r = Fitting.linear(list_1_1_B, list_1_1_V)
        self.data['sen'] = a / 5 * 50
    def draw_graph(self):
        # 第一个表格
        list_1_1_B = self.data['list_1_1_B']
        list_1_1_V = self.data['list_1_1_V']
        plt.clf()
        plt.scatter(list_1_1_B, list_1_1_V, marker='+',s=3)
        a, b, r = Fitting.linear(list_1_1_B, list_1_1_V)
        _x = linspace(min(list_1_1_B), max(list_1_1_B))
        _y = a * _x + b
        plt.plot(_x, _y)
        plt.savefig("1_1.jpg")
        # 第二个表格
        x = self.data['list_1_2_angle']
        y = self.data['list_1_2_V']
        # fun = interpolate.interp1d(x, y, kind='quadratic')
        fun = lambda x, a: a * cos(pi * x / 180)
        popt, pcov = curve_fit(fun, x, y)
        _x = linspace(min(x), max(x))
        _y = fun(_x, popt[0])
        plt.clf()
        plt.plot(_x, _y)
        plt.savefig("1_2.jpg")
        # 第三个表格
        x = arange(-5, 6) / 10
        y = self.data['list_2_1_BxGauss']
        # fun = interpolate.interp1d(x, y, kind='quadratic')
        fargs = polyfit(x, y, 2)
        fun = lambda x : fargs[2] + fargs[1] * x + fargs[0] * (x**2)
        _x = linspace(min(x),max(x))
        _y = fun(_x)
        plt.clf()
        plt.scatter(x, y)
        plt.plot(_x,_y)
        plt.savefig("2_1.jpg")
        # Finished

    def fill_report(self):
        # 填充一行表格, 参数: word中键的格式(如: "a_%d", 包含一个%d位置); 数据(数组)名称
        def fill_a_line(keyfmt:str, dataname:str, outfmt="%g"):
            list_data = self.data[dataname]
            for i, di in enumerate(list_data,start=1):
                k = keyfmt % i
                self.report_data[k] = outfmt % di
        # 第一个表
        fill_a_line("1_1_I%d", "list_1_1_I")
        fill_a_line("1_1_B%d", "list_1_1_B")
        fill_a_line("1_1_V%d", "list_1_1_V", "%.3f")
        # 第二个表
        fill_a_line("1_2_a%d", "list_1_2_angle")
        fill_a_line("1_2_V%d", "list_1_2_V", "%.3f")
        # 第三个表
        fill_a_line("2_1_x%d", "list_2_1_BxB0", "%.3f")
        fill_a_line("2_1_v%d", "list_2_1_BxV", "%.3f")
        fill_a_line("2_1_g%d", "list_2_1_BxGauss", "%.3f")
        # 第四个表
        list_2_2 = self.data['list_2_2']
        for il, aline in enumerate(list_2_2, start=1):
            for i, di in enumerate(aline,start=1):
                k = "2_2_%d%d" % (il, i)
                self.report_data[k] = "%.3f" % di
        # 最后一个地磁场表格
        list_3 = self.data['list_3']
        self.report_data['3_u1'], self.report_data['3_u2'], self.report_data['3_u'], self.report_data['3_B'] = ["%.3f" % _i for _i in list_3]
        
        # 其他的单个数据
        self.report_data['1_1_sen'] = "%e" % self.data['sen']
        self.report_data['1_2_cosA'] = "%.3f" % self.data['list_1_2_V'][0]

        # 插入图片
        self.report_pics['Graph-1-1'] = self.cwd + "1_1.jpg"
        self.report_pics['Graph-1-2'] = self.cwd + "1_2.jpg"
        self.report_pics['Graph-2-1'] = self.cwd + "2_1.jpg"

        # 输出Word
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.load_insert_pic(self.report_pics)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

    

if __name__ == '__main__':
    Mg = Magnetic()