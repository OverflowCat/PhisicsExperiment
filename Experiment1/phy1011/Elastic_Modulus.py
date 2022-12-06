import xlrd
import os, sys, shutil
from numpy import sqrt, pi, abs


sys.path.append("../..")
from GeneralMethod.PyCalcLib import Fitting, Method
from GeneralMethod.Report import Report


class ElasticModulus:
    def __init__(self, cwd=""):  # 初始化实验类时给一个路径参数

        self.PREVIEW_FILENAME = cwd + "Preview.pdf"
        self.DATA_SHEET_FILENAME = cwd + "data.xlsx"
        self.REPORT_TEMPLATE_FILENAME = cwd + "Elastic_Modulus_empty.docx"  # 实验报告模板（未填数据）的名称
        self.REPORT_OUTPUT_FILENAME = cwd + "../../Report/Experiment1/1011Report.docx"  # 最后生成实验报告的相对路径

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
        D_count = 10
        # for cidx in range(1, 7):
        for cidx in range(1, D_count + 1):
            D_raw_arr.append(float(ws.cell_value(9, cidx)))
            D_arr.append(float(ws.cell_value(10, cidx)))
        D_raw_ave = sum(D_raw_arr) / len(D_raw_arr) # float(ws.cell_value(9, 7))
        D_ave = sum(D_arr) / len(D_arr) # float(ws.cell_value(10, 7))
        # 存储
        self.data.update({"list_D_raw":D_raw_arr, "list_D": D_arr})
        self.data.update({"num_D_raw": D_raw_ave, "num_D": D_ave})

        # 读取c表格
        m_arr = []; r_p_arr = []; r_n_arr = []; r_ave_arr = []
        c_count = 10
        # for cidx in range(1, 9):
        for cidx in range(1, c_count + 1):
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

        # 读取下一个表格
        ws = xlrd.open_workbook(filename).sheet_by_index(1)
        # 从excel中读取数据
        # 基础数据（数组，0：塑料圆柱，1：金属圆筒，2：圆球，3：金属杆，4：滑块）
        basic_data = []
        basic = {'m': float(ws.cell_value(3, 1)), 'd': float(ws.cell_value(4, 1))}
        basic_data.append(basic)
        basic = {'m': float(ws.cell_value(3, 2)), 'din': float(ws.cell_value(4, 3)),
                 'dout': float(ws.cell_value(6, 3))}
        basic_data.append(basic)
        basic = {'m': float(ws.cell_value(3, 4)), 'd': float(ws.cell_value(4, 4))}
        basic_data.append(basic)
        basic = {'m': float(ws.cell_value(3, 5)), 'l': float(ws.cell_value(4, 5))}
        basic_data.append(basic)
        basic = {'m1': float(ws.cell_value(3, 6)), 'm2': float(ws.cell_value(3, 7)), 'din': float(ws.cell_value(4, 7)),
                 'dout': float(ws.cell_value(5, 7)), 'h': float(ws.cell_value(6, 7))}
        basic_data.append(basic)

        # 每次计时周期数
        times = float(ws.cell_value(8, 2))

        # 摆动时间
        time_meas = []
        组数 = 5
        for row in range(5):
            time_meas.append([float(i) for i in ws.row_values(row + 12, 3, 3 + 组数)])

        # 平行轴定理
        time_meas_para = []
        for row in range(5):
            time_meas_para.append([float(i) for i in ws.row_values(row + 20, 3, 6)])

        # 存储数据
        self.data.update({"basic_data": basic_data, "times": times, "time_meas": time_meas,
                          "time_meas_para": time_meas_para})
        # Finish!

    def calc_data(self):
        list_r = self.data['list_r']
        list_dif_r, ave_c = Method.successive_diff(list_r) # ?????
        print(list_r)
        print(list_dif_r)
        print(ave_c)
        divide_by_5 = lambda x: x / 5
        list_dif_r = list(map(divide_by_5, list_dif_r))
        ave_c = divide_by_5(ave_c)

        self.data.update({"list_c": list_dif_r, "num_c": ave_c})
        # 计算弹性模量E
        m, g, L, H = self.data['num_m'], self.data['num_g'], self.data['num_L'], self.data['num_H']
        D, b, ave_c = self.data['num_D'], self.data['num_b'], self.data['num_c']
        PI = pi
        砝码重量系数 = 10
        E = (16 * 砝码重量系数 * m * g * L * H) / (PI * (D**2) * b * ave_c)
        分子 = f"16 × {砝码重量系数} × {m} × {g} × {L} × {H}"
        分母 = f"π × ({D})² × {b} × {ave_c}"
        E *= 1e6 # 单位换算
        print(f"""
     {分子}
E = {'-' * (2 + max(len(分子), len(分母)))}
     {分母}
  = {E}
        """)
        self.data['num_E'] = E

        # 计算转动惯量
        basic_data = self.data['basic_data']
        times = self.data['times']
        time_meas = self.data['time_meas']
        time_meas_para = self.data['time_meas_para']
        # 计算真实周期
        for i, row in enumerate(time_meas):
            time_meas[i].append(Method.average(row))
            time_meas[i].append(time_meas[i][-1] / times)
        for i, row in enumerate(time_meas_para):
            time_meas_para[i].append(Method.average(row))
            time_meas_para[i].append(time_meas_para[i][-1] / times)

        co = 4 * (pi ** 2)  # 提前计算一个小系数

        # 计算K
        I0 = 1 / 8 * basic_data[0]['m'] * (basic_data[0]['d'] ** 2) / 1e9
        print("4*pi**2:", co, "I_0:", I0, "T_1:", time_meas[1][-1], "T_0:", time_meas[0][-1])
        K = co * I0 / (time_meas[1][-1] ** 2 - time_meas[0][-1] ** 2)

        co = K / co

        # 金属圆筒
        I1 = co * (time_meas[2][-1] ** 2 - time_meas[0][-1] ** 2)
        I1_theo = 1 / 8 * basic_data[1]['m'] * (basic_data[1]['din'] ** 2 + basic_data[1]['dout'] ** 2) / 1e9
        e1 = abs(I1 - I1_theo) / I1_theo * 100

        # 圆球
        I2 = co * (time_meas[3][-1] ** 2)
        I2_theo = 1 / 10 * basic_data[2]['m'] * (basic_data[2]['d'] ** 2) / 1e9
        e2 = abs(I2 - I2_theo) / I2_theo * 100

        # 金属杆
        I3 = co * (time_meas[4][-1] ** 2)
        I3_theo = 1 / 12 * basic_data[3]['m'] * (basic_data[3]['l'] ** 2) / 1e9
        e3 = abs(I3 - I3_theo) / I3_theo * 100

        # 验证平行轴定理
        x, y = [], []
        for T in time_meas_para:
            y.append(T[-1] ** 2)
        for i in range(5):
            x.append((i * 5 + 5) ** 2)
        [k, b, r] = Fitting.linear(x, y, False)
        k = k * 10000
        J1 = 1 / 16 * basic_data[4]['m1'] * (basic_data[4]['din'] ** 2 + basic_data[4]['dout'] ** 2) \
             + 1 / 12 * basic_data[4]['m1'] * (basic_data[4]['h'] ** 2)
        J1 = J1 / 1e9
        J2 = 1 / 16 * basic_data[4]['m2'] * (basic_data[4]['din'] ** 2 + basic_data[4]['dout'] ** 2) \
             + 1 / 12 * basic_data[4]['m2'] * (basic_data[4]['h'] ** 2)
        J2 = J2 / 1e9
        k_theo = (basic_data[4]['m1'] + basic_data[4]['m2']) / co / 1000
        b_theo = (J1 + J2 + I3) / co

        # 更新数据
        self.data.update({'time_meas': time_meas, 'time_meas_para': time_meas_para, 'I1': I0, 'K': K,
                          'I2': I1, 'I2_theo': I1_theo, 'e2': e1,
                          'I3': I2, 'I3_theo': I2_theo, 'e3': e2,
                          'I4': I3, 'I4_theo': I3_theo, 'e4': e3,
                          '2k': k, '2k_theo': k_theo, '2b': b, '2b_theo': b_theo, '2r': r,
                          'J1': J1, 'J2': J2})
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
        u_E_E = self.uncertainty['u_E_E'] = sqrt((u_L/L)**2 + (u_H/H)**2 + 10 * (u_D/D)**2 + (u_b/b)**2 + (u_c/c)**2)
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

        # 扭转法

        # 基础数据
        for i, basic in enumerate(self.data['basic_data']):
            for key in basic.keys():
                self.report_data['s21' + key + str(i)] = "%.2f" % basic[key]
        # 每次测量周期数
        self.report_data['times'] = self.data['times']
        # 测量周期
        for i, T in enumerate(self.data['time_meas']):
            for j, detail in enumerate(T):
                self.report_data['s212' + str(i) + str(j)] = "%.2f" % detail
        # 平行轴定理
        for i, T in enumerate(self.data['time_meas_para']):
            for j, detail in enumerate(T):
                self.report_data['s213' + str(i) + str(j)] = "%.2f" % detail
        # 其它计算值
        self.report_data['I1'] = "%.4e" % self.data['I1']
        self.report_data['I2'] = "%.4e" % self.data['I2']
        self.report_data['I3'] = "%.4e" % self.data['I3']
        self.report_data['I4'] = "%.4e" % self.data['I4']
        self.report_data['I2_theo'] = "%.4e" % self.data['I2_theo']
        self.report_data['I3_theo'] = "%.4e" % self.data['I3_theo']
        self.report_data['I4_theo'] = "%.4e" % self.data['I4_theo']
        self.report_data['e2'] = "%.2f" % self.data['e2']
        self.report_data['e3'] = "%.2f" % self.data['e3']
        self.report_data['e4'] = "%.2f" % self.data['e4']
        self.report_data['K'] = "%.2e" % self.data['K']
        self.report_data['2k'] = "%d" % self.data['2k']
        self.report_data['2b'] = "%.4f" % self.data['2b']
        self.report_data['2r'] = "%.4f" % self.data['2r']
        self.report_data['2k_theo'] = "%d" % self.data['2k_theo']
        self.report_data['2b_theo'] = "%.4f" % self.data['2b_theo']
        self.report_data['J1'] = "%.4e" % self.data['J1']
        self.report_data['J2'] = "%.4e" % self.data['J2']


        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)


if __name__ == '__main__':
    EM = ElasticModulus()