from pandas import read_excel

import os, sys, shutil
sys.path.append("../..")
from GeneralMethod.PyCalcLib import Method
from GeneralMethod.Report import Report

class SolarBattery:
    def __init__(self, cwd=""):
        self.in_path = cwd + "in.xlsx"
        self.out_path = cwd + "out.xlsx"
        self.template_path = cwd + "SolarBattery_empty.docx"
        self.outdocx_path = cwd + "../../Report/Experiment2/2042Report.docx"
        print("Data Process Begin!")
        self.show()
        print("Fill Report Begin!")
        self.write2word()
        Method.start_file(self.outdocx_path)
        print("All Done!")

    def show(self):
        try:
            print("reading data from {}".format(self.in_path))
            df = read_excel(self.in_path)

            print("pasring result")
            for row in [7, 15, 23, 31]:
                for col in range(3, 23):
                    # R
                    df.iloc[row, col] = round(1000 * df.iloc[row-1, col] / df.iloc[row-2, col], 1)
                    # P
                    df.iloc[row+1, col] = round(df.iloc[row-1, col] * df.iloc[row-2, col], 2)
            for idx in range(4):
                p_row = list(df.iloc[8*idx+8, 3:23])
                # Pmax
                df.iloc[40, 3+idx] = max(p_row)
                # Rmax
                df.iloc[35, 3+idx] = df.iloc[8*idx+7, p_row.index(df.iloc[40, 3+idx])+3]
                # Ri
                df.iloc[36, 3+idx] = 1000 * df.iloc[8*idx+3, 3] / int(df.iloc[8*idx+2, 3][:-2])
                # Rmax/Ri
                df.iloc[37, 3+idx] = round(df.iloc[35, 3+idx] / df.iloc[36, 3+idx], 3)
                # U0*IS
                df.iloc[41, 3+idx] = round(df.iloc[8*idx+3, 3] * int(df.iloc[8*idx+2, 3][:-2]), 2)
                # F=Pmax/(U0*IS)
                df.iloc[42, 3+idx] = round(df.iloc[40, 3+idx] / df.iloc[41, 3+idx], 3)
                # Fmean
                df.iloc[44, 3+idx] = round(sum(list(df.iloc[42, 3:7])) / 4, 3)

                df.iloc[35, 3+idx] = round(df.iloc[35, 3+idx], 1)
                df.iloc[36, 3+idx] = round(df.iloc[36, 3+idx], 1)

            # print("saving to {}".format(self.out_path))
            self.df = df
            # df.to_excel(self.out_path, index=False, header=False)
            print("done")
        except Exception as ex:
            print("exception encountered:\n{}".format(ex))
            input()
    def write2word(self):
        self.report_data = {}
        # 分别填入四个表格
        # self.df就是表格，从第0行0列开始, 用self.df.iloc[i, j]访问第i行第j列
        for igroup in range(1, 5): # 第1组到第4组
            # I, U, R, P四个变量
            # 每组的I对应的行号分别是5, 13, 21, 29, 8 * igroup - 3
            r_I = 8 * igroup - 3
            self.report_data["%d-I0" % igroup] = self.df.iloc[r_I - 3, 3]
            self.report_data["%d-U0" % igroup] = self.df.iloc[r_I - 2, 3]
            for idx in range(1, 21): # 每个物理量的序号是1到20
                kw = "%d-I%d" % (igroup, idx)
                self.report_data[kw] = "%g" % self.df.iloc[r_I, idx + 2]
                kw = "%d-U%d" % (igroup, idx)
                self.report_data[kw] = "%g" % self.df.iloc[r_I + 1, idx + 2]
                kw = "%d-R%d" % (igroup, idx)
                self.report_data[kw] = "%g" % self.df.iloc[r_I + 2, idx + 2]
                kw = "%d-P%d" % (igroup, idx)
                self.report_data[kw] = "%g" % self.df.iloc[r_I + 3, idx + 2]
            # 填充最后的小表格里的数据
            self.report_data['%d-Rm' % igroup] = "%g" % self.df.iloc[35, igroup + 2]
            self.report_data['%d-Ri' % igroup] = "%g" % self.df.iloc[36, igroup + 2]
            self.report_data['%d-RmRi' % igroup] = "%g" % self.df.iloc[37, igroup + 2]

            self.report_data['%d-Pm' % igroup] = "%g" % self.df.iloc[40, igroup + 2]
            self.report_data['%d-U0Is' % igroup] = "%g" % self.df.iloc[41, igroup + 2]
            self.report_data['%d-F' % igroup] = "%g" % self.df.iloc[42, igroup + 2]
        # 最后一个物理量
        self.report_data['F_mean'] = "%g" % self.df.iloc[44, 6]
        # 调用Report
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.template_path, self.outdocx_path)
if __name__ == "__main__":
    SolarBattery()
