# 主程序，只包括选择实验和运行函数，以后会加上说明一类的

# 基物实验1
import Experiment1.phy1011.Elastic_Modulus as Phy1_1011
import Experiment1.phy1021.thermology as Phy1_1021
import Experiment1.phy1022.heatConductivity as Phy1_1022
import Experiment1.phy1031.Oscillograph as Phy1_1031
import Experiment1.phy1041.RMeasure as Phy1_1041
import Experiment1.phy1042.Kelvinbridge as Phy1_1042
import Experiment1.phy1051.Potentiometer as Phy1_1051
import Experiment1.phy1061.f_measure as Phy1_1061
import Experiment1.phy1062.Parallel as Phy1_1062
import Experiment1.phy1071.Spectrometer as Phy1_1071
import Experiment1.phy1081.Laser as Phy1_1081
import Experiment1.phy1082.Sodium as Phy1_1082
import Experiment1.phy1091.Michelson as Phy1_1091_1
import Experiment1.phy1091.NewtonRing as Phy1_1091_2
# 基物实验2
import Experiment2.MillikanOilDrop.MillikanOilDrop as Millikan
import Experiment2.SolarBattery.SolarBattery as sb
import Experiment2.Frank_Hertz.frank_hertz as fh

# TODO: 在main里调用子程序可能有路径问题
# 解决方法：参考我写的1011，在初始化类的构造函数__init__里加一个路径参数，默认是空(由本文件调用)，如果在main里调用就加上这个实验相对main的路径

if __name__ == '__main__':
	# 主程序，只引用模块
	try:
		term = input("请选择实验所属：\n1：基物实验1\n2：基物实验2\n（请输入序号）：")
		if term == "1":
			print("目前基物实验1中可计算的实验：")
			print("\t弹性模量实验（1011）")
			print("\t测量水的溶解热+焦耳热功当量（1021）")
			print("\t稳态法测不良导体热导率（1022）")
			print("\t示波器的使用（1031）")
			print("\t电阻的测量（1041）")
			print("\t双电桥法测电阻实验（1042）")
			print("\t电位差计及其应用（1051）")
			print("\t物距像距法测透镜焦距（1061）")
			print("\t平行光管法测透镜焦距（1062）")
			print("\t分光仪实验（1071）")
			print("\t光的干涉实验（1081）")
			print("\t钠光干涉实验（1082）")
			print("\t光的分振幅法干涉实验（1091）")
			exp = input("请输入实验序号：").strip()
			if exp == "1011":
				Phy1_1011.ElasticModulus(cwd="Experiment1/phy1011/")
			elif exp == "1021":
				Phy1_1021.thermology(cwd="Experiment1/phy1021/")
			elif exp == "1022":
				Phy1_1022.heatConductivity(cwd="Experiment1/phy1022/")
			elif exp == "1031":
				Phy1_1031.Oscillograph(cwd="Experiment1/phy1031/")
			elif exp == "1041":
				Phy1_1041.RMeasure(cwd="Experiment1/phy1041/")
			elif exp == "1042":
				Phy1_1042.Kelvinbridge(cwd="Experiment1/phy1042/")
			elif exp == "1051":
				Phy1_1051.Potentiometer(cwd="Experiment1/phy1051/")
			elif exp == "1061":
				Phy1_1061.f_measure(cwd="Experiment1/phy1061/")
			elif exp == "1062":
				Phy1_1062.Parallel(cwd="Experiment1/phy1062/")
			elif exp == "1071":
				Phy1_1071.Spectrometer(cwd="Experiment1/phy1071/")
			elif exp == "1081":
				Phy1_1081.Laser(cwd="Experiment1/phy1081/")
			elif exp == "1082":
				Phy1_1082.Sodium(cwd="Experiment1/phy1082/")
			elif exp == "1091":
				print("\t该模块内目前包括的子实验：")
				print("\t\t1. 迈克尔逊干涉")
				print("\t\t2. 牛顿环干涉")
				eidx = input("请输入子实验序号").strip()
				if eidx == "1":
					Phy1_1091_1.Michelson()
				elif eidx == "2":
					Phy1_1091_2.NewtonRing()
			else:
				print("很抱歉。暂时没有相应实验的数据处理程序。")
		elif term == "2":
			print("目前基物实验2中可计算的实验：")
			print("\t1、密里根油滴实验")
			print("\t2、太阳能电池特性实验")
			print("\t3、弗兰克赫兹实验")
			exp = input("请输入实验序号：")
			if exp == "1":
				Millikan.Millikan()
			elif exp == "2":
				sb.SolarBattery()
			elif exp == "3":
				fh.FrankHertz()
			else:
				print("很抱歉。暂时没有相应的数据处理程序。")
	except ValueError:
		print("请输入一个数字")
	input("点击任意键退出")
