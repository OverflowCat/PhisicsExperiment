# 主程序，只包括选择实验和运行函数，以后会加上说明一类的

# 基物实验1
import Experiment1.phy1011.Elastic_Modulus as Phy1_1011_1

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
			print("\t光的分振幅法干涉实验（1091）")
			exp = input("请输入实验序号：").strip()
			if exp == "1011":
				print("\t该模块内目前包括的子实验：")
				print("\t\t1. 拉伸法测量钢丝的弹性模量")
				eidx = input("请输入子实验序号").strip()
				if eidx == "1":
					Phy1_1011_1.Elastic_Modulus(cwd="Experiment1/phy1011/")
				else:
					print("很抱歉。暂时没有相应的数据处理程序。")
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
