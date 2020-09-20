# 基物实验通用程序



## 简介

本程序是基于python的基物实验辅助工具，主要功能是提供基物实验1（基物实验2的部分也正在开发中Orz）中各实验的预习报告以及对实验数据进行处理并生成完整的实验报告。



## 安装与使用

在windows系统里解压打包文件然后直接运行main.exe即可使用（exe文件运行较慢，请耐心等候），操作流程如下：

* 选择实验所属

  * 输入1: 基物实验1

  * 输入2: 基物实验2（正在开发中）

* 输入实验代号（10XX）

* 选择程序功能

  * 输入1:实验预习

  * 输入2:数据处理

* 实验预习

  * 程序将自动打开pdf版实验预习报告

  * 或手动打开 `Experiment1\10XX\Preview.pdf`

* 数据处理

  * 程序自动打开数据输入文件  `Experiment1\10XX\data.xlsx` ，按照excel文件内提示将实验所得数据填入表中，保存并关闭文件

  * 按回车继续，并等待程序生成实验报告，完成后将自动打开word版完整实验报告`Report\Experiment1\10XXReport.docx`

  

<div  align="center">    
<img src="/pics/pic2.png" width = 800 height = 500 />
</div>

<center>程序操作流程图</center>



## 版本

当前版本号：v1.0.0

当前支持的实验有：
* 弹性模量实验（1011）
* 测量水的溶解热+焦耳热功当量（1021）
* 稳态法测不良导体热导率（1022）
* 示波器的使用（1031）
* 电阻的测量（1041）
* 双电桥法测电阻实验（1042）
* 电位差计及其应用（1051）
* 物距像距法测透镜焦距（1061）
* 平行光管法测透镜焦距（1062）
* 分光仪实验（1071）
* 光的干涉实验（1081）
* 钠光干涉实验（1082）
* 光的分振幅法干涉实验（1091）



## 开发计划

* 目前开发的第一阶段（即基物实验1）主体部分已经基本完成
* 本学期的主要任务是日常维护本程序并且不断完善基物实验1中的各个实验：
  * 根据用户使用反馈不断修复bug，改善用户体验（增加分级菜单及回退功能）
  * 补全本学期新开设的实验及假期没有完成的实验

* 预计将基物实验2的开发工作将于本学期结束后的寒假开始
* 现在急需招募愿意加入我们基物实验程序开发团队的同学！
  * 面向全体18、19还有20级同学（应该是以19级正在做基物实验的同学们为主力，毕竟需求是第一生产力）
  * 掌握python的同学优先
  * 对本项目感兴趣的同学请看下面的联系方式



## 相关项目

团队里另外一组小伙伴们做的小程序，详情请微信搜索“北航基物实验”～



## 联系方式

* 本项目的gitee地址如下，欢迎大家来star and fork：

  https://gitee.com/PhisicsExperiment/PhysicsExperiment.git

* 如果有什么问题或者改进建议，请随时联系我们！

  你可以直接在“北航基物实验”微信小程序内的“反馈”页面留下你的问题（推荐），或者添加下面的微信来进行反馈。

* 想加入本项目的同学也请直接扫下面的二维码联系我们。

  

<div  align="center">    
<img src="/pics/QRcode.jpg" width = 200 height = 200 />
</div>



## 贡献者

##### python组：

董翰元 何梓心 柯瑞奇 李茜 马文韬 朴宣夷 宋子龙 王衍 吴逸霄 张智博 褚蒙 卓乐

##### 小程序组：

曹北健 杜晨鸿 杜朋蕙 荆煦添 王雪灏 王肇凯 朱乐言

##### 感谢以上各位的付出与努力！