__author__ = 'zds'
# -*- coding: utf-8 -*-
# @Time    : 2021/6/7 0007 上午 11:53
# @Author  : zds
# @FileName: config.py
# @Software: PyCharm
from os import path

PATH_TEMP = path.abspath('.') + '\\'
PATH_TEMP_NAME1 = "temp_jssq"
PATH_TEMP_NAME2 = "temp2_jssq"
PATH_TEMP_NAME3 = "temp3_jssq"
PATH_TEMP_NAME4 = "temp4_jssq"
PATH_J = PATH_TEMP + "dick_j.txt"
DATA_DICK = {
                "登记表单展示": str,
                "电子税票号码": str,
                "登记序号": str,
                "社会信用代码（纳税人识别号）": str,
                "证照编号": str,
                "法定代表人（负责人、业主）身份证件号码": str,
                "财务负责人身份证件号码": str,
                "办税人身份证件号码": str,
                "税收档案编号": str,
                "社会信用代码": str,
                "原纳税人识别号": str
            }
