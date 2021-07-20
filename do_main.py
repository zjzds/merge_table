__author__ = 'zds'
# -*- coding: utf-8 -*-
# @Time    : 2021/6/7 0007 下午 2:27
# @Author  : zds
# @FileName: do_main.py
# @Software: PyCharm
from shutil import unpack_archive, copytree, rmtree
from os import mkdir, listdir, rename
from pandas import DataFrame, read_csv, concat, merge, read_excel
from bs4 import BeautifulSoup
from tkinter import messagebox
from config import *
import win32con
import win32api
from ast import literal_eval
from time import sleep, time

p_c = PATH_TEMP
p_c1_name = PATH_TEMP_NAME1
p_c2_name = PATH_TEMP_NAME2
p_c3_name = PATH_TEMP_NAME3
p_c4_name = PATH_TEMP_NAME4
path_set_j = PATH_J


def msg(text):
    m = messagebox.showinfo("Notice", text)
    return m


def read_file_to_html(path_rf, p_c_name):
    """
    复制需要合并的压缩文件并解压或文件夹中的文件至temp文件夹，并把所有文件改名为*.html,
    :param :path:
    :return:flag 1 or 0
    """
    try:
        p_path = p_c + p_c_name
        if path_rf.endswith('.zip'):  # 压缩文件
            unpack_archive(path_rf, p_path)
            re_filename(p_path)
            return p_path
        else:                       # 一般目录
            if path.exists(path_rf):
                p = copytree(path_rf, p_path)
                re_filename(p)
                return p
            else:
                msg('目录地址有误')
    except Exception as e:
        msg(e)


def re_filename(p):
    """
    把目录内的所有电子表文件改名为"*.html"文件，便于bs4解析。
    :param: p
    :return:None
    """
    if path.exists(p):
        win32api.SetFileAttributes(p, win32con.FILE_ATTRIBUTE_HIDDEN) #设置p为隐藏文件
        for f in listdir(p):
            rename(p + '\\' + f, p + '\\' + f[:-3] + 'html')
    else:
        msg('目录地址有误')


def rm_temp(path_r):
    """
    删除临时目录
    :param: path
    :return:None
    """
    try:
        rmtree(path_r)
    finally:
        pass


def parse_html2excel2(path_p, file_name, p_c_name, n):
    global p_c
    # global p_c2_name
    file_path = path_p + '\\' + file_name
    try:
        with open(file_path, 'r+', encoding='UTF-8') as f:
            str_t = f.read()
            wb = str_t.strip().replace('\ufeff', '')
        soup = BeautifulSoup(wb, 'lxml')  # 解析html
    except Exception as e:
        msg("文件格式不符，只能对金三导出的.zip或xls文件进行合并！")
        return
    table_ys = soup.findAll("table")[1].findAll("tr")  # 读取第二个表格，Excel文件转成的html。
    list1 = []
    try:
        num = 0
        if table_ys[-1].get_text()[0:3] == '合计行':
            table_ys_x = table_ys[0:-1]
        else:
            table_ys_x = table_ys
        for tr in table_ys_x:
            list_temp = []
            num += 1
            cols = tr.findAll('td')
            for td in cols:
                val = td.text
                list_temp.append(val)
            list1.append(list_temp)
        p = p_c + p_c_name
        f1 = str(n) + '.csv'
        if not path.exists(p):
            mkdir(p)
            win32api.SetFileAttributes(p, win32con.FILE_ATTRIBUTE_HIDDEN) # 设置p为隐藏文件
        save_path = path.join(p, f1)
        df = DataFrame(list1)
        df.to_csv(save_path, header=False, index=False)
        return p, num
    except Exception as e:
        msg(e)


def long_num_str(d):
    d = str(d)+'\t'
    return d


def m_table(path_m, out_pp):
    """
    合并表格
    :param: path
    :return:
    """
    try:
        dir_list = []
        data1 = DATA_DICK
        data = {}
        data.update(data1)
        if path.exists(path_set_j):
            with open(path_set_j, 'r') as fr:
                dict2 = literal_eval(fr.read())  # 读取的str转换为字典
                for k in dict2.keys():               # 把从文本中读出来的字典里的"<class 'str'>"按类型转化成str等
                    s = dict2[k][-5:-2]
                    if s == 'str':
                        dict2[k] = str
                data.update(dict2)
        for file_name in listdir(path_m):
            dir_list.append(read_csv(path_m + "\\" + file_name, converters=data))
        df = concat(dir_list)
        for s_k in data.keys():
            try:
                df[s_k] = df[s_k].map(long_num_str)
            except KeyError as e:
                pass
        df.to_csv(out_pp, index=False, encoding='utf_8_sig')
        return 1
    except Exception as e:
        msg("合并失败：{}".format(e))
        return -1


def read_table_xls(p_m_t, data):
    return read_excel(p_m_t, converters=data)


def read_table_csv(p_m_t, data):
    return read_csv(p_m_t, low_memory=False, converters=data)


def m_table2(path_m, path_1_3_1, path_1_4_2, path_g_2_3, out_pp, flag):
    """
    合并表格,如果用read_csv读，flag = 1,用read_excel读，flag = 2
    :param: path
    :return:
    """
    try:
        dir_list = []
        data1 = DATA_DICK
        data = {}
        data.update(data1)
        if path.exists(path_set_j):
            with open(path_set_j, 'r') as fr:
                dict2 = literal_eval(fr.read())  # 读取的str转换为字典
                for k in dict2.keys():               # 把从文本中读出来的字典里的"<class 'str'>"按类型转化成str等
                    s = dict2[k][-5:-2]
                    if s == 'str':
                        dict2[k] = str
                data.update(dict2)
        for file_name in listdir(path_m):
            p_m_t = path_m + "\\" + file_name
            if flag ==1:
                dir_list.append(read_table_csv(p_m_t, data))
            elif flag == 2:
                dir_list.append(read_table_xls(p_m_t, data))
            else:
                msg("表格格式不符合！")
                return -1
        # df = merge(dir_list[0],dir_list[1],how=path_g_2_3,left_on=str(path_1_3_1),right_on=str(path_1_4_2))
        path_1_3_1 = path_1_3_1.split(',')
        path_1_4_2 = path_1_4_2.split(',')
        df = merge(dir_list[0], dir_list[1], how=path_g_2_3, left_on=path_1_3_1, right_on=path_1_4_2)
        for s_k in data.keys():
            try:
                df[s_k] = df[s_k].map(long_num_str)
            except KeyError as e:
                pass
        df.to_csv(out_pp, index=False,encoding='utf_8_sig')
        return df.shape[0]
    except Exception as e:
        msg("合并失败：{}".format(e))
        return -1


def do_work(path_d, out_p, s_t):
    # global p_c2_name
    temp = p_c + p_c1_name
    if path.exists(temp):
        rm_temp(temp)
    temp2 = p_c + p_c2_name
    if path.exists(temp2):
        rm_temp(temp2)
    return_path = read_file_to_html(path_d, p_c1_name)
    i = 1
    num_t = 0
    try:
        for f_name in listdir(return_path):
            s_path, num = parse_html2excel2(return_path, f_name, p_c2_name, i)
            i += 1
            num_t += num
    except Exception as e:
        return "",False
    fl = m_table(s_path,out_p)
    if fl == -1:
        return "", False
    sleep(3)
    obj = listdir(return_path)
    l = num_t - len(obj)
    try:
        rm_temp(return_path)
        rm_temp(s_path)
    except IOError as e:
        pass
    txt = "合并完成，共合并得到数据{}条用时: {:.0f}秒".format(l, (time() - s_t))
    return txt, True


def do_work2(path_1_1, path_1_2, path_1_3, path_1_4, out_1, path_g_2, s_t1):
    if path.split(path_1_1)[0] != path.split(path_1_2)[0]:
        msg("合并的文件应放在同一文件夹下！")
        return "", False
    path_1_ml = path.split(path_1_1)[0]
    file_name1_suffix = path.split(path_1_1)[1]
    file_name2_suffix = path.split(path_1_2)[1]
    if file_name1_suffix.endswith(('.xls','.xlsx')) and file_name2_suffix.endswith(('.xls','.xlsx')):
        try:
            read_excel(path_1_1)
            read_excel(path_1_2)
            f5 = m_table2(path_1_ml, path_1_3, path_1_4, path_g_2, out_1,2)
            if f5 == -1:
                return "", False
            txt = "合并完成，共合并得到数据{}条用时: {:.0f}秒".format(f5, (time() - s_t1))
            return txt, True
        except Exception as e:
            temp3 = p_c + p_c3_name
            if path.exists(temp3):
                rm_temp(temp3)
            temp4 = p_c + p_c4_name
            if path.exists(temp4):
                rm_temp(temp4)
            return_path = read_file_to_html(path_1_ml, p_c3_name)
            i = 1
            try:
                for f_name in listdir(return_path):
                    s_path, num = parse_html2excel2(return_path, f_name, p_c4_name, i)
                    i += 1
            except Exception as e:
                return "", False
            f3 = m_table2(s_path,path_1_3,path_1_4 ,path_g_2,out_1,1)
            if f3 == -1:
                return "", False
            try:
                rm_temp(return_path)
                rm_temp(s_path)
            except IOError as e:
                pass
            txt = "合并完成，共合并得到数据{}条用时: {:.0f}秒".format(f3, (time() - s_t1))
            return txt, True
    elif file_name1_suffix.endswith('.csv') and file_name2_suffix.endswith('.csv'):
        f4 = m_table2(path_1_ml, path_1_3, path_1_4, path_g_2, out_1,1)
        if f4 == -1:
            return "", False
        txt = "合并完成，共合并得到数据{}条用时: {:.0f}秒".format(f4, (time() - s_t1))
        return txt, True
    else:
        return "", False

