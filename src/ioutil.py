# -*- coding: utf-8 -*-
import os


def listdir(path, list_name, return_type):  # 传入存储的list
    '''
    save all file path in the dir @path in the list @list_name
    :param path: dir
    :param list_name: return list with all path
    :param return_type: "name": return the filename, "path" return the file path
    :return:
    '''
    for file in os.listdir(path):
        file_path = os.path.join(path, file)
        if os.path.isdir(file_path):
            listdir(file_path, list_name, return_type)
        else:
            if return_type == "path":
                list_name.append(file_path)
            else:
                list_name.append(file_path)


def read_config_work1():
    '''
    read the config file of work1
    :return:
    '''
    read_labels = []
    result_file_name = "汇总.xls"
    sheet_name = "人行填写"
    statistic_labels = "汇总"
    company_name_list = []
    with open("../config/work1.conf", "r", encoding="utf-8") as f:
        for line in f.readlines():
            line_list = line.replace("\n", "").split(":")
            if line_list[0] == "Company_name":
                company_name_list.append(line_list[1])
            elif line_list[0] == "Sheet_Name":
                sheet_name = line_list[1]
            elif line_list[0] == "Read_Label":
                read_labels.append(line_list[1])
            elif line_list[0] == "Result_File_Name":
                result_file_name = line_list[1] + ".xls"
            elif line_list[0] == "Statistic_Label":
                statistic_labels = line_list[1]
            else:
                pass
    return read_labels,result_file_name,sheet_name,statistic_labels,company_name_list
