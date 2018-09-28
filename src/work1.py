from openpyxl import Workbook

from openpyxl import load_workbook

from openpyxl.writer.excel import ExcelWriter

import openpyxl as xlsx
import pandas as pd
import numpy as np
import xlrd
import time
import logging
import math
import logging.config
import ioutil as il

logging.config.fileConfig("./logging.conf")
logger_name = "work1"
logger = logging.getLogger(logger_name)


def main():
    read_labels, result_file_name, sheet_name, statistic_label, company_name_list = il.read_config_work1()
    if len(read_labels) <= 0:
        logger.error('统计时间段(Read_Label)为空或配置出错。')
        return
    if result_file_name == ".xls":
        logger.error('汇总输出文件名(Result_File_Name)为空或配置出错。')
        return
    if sheet_name == "":
        logger.error('工作簿名称(Sheet_Name)为空或配置出错。')
        return
    if statistic_label == "":
        logger.error('汇总行标签(Statistic_Label)为空或配置出错。')
        return
    if len(company_name_list) <= 0:
        logger.error('公司名称(Company_name)为空或配置出错。')
        return
    file_index = 0
    statistic_count = 0
    df_headers = None
    file_dir = r"../data/"
    file_path_list = []
    il.listdir(file_dir, file_path_list, "path")
    statistic_df = pd.DataFrame()
    df_section_index_list = [2-1]
    for read_lable in read_labels:
        read_lable_index = 1
        for file_path in file_path_list:
            file_name = file_path.split("/")[-1]
            if file_name != result_file_name:
                try:
                    if "(" in file_name:
                        file_date = file_name.split("(")[0]
                        file_city = file_name.split("(")[1].split(")")[0]
                        file_company = file_name.split(")")[1].split(".")[0]
                    else:
                        file_date = file_name.split("（")[0]
                        file_city = file_name.split("（")[1].split("）")[0]
                        file_company = file_name.split("）")[1].split(".")[0]
                except Exception as e:
                    logger.error("文件名 " + file_name + " 命名格式错误！！！如果该文件不需要统计，请忽略。")
                    continue
                if file_company in company_name_list:
                    df_dict = pd.read_excel(file_path, sheet_name=None, header=None)
                    # try:
                    for k, v in df_dict.items():
                        if sheet_name in k:
                            if file_index == 0:
                                df_headers = v.iloc[3:5, :]
                            df_data = pd.concat((v.iloc[:, 0:2].fillna(method="pad"), v.iloc[:, 2:]), axis=1)
                            for index, row in df_data.iterrows():
                                if row[0] == read_lable and not row[1] == statistic_label:
                                    if len(set(row[2:])) <= 1 and len(
                                            set([i for i in list(row[2:]) if not math.isnan(i)])) < 1:
                                        # print(set(row[2:]))
                                        pass
                                    else:
                                        row[1] = read_lable_index
                                        df_headers = df_headers.append(row,ignore_index=True)
                                        read_lable_index += 1
                                elif row[0] == read_lable and row[1] == statistic_label and statistic_count == 0:
                                    statistic_df = row
                                    statistic_count += 1
                    file_index += 1
                    # except Exception as e:
                    #     logger.error('Failed to load the excel: ' + str(e))
                    #     logger.error("[" + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(
                    #         time.time())) + "] Please check the file " + file_name + ", its format is not right).")
        df_headers = df_headers.append(statistic_df, ignore_index=True)
        df_headers = df_headers.append(statistic_df, ignore_index=True)
        df_section_index_list.append(len(df_headers))
        if len(df_headers) > df_section_index_list[-2]:
            df_section = df_headers.iloc[df_section_index_list[-2]+1:df_section_index_list[-1]-1,:]
            df_headers.iloc[-1,2] = np.sum(df_section.iloc[:,2].values)
        statistic_count = 0
    df_headers.to_excel(file_dir + result_file_name, header=None, index=None)


if __name__ == "__main__":
    main()
