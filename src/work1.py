# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import math
import logging.config
import ioutil as il
from datetime import datetime

logging.config.fileConfig("./logging.conf")
logger_name = "work1"
logger = logging.getLogger(logger_name)


def isVaildDate(date):
    try:
        if ":" in date and "-" in date and date.count("-") == 2:
            result_time = datetime.strptime(date, "%Y-%m-%d %H:%M:%S")
        elif ":" in date and "-" in date and date.count("-") == 1:
            result_time = datetime.strptime(date, "%Y-%m %H:%M:%S")
        elif ":" in date and "/" in date and date.count("/") == 2:
            result_time = datetime.strptime(date, "%Y/%m/%d %H:%M:%S")
        elif ":" in date and "/" in date and date.count("/") == 1:
            result_time = datetime.strptime(date, "%Y/%m %H:%M:%S")
        elif ":" in date and "." in date and date.count(".") == 2:
            result_time = datetime.strptime(date, "%Y.%m.%d %H:%M:%S")
        elif ":" in date and "." in date and date.count(".") == 1:
            result_time = datetime.strptime(date, "%Y.%m %H:%M:%S")
        elif "-" in date and date.count("-") == 2:
            result_time = datetime.strptime(date, "%Y-%m-%d")
        elif "-" in date and date.count("-") == 1:
            result_time = datetime.strptime(date, "%Y-%m")
        elif "/" in date and date.count("/") == 2:
            result_time = datetime.strptime(date, "%Y/%m/%d")
        elif "/" in date and date.count("/") == 1:
            result_time = datetime.strptime(date, "%Y/%m")
        elif "." in date and date.count(".") == 2:
            result_time = datetime.strptime(date, "%Y.%m.%d")
        elif "." in date and date.count(".") == 1:
            result_time = datetime.strptime(date, "%Y.%m")
        else:
            result_time = datetime.strptime(date, "%Y/%m/%d")
        return True, result_time
    except:
        return False, None


def isfloat(str):
    try:
        float(str)
    except ValueError:
        return False
    return True


def get_loan_time_limit(loan_time_limit, loan_start_date):
    loan_time_limit_in_month = None
    if type(loan_time_limit) == str:
        if "年" in loan_time_limit:
            # print(loan_time_limit,list(filter(str.isdigit, loan_time_limit))[0])
            loan_time_limit_in_number = int(''.join(list(filter(str.isdigit, loan_time_limit.split("年")[0]))))
            loan_time_limit_in_month = loan_time_limit_in_number * 12
        elif "月" in loan_time_limit:
            loan_time_limit_in_number = int(''.join(list(filter(str.isdigit, loan_time_limit.split("月")[0]))))
            loan_time_limit_in_month = loan_time_limit_in_number
        elif "天" in loan_time_limit:
            loan_time_limit_in_number = int(''.join(list(filter(str.isdigit, loan_time_limit.split("天")[0]))))
            loan_time_limit_in_month = loan_time_limit_in_number / 30
        elif "日" in loan_time_limit:
            loan_time_limit_in_number = int(''.join(list(filter(str.isdigit, loan_time_limit.split("日")[0]))))
            loan_time_limit_in_month = loan_time_limit_in_number / 30
        elif isVaildDate(loan_time_limit)[0]:
            loan_end_date = isVaildDate(loan_time_limit)[1]
            delta = loan_end_date - loan_start_date
            loan_time_limit_in_month = delta.days / 30
    elif isinstance(loan_time_limit, datetime):
        delta = loan_time_limit - datetime.strptime(loan_start_date, "%Y-%m-%d")
        loan_time_limit_in_month = delta.days / 30
    if loan_time_limit_in_month is not None:
        return True, int(loan_time_limit_in_month)
    else:
        return False, None


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
    df_section_index_list = [2 - 1]
    for read_lable in read_labels:
        read_lable_index = 1
        this_df_data = pd.DataFrame()
        this_df_tiles = pd.DataFrame()
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
                    find_sheet = False
                    for k, v in df_dict.items():
                        if sheet_name in k:
                            find_sheet = True
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
                                        append_flag = True
                                        row[1] = read_lable_index
                                        # 检查填写是否为数字的情况相关内容
                                        for number_check_index in [2, 5, 6, 7, 8, 11, 14, 15, 16, 17, 23, 24, 26, 29,
                                                                   30]:
                                            number_value = row[number_check_index]
                                            if isfloat(number_value):
                                                pass
                                            else:
                                                append_flag = False
                                                logger.error(
                                                    "文件 " + file_name + " 中第" + str(index + 1) + "行(第" + str(
                                                        number_check_index + 1) + ")列应使用全数字填写，目前表格不符合要求，未计入统计")
                                        # 处理发放日相关内容
                                        for date_index in [3, 12, 21, 27]:
                                            if isinstance(row[date_index], datetime):
                                                row[date_index] = datetime.strftime(row[date_index], "%Y-%m-%d")
                                            elif isVaildDate(row[date_index])[0]:
                                                row[date_index] = datetime.strftime(isVaildDate(row[date_index])[1],
                                                                                    "%Y-%m-%d")
                                            elif isfloat(row[date_index]) and isfloat(
                                                    row[date_index - 1]) and math.isnan(row[date_index]) and math.isnan(
                                                row[date_index - 1]):
                                                pass
                                            else:
                                                append_flag = False
                                                logger.error(
                                                    "文件 " + file_name + " 中第" + str(index + 1) + "行发放日(第" + str(
                                                        date_index + 1) + ")列格式错误,未计入统计")
                                        # 处理贷款期限相关内容
                                        for loan_time_limit_index in [4, 13, 22, 28]:
                                            loan_time_limit = row[loan_time_limit_index]
                                            loan_time_limit_flag, loan_time_limit_in_month = get_loan_time_limit(
                                                loan_time_limit, row[loan_time_limit_index - 1])
                                            if loan_time_limit_flag:
                                                row[loan_time_limit_index] = loan_time_limit_in_month
                                            elif isfloat(row[loan_time_limit_index]) and isfloat(
                                                    row[loan_time_limit_index - 2]) and math.isnan(
                                                row[loan_time_limit_index]) and math.isnan(
                                                row[loan_time_limit_index - 2]):
                                                pass
                                            else:
                                                append_flag = False
                                                logger.error(
                                                    "文件 " + file_name + " 中第" + str(index + 1) + "行贷款期限(第" + str(
                                                        loan_time_limit_index + 1) + ")列格式错误,未计入统计")

                                        # 处理政策资金补偿相关内容
                                        for other_income_index in [9]:
                                            other_income = row[other_income_index]
                                            if other_income == "无" or other_income == "否":
                                                row[other_income_index] = 0
                                            elif isfloat(row[other_income_index]) and math.isnan(
                                                    row[other_income_index]):
                                                row[other_income_index] = 0
                                            elif isfloat(row[other_income_index]):
                                                pass
                                            else:
                                                append_flag = False
                                                logger.error(
                                                    "文件 " + file_name + " 中第" + str(
                                                        index + 1) + "行政策补偿金(第" + str(
                                                        other_income_index + 1) + ")列格式错误,未计入统计")
                                        if append_flag:
                                            sub_append_flag = True
                                            # 处理存在月数相关
                                            active_month = [1, 1, 1, 1]
                                            first_date_quarter_dic = {
                                                1: "-01-01",
                                                2: "-04-01",
                                                3: "-07-01",
                                                4: "-10-01"
                                            }
                                            last_date_last_month_quarter_dic = {
                                                1: "-03-31",
                                                2: "-06-30",
                                                3: "-09-30",
                                                4: "-12-31"
                                            }
                                            last_date_month_dic = {
                                                1: "31",
                                                2: "28",
                                                3: "31",
                                                4: "30",
                                                5: "31",
                                                6: "30",
                                                7: "31",
                                                8: "31",
                                                9: "30",
                                                10: "31",
                                                11: "30",
                                                12: "31"
                                            }
                                            if "季度" in row[0]:
                                                if "(" in row[0]:
                                                    current_quarter = int(row[0].split("(")[0][-3])
                                                elif "（" in row[0]:
                                                    current_quarter = int(row[0].split("（")[0][-3])
                                                else:
                                                    current_quarter = int(row[0][-3])
                                                current_year = row[0][:4]
                                                current_month_date = datetime.strptime(
                                                    current_year + last_date_last_month_quarter_dic[current_quarter],
                                                    "%Y-%m-%d")
                                                file_month_date = datetime.strptime(
                                                    file_date + last_date_month_dic[int(file_date[-2:])], "%Y%m%d")
                                                if file_month_date > current_month_date:
                                                    file_month_date = current_month_date
                                                file_year = file_date[:4]
                                                loop_index = 0
                                                for date_index in [3, 12, 21, 27]:
                                                    if not (isfloat(row[date_index - 1]) and math.isnan(
                                                            row[date_index - 1])):
                                                        loan_time_limit_in_month = row[date_index + 1]
                                                        start_date = row[date_index]
                                                        delta = file_month_date - datetime.strptime(start_date,
                                                                                                    "%Y-%m-%d")
                                                        if delta.days < 90:
                                                            active_month[loop_index] = int((delta.days + 1) / 30)
                                                        elif delta.days >= 90 and delta.days < loan_time_limit_in_month * 30:
                                                            current_quarter_delta = file_month_date - datetime.strptime(
                                                                current_year + first_date_quarter_dic[current_quarter],
                                                                "%Y-%m-%d")
                                                            active_month[loop_index] = int(
                                                                (current_quarter_delta.days + 1) / 30)
                                                        else:
                                                            sub_append_flag = False
                                                            logger.error(
                                                                "文件 " + file_name + " 中第" + str(
                                                                    index + 1) + "行(第" + str(
                                                                    date_index + 1) + ")列对应贷款已经超过贷款期限一月以上,未计入统计")
                                                    else:
                                                        active_month[loop_index] = 0
                                                    loop_index += 1
                                            elif "月" in row[0]:
                                                loop_index = 0
                                                for date_index in [3, 12, 21, 27]:
                                                    if not (isfloat(row[date_index - 1]) and math.isnan(
                                                            row[date_index - 1])):
                                                        loan_time_limit_in_month = row[date_index + 1]
                                                        start_date = row[date_index]
                                                        delta = file_month_date - datetime.strptime(start_date,
                                                                                                    "%Y-%m-%d")
                                                        if delta.days < loan_time_limit_in_month * 30:
                                                            pass
                                                        else:
                                                            sub_append_flag = False
                                                            logger.error(
                                                                "文件 " + file_name + " 中第" + str(
                                                                    index + 1) + "行(第" + str(
                                                                    date_index + 1) + ")列对应贷款已经超过贷款期限一月以上,未计入统计")
                                                    else:
                                                        active_month[loop_index] = 0
                                                    loop_index += 1
                                            else:
                                                sub_append_flag = False
                                                logger.error(
                                                    "文件 " + file_name + " 中第" + str(
                                                        index + 1) + "行(第" + str(
                                                        1) + ")命名格式错误,未计入统计")
                                            row = pd.Series(
                                                np.concatenate([row.values, np.array(active_month)], axis=0))
                                            if sub_append_flag:
                                                row = row.fillna(0)
                                                this_df_data = this_df_data.append(row, ignore_index=True)
                                                read_lable_index += 1
                                elif row[0] == read_lable and row[1] == statistic_label and statistic_count == 0:
                                    statistic_df = row
                                    statistic_count += 1
                    if not find_sheet:
                        logger.error("文件 " + file_name + " 中不包含‘表3-人行填写’Sheet页。")
                    file_index += 1
                    # except Exception as e:
                    #     logger.error('Failed to load the excel: ' + str(e))
                    #     logger.error("[" + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(
                    #         time.time())) + "] Please check the file " + file_name + ", its format is not right).")
        if len(this_df_data) > 0 and "预计" not in this_df_data.iloc[-1, 0]:
            this_df_tiles = this_df_tiles.append(statistic_df, ignore_index=True)
            this_df_tiles = this_df_tiles.append(statistic_df, ignore_index=True)  # 多附加一行作为占位用的
            # 贷款金额汇总
            for summary_index in [2, 11, 19, 26]:
                this_df_tiles.iloc[-1, summary_index] = np.sum(this_df_data.iloc[:, summary_index].values)
            # 贷款利率汇总
            loan_amount_dic = {5: 2, 14: 11, 23: 19, 29: 26}
            for summary_index in [5, 14, 23, 29]:
                this_df_tiles.iloc[-1, summary_index] = np.sum(
                    np.multiply(this_df_data.iloc[:, loan_amount_dic[summary_index]].values,
                                this_df_data.iloc[:, summary_index].values)) / this_df_tiles.iloc[
                                                            -1, loan_amount_dic[summary_index]]
            # 担保费汇总
            summary_loop_index = 0
            loan_time_limit_dic = {6: 4, 15: 13}
            for summary_index in [6, 15]:
                this_df_tiles.iloc[-1, summary_index] = np.sum(
                    np.multiply(np.divide(this_df_data.iloc[:, summary_index].values,
                                          this_df_data.iloc[:, loan_time_limit_dic[summary_index]].values),
                                this_df_data.iloc[:, -4 + summary_loop_index].values))
                summary_loop_index += 1
            # 评估费汇总
            summary_loop_index = 0
            loan_time_limit_dic = {7: 4, 16: 13}
            for summary_index in [7, 16]:
                this_df_tiles.iloc[-1, summary_index] = np.sum(
                    np.multiply(np.divide(this_df_data.iloc[:, summary_index].values,
                                          this_df_data.iloc[:, loan_time_limit_dic[summary_index]].values),
                                this_df_data.iloc[:, -4 + summary_loop_index].values))
                summary_loop_index += 1
            # 其他费用汇总
            summary_loop_index = 0
            loan_time_limit_dic = {8: 4, 17: 13, 24: 22, 30: 28}
            for summary_index in [8, 17, 24, 30]:
                this_df_tiles.iloc[-1, summary_index] = np.sum(
                    np.multiply(np.divide(this_df_data.iloc[:, summary_index].values,
                                          this_df_data.iloc[:, loan_time_limit_dic[summary_index]].values),
                                this_df_data.iloc[:, -4 + summary_loop_index].values))
                summary_loop_index += 1
            # 政策补偿金汇总
            summary_loop_index = 0
            loan_time_limit_dic = {9: 4}
            for summary_index in [9]:
                this_df_tiles.iloc[-1, summary_index] = np.sum(
                    np.multiply(
                        np.divide(this_df_data.iloc[:, summary_index].values,
                                  this_df_data.iloc[:, loan_time_limit_dic[summary_index]].values),
                        this_df_data.iloc[:, -4 + summary_loop_index].values))
                summary_loop_index += 1
            # 当期贷款融资费用
            summary_loop_index = 0
            this_df_tiles.iloc[-1, :] = this_df_tiles.iloc[-1, :].fillna(0)
            for summary_index in [10, 18, 25, 31]:
                if summary_index == 10:
                    current_interest = np.sum(
                        np.multiply(np.multiply(this_df_data.iloc[:, 2], this_df_data.iloc[:, 5]) / 12,
                                    this_df_data.iloc[:, -4 + summary_loop_index]))
                    this_df_tiles.iloc[-1, summary_index] = current_interest + np.sum(
                        this_df_tiles.iloc[-1, 6:10].values)
                elif summary_index == 18:
                    current_interest = np.sum(
                        np.multiply(np.multiply(this_df_data.iloc[:, 11], this_df_data.iloc[:, 14]) / 12,
                                    this_df_data.iloc[:, -4 + summary_loop_index]))
                    this_df_tiles.iloc[-1, summary_index] = current_interest + np.sum(
                        this_df_tiles.iloc[-1, 15:18].values)
                elif summary_index == 25:
                    current_interest = np.sum(
                        np.multiply(np.multiply(this_df_data.iloc[:, 19], this_df_data.iloc[:, 23]) / 12,
                                    this_df_data.iloc[:, -4 + summary_loop_index]))
                    this_df_tiles.iloc[-1, summary_index] = current_interest + np.sum(
                        this_df_tiles.iloc[-1, 24])
                elif summary_index == 31:
                    current_interest = np.sum(
                        np.multiply(np.multiply(this_df_data.iloc[:, 26], this_df_data.iloc[:, 29]) / 12,
                                    this_df_data.iloc[:, -4 + summary_loop_index]))
                    this_df_tiles.iloc[-1, summary_index] = current_interest + np.sum(
                        this_df_tiles.iloc[-1, 30])
                summary_loop_index += 1
        df_headers = pd.concat([df_headers, this_df_data, this_df_tiles], axis=0)
        statistic_count = 0
    df_headers.to_excel(file_dir + result_file_name, header=None, index=None)


if __name__ == "__main__":
    main()
    str = input(u"按两次回车键关闭窗口。")
