"""
功能：根据用户的输入内容，新建一个公司的文件夹，并将所有的报告保存其中
"""

import os
from docxtpl import DocxTemplate
from datetime import datetime, timedelta
import random


# 根据用户输入的日期字符串，返回两个日期（当前日期，检测到期日期）
def format_date(date_str):
    date = datetime.strptime(date_str, "%Y%m%d")
    formatted_date = date.strftime("%Y 年 %m  月 %d  日")
    next_year = date + timedelta(days=365)
    previous_day = next_year - timedelta(days=0)
    formatted_next_year_date = previous_day.strftime("%Y 年 %m  月 %d  日")
    return formatted_date, formatted_next_year_date


# 根据序号,返回三位的格式化编号，不足三位以0补充
def return_format_num(num):
    if num < 10:
        return f"00{num}"
    elif num < 100:
        return f"0{num}"
    else:
        return f"{num}"


# 根据日期和总体的探头序号，返回报告的文件编号
# 格式：ZJYX-202306140001
def get_file_num(date, num, start_num):
    return f"ZJYX-{date}0{return_format_num(num + start_num)}"


# 根据区域和各个区域的数量生成并返回所有的探头编号
def create_all_alerts_num_list(sections, sections_num):
    all_alerts_num = []  # 所有探头编号的列表
    for i in range(len(sections)):
        for j in range(sections_num[i]):
            alert_num = f"{sections[i]}{return_format_num(j + 1)}"
            all_alerts_num.append(alert_num)
    return all_alerts_num


# 根据用户的界面输入，输出一个列表，列表中是可以直接替换word模板中数据的公司信息
def all_company_to_save(user_input):
    all_company_data = []

    company_name = user_input["company_name"]
    all_nums = user_input["all_nums"]
    date = user_input["date"]
    temperature = user_input["temperature"]
    humidity = user_input["humidity"]
    sections = user_input["sections"]
    sections_num = user_input["sections_num"]
    start_num = user_input["start_num"]

    all_alerts_num = create_all_alerts_num_list(sections, sections_num)

    for i in range(all_nums):
        new_company = {
            "file_num": get_file_num(date, i, start_num),
            "company_name": company_name,
            "alert_type": "AEC2332",
            "alert_factory": "成都安可信电子股份有限公司",
            "dongzuozhi": "ankexindongzuo",
            "alert_num": all_alerts_num[i],
            "date_now": format_date(date)[0],
            "date_next": format_date(date)[1],
            "temperature": temperature,
            "humidity": humidity,
            "random_chongfu": round(random.uniform(0.0, 2.0), 1),
            "action_time": random.randint(7, 25)
        }
        all_company_data.append(new_company)

    return all_company_data


def write_save_all_company(one_company_input_data):
    os.makedirs(f"{one_company_input_data['company_name']}")
    doc = DocxTemplate("model.docx")
    all_new_company = all_company_to_save(one_company_input_data)
    for company in all_new_company:
        doc.render(company)
        doc.save(f"{one_company_input_data['company_name']}/{company['file_num']}-{company['alert_num']}.docx")
    return True


if __name__ == "__main__":
    user_input = {
        "company_name": "绍兴鑫沃电子有限公司",
        "all_nums": 10,
        "date": "20230614",
        "temperature": 30,
        "humidity": 35,
        "sections": ["厨房", "大厅"],
        "sections_num": [3, 7],
        "start_num": 7
    }
    write_save_all_company(user_input)
