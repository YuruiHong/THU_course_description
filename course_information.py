import requests
import pandas as pd
import numpy as np
from urllib.parse import quote
from bs4 import BeautifulSoup

# 读取Excel文件，无Header，可能有空格需要处理
excel_file = '2024123456_undergraduate major_en.xlsx' # 使用下载的默认格式，替换为自己的文件名
df = pd.read_excel(excel_file, header=None)

# 假设课程名在第一列
course_names = df.iloc[:, 0].tolist()

# 设置固定的URL部分
base_url = "https://zhjwxkyw.cic.tsinghua.edu.cn/xkYjs.vxkYjsJxjhBs.do?m=kkxxSearch&page=-1&token=f1440a358d6fe6e27572a3895645ef14&p_sort.p1=&p_sort.p2=&p_sort.asc1=true&p_sort.asc2=true&p_xnxq={}&pathContent=Course+information&showtitle=&p_kch=&p_kcm={}&p_zjjsxm=&p_kkdwnm=&p_kcflm=&p_skxq=&p_skjc=&p_xkwzsm=&p_rxklxm=&p_kctsm=&p_ssnj=&p_bkskyl_ig=&p_yjskyl_ig=&goPageNumber=1"

# 请求头
headers = {
    'Cookie': 'JSESSIONID=xxx; serverid=xxx' # 替换为实际的Cookie值，注意zhjwxkyw.cic.tsinghua.edu.cn和zhjwxk.cic.tsinghua.edu.cn 不共享Cookie
}

# 定义一个子函数来提取学时和课程内容
def extract_course_details(course_url):
    # 请求课程详细页面
    course_response = requests.get(course_url, headers=headers)
    
    # 检查响应状态
    if course_response.status_code == 200:
        soup = BeautifulSoup(course_response.content, 'html.parser')
        
        # 找到所有 <td> 元素
        td_elements = soup.find_all('td')
        
        # 假设学时位于第6个 <td>，课程内容位于第11个 <td>，手动计算索引
        try:
            total_hours = str(td_elements[6].text.strip())  # 第7个 <td> 存储学时
            df.iloc[i, 2] = total_hours
            course_description = str(td_elements[10].text.strip())  # 第12个 <td> 存储课程简介
            df.iloc[i, 3] = course_description
            # english_description = str(td_elements[12].text.strip())  # 第13个 <td> 存储课程简介
            # if english_description: # 如果有英文描述，用英文描述替换中文描述
            #     df.iloc[i, 3] = english_description
        except IndexError:
            print("Failed to extract details, please check the HTML structure.")
    else:
        print(f"Failed to retrieve course details from {course_url}, Status Code: {course_response.status_code}")

# 遍历课程名称并获取详细信息
for semester in ['2024-2025-1', '2023-2024-2', '2023-2024-3']: # 需要手动设置学期范围
    for i, course in enumerate(course_names):
        # 将课程名称转为URL编码
        encoded_course = quote(str(course))

        # 构建完整的URL
        url = base_url.format(semester, encoded_course)
        
        # 发送GET请求
        response = requests.get(url, headers=headers)
        
        # 检查响应状态
        if response.status_code == 200:
            # 解析HTML内容
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # 提取class为trr2的元素
            trr2_elements = soup.find_all(class_='trr2')
            
            if trr2_elements:
                # 找到第一个trr2元素
                trr2_element = trr2_elements[0]
                
                # 找到第一个带有a标签的td元素，并获取href
                a_tag = trr2_element.find('a', class_='mainHref')
                if a_tag and a_tag.get('href'):
                    # 获取href链接并构建完整URL
                    href_link = a_tag.get('href')
                    full_url = "https://zhjwxkyw.cic.tsinghua.edu.cn/" + href_link
                    
                    # 调用子函数提取详细信息
                    extract_course_details(full_url)
            else:
                print(f"No trr2 element found for course: {course}")

df.to_excel(excel_file, index=False, header=False)