"""

用于读取最大的模型连续使用量，给出最大环数，起始环号，结束环号
程鹏      2022/10/10
Version 1.0

"""


import pandas as pd


def excel_one_line_to_list():
    """
    读取excel的某一列数据
    :return:
    """
    df = pd.read_excel("自动情况20221010.xlsx", usecols=[7],
                       names=None)  # 读取项目名称列,不要列名
    df_li = df.values.tolist()
    my_excel_line = []
    for s_li in df_li:
        my_excel_line.append(s_li[0])
    # print(my_excel_line)
    return my_excel_line

def find_max_continue(org):
    """

    :param org: 输入起始环号
    :return: 最大连续使用量，起始环号，结束环号
    """
    count = 0
    count_max = 0
    finish = 0
    i = 0
    my_list = excel_one_line_to_list()
    while i < len(my_list):
        if my_list[i] > 0:
            count += 1
            if count > count_max:
                count_max = count
                finish = i+1            # 对应位置位加一才是此数据的位置
        else:
            count = 0
        i += 1
    start = finish - count_max
    print(f"此列表中的模型最大连续使用{count_max}环，起始于{start + org}环，结束于{finish + org -1}环")
    return count_max, start + org, finish + org - 1


if __name__ == '__main__':
    # excel_one_line_to_list()
    result = find_max_continue(438)
    print(result)
