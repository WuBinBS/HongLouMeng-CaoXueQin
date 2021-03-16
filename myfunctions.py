# coding: UTF-8
import docx
from scipy import stats
from prettytable import PrettyTable
import numpy as np


# 函数功能: 读入初步处理后的《红楼梦》, 并进行数据清洗和整理
def statistics_reading_cleaning_arranging():
    # 读入原《红楼梦》Word文档删去位于文档前后部分的目录和其他冗余信息后剩下的部分
    file = docx.Document('A:\\shuli\\《红楼梦》删.docx')

    # 将《红楼梦》一百二十回的内容装入一个字典hlm, 第x回对应的键为'第x回', 键的值为该回的正文内容（不分段落）
    hlm = {}

    # 进行第一层处理: 1.检测关键字以划分一百二十回;
    # 2. 清洗掉每回前的回前墨部分， 每回之后的回后评部分.

    switch = 1  # 控制是否将当前光标所指的文字收集起来, 1表示收集, 0表示不收集
    chapter = 1  # 定位当前的回

    for para in file.paragraphs:

        # 若本段为空段, 则忽略本段内容, 且若之前遇到'回前墨'或'回后评', 则将文字收集开关打开
        if para.text == '':
            if switch == 0:
                switch = 1
            continue

        # 若本段非空段的情况
        else:
            # 检测本段开头有无关键字'第', ’回'
            # 若检测到, 则在字典创建下一回的键, 初始化其值为空列表[]
            # 并设置文字收集开关为1
            if para.text[0] == '第' and para.text[-1] == '回':
                location = '第' + str(chapter) + '回'
                chapter += 1
                hlm[location] = ''

            # 检测本段开头有无关键字'回前墨'或'回后评'
            # 若检测到, 则将文字收集开关调整为0
            elif '回前墨' in para.text[0:min(5, len(para.text))] or \
                    '回后评' in para.text[0:min(5, len(para.text))]:
                switch = 0

        # 若文字收集开关switch是1, 则将本段内容中除了空格外的内容收录到字典hlm对应的回中
        if switch == 1:
            for word in para.text:
                if word != ' ':
                    hlm[location] += word

    # 进行第二层处理: 去除每一回最后的注释
    # 先针对[数字]的注释
    for k in range(1, 121):
        t = []
        for pos in range(0, len(hlm['第' + str(k) + '回']) - 3):
            if hlm['第' + str(k) + '回'][pos: pos + 3] == '[1]' or \
                    hlm['第' + str(k) + '回'][pos: pos + 3] == '【1】':
                t.append(pos)
        if t != []:
            hlm['第' + str(k) + '回'] = hlm['第' + str(k) + '回'][0:t[-1]]

    # 再针对[汉字]的注释
    for k in range(1, 121):
        t = []
        for pos in range(0, len(hlm['第' + str(k) + '回']) - 3):
            if hlm['第' + str(k) + '回'][pos: pos + 3] == '[一]' or \
                    hlm['第' + str(k) + '回'][pos: pos + 3] == '【一】':
                t.append(pos)
        if t != []:
            hlm['第' + str(k) + '回'] = hlm['第' + str(k) + '回'][0:t[-1]]

    # 进行第三层处理: 去除每一回正文中的注释角标[x]
    switch = 1  # 控制是否保留文字的开关, 1表示保留, 0表示不保留
    for k in range(1, 121):
        t = ''
        for word in hlm['第' + str(k) + '回']:
            if word == '[' or word == '【':
                switch = 0
            elif word == ']' or word == '】':
                switch = 1
                continue

            if switch == 1:
                t += word
        hlm['第' + str(k) + '回'] = t

    return hlm


# 函数功能: 进行单因素方差分析
def anova1(data, alpha=0.05):
    # data是单因素各水平抽样数据, alpha是显著性水平
    s = len(data)  # 水平数
    n_j = []  # 各水平下的样本容量
    n = 0  # 样本总容量
    for item in data.values():
        n_j.append(len(item))
        n += n_j[-1]

    X_mean = 0
    for item in data.values():
        X_mean += sum(np.array(item))
    X_mean = X_mean / n

    SST = 0
    for item in data.values():
        SST += sum(np.array(item) ** 2)
    SST -= n * (X_mean ** 2)

    SSA = 0
    k = 0
    for item in data.values():
        SSA += n_j[k] * (np.mean(np.array(item)) ** 2)
        k += 1
    SSA -= n * (X_mean ** 2)

    SSE = SST - SSA

    table = PrettyTable()
    table.add_column('方差来源', ['因素', '误差', '总和'])
    table.add_column('平方和', [SSA, SSE, SST])
    table.add_column('自由度', [s - 1, n - s, n - 1])
    table.add_column('均方', [SSA / (s - 1), SSE / (n - s), '_'])
    table.add_column('F比', [(SSA * (n - s)) / (SSE * (s - 1)), '_', '_'])
    table.add_column('F临界值', [stats.f.isf(alpha, dfn=s - 1, dfd=n - s), '_', '_'])

    return table


# 函数功能: 进行双边的秩和检验
def kruskal_wallis_bilateral_test_over10(data, alpha=0.05):
    # 要求检验的总体个数为2, 且n1, n2均不小于10
    if len(data) != 2:
        return

    all1 = list(data.values())[0]
    all2 = list(data.values())[1]
    n1 = len(all1)
    n2 = len(all2)
    if n1 < 10 or n2 < 10:
        return

    # 计算各个观察值的秩
    all = np.sort(np.array(all1 + all2))
    rank = np.array(range(1, n1 + n2 + 1))
    group = []
    backup = []
    for item in all:
        if item not in backup:
            pos = np.where(all == item)[0]
            if len(pos) > 1:
                backup.append(item)
                group.append(len(pos))
                t = np.sum(all[pos]) / len(pos)
                rank[pos] = t

    # 计算来自第1个总体的秩和
    R1 = 0
    for k in range(n1 + n2):
        if all[k] in all1:
            R1 += rank[k]

    # 计算正态总体的均值和方差
    mju = n1 * (n1 + n2 + 1) / 2
    if group == []:
        sigma = np.sqrt(n1 * n2 * (n1 + n2 + 1) / 12)
    else:
        n = n1 + n2
        t = np.sum((np.array(group, dtype='float') ** 3)) - \
            np.sum(np.array(group, dtype='float'))
        sigma = np.sqrt(
            n1 * n2 * (n * (n ** 2 - 1) - t) / (12 * n * (n - 1))
        )

    # 计算拒绝域
    t = stats.norm.isf(alpha)
    C_L = mju - t * sigma
    C_U = mju + t * sigma
    print('秩相同的组的组数为:', len(group))
    print('第一样本的秩和R1的观察值为r1=', R1)
    print('拒绝域为: r1 < ' + str(C_L) + ', r1 > ' + str(C_U))
    if R1 < C_L or R1 > C_U:
        print('两个总体的数据有显著差异')
    else:
        print('两个总体的数据无显著差异')

    return R1, C_L, C_U, len(group)


# 函数功能: 在单因素方差分析结果是有显著差异的情况下, 计算效应尺度的大小
# 限制: 要求单因素的水平数为2, 即只有两组
def effect_scale(data):
    if len(data) != 2:
        return
    t1 = np.array(list(data.values())[0])
    t2 = np.array(list(data.values())[1])
    t = 2 * np.abs(np.mean(t1) - np.mean(t2)) / (np.std(t1, ddof=1) + np.std(t2, ddof=1))
    print('两组样本的效应尺度为', t)

    if t < 0.2:
        print('差异无实际意义')
    elif t < 0.5:
        print('差异较小')
    elif t < 0.8:
        print('差异程度为中等')
    else:
        print('差异较大')

    return t
