# coding: UTF-8
import docx
import myfunctions as my
import matplotlib.pyplot as plt
from prettytable import PrettyTable
import numpy as np
from scipy import stats
import statsmodels.api as sm

# 使用说明: 该.py文件需要配合myfunctions.py和《红楼梦》删,docx
# 直接运行即可, 所有结果将会输出

# 读入数据并清洗和整理
hlm = my.statistics_reading_cleaning_arranging()

# 读入红楼梦原版资料并统计原版资料总字数
file = docx.Document('A:\\shuli\\《红楼梦》删.docx')
word_symbol_original_count_all = 0
for para in file.paragraphs:
    if para.text != '':
        word_symbol_original_count_all += len(para.text)
del file, para

# 统计量: 每一回的字数和标点符号总数, 并可视化
word_symbol_count = {}
for k in range(1, 121):
    word_symbol_count['第' + str(k) + '回'] = len(hlm['第' + str(k) + '回'])
del k

plt.figure()
plt.bar(range(1, 41), list(word_symbol_count.values())[0:40])
plt.bar(range(41, 81), list(word_symbol_count.values())[40:80])
plt.bar(range(81, 121), list(word_symbol_count.values())[80:120])
plt.show()

plt.figure()
plt.boxplot(word_symbol_count.values())
plt.show()

t = list(word_symbol_count.values()).index(np.min(list(word_symbol_count.values())))
print('字符数最少的回数是', list(word_symbol_count.keys())[t],
      '字符数为', np.min(list(word_symbol_count.values())))
t = list(word_symbol_count.values()).index(np.max(list(word_symbol_count.values())))
print('字符数最多的回数是', list(word_symbol_count.keys())[t],
      '字符数为', np.max(list(word_symbol_count.values())))
print('---------------------------------------------------------------------\n')
del t

# 计算结果并可视化: 计算清洗掉的数据占原版资料的比重, 然后画成饼图
table = PrettyTable()
t = sum(list(word_symbol_count.values()))
table.add_column(' ', ['清洗前的数据共有字数', '清洗后的数据共有字数', '清洗掉的数据占原数据比重'])
table.add_column(' ', [word_symbol_original_count_all, t,
                       (word_symbol_original_count_all - t) / word_symbol_original_count_all])
print(table)
print('---------------------------------------------------------------------\n')

plt.figure()
plt.pie([t, word_symbol_original_count_all - t],
        explode=[0.2, 0.05],
        labels=['Words After Data Washing', 'Words Washed'],
        colors=['lightblue', 'lightgray'],
        autopct='%1.2f%%',
        shadow=True,
        startangle=100
        )
plt.axis('equal')
plt.show()
del table, t

# 提取数据: 给出前1-40回和41-80回单字使用次数的字典
hlm_words_1_40 = {}
hlm_words_41_80 = {}
for k in range(1, 41):
    for word in hlm['第' + str(k) + '回']:
        if word in hlm_words_1_40.keys():
            hlm_words_1_40[word] += 1
        else:
            hlm_words_1_40[word] = 1
    for word in hlm['第' + str(k + 40) + '回']:
        if word in hlm_words_41_80.keys():
            hlm_words_41_80[word] += 1
        else:
            hlm_words_41_80[word] = 1
del k, word

# 比对1-40回和41-80回各自的前k个高频字中有多少个重复的
hot_words_1_40 = []
hot_words_41_80 = []
k = 100

t1 = sorted(list(hlm_words_1_40.values()), reverse=True)[0:k]
t2 = hlm_words_1_40.copy()
for value in t1:
    word = list(t2.keys())[list(t2.values()).index(value)]
    hot_words_1_40.append(word)
    del t2[word]

t1 = sorted(list(hlm_words_41_80.values()), reverse=True)[0:k]
t2 = hlm_words_41_80.copy()
for value in t1:
    word = list(t2.keys())[list(t2.values()).index(value)]
    hot_words_41_80.append(word)
    del t2[word]

del k, t1, t2, value, word

num = 0
for item in hot_words_1_40:
    if item in hot_words_41_80:
        num += 1
        print('1-40回和41-80回中均有高频字 ' + item)
print('1-40回和41-80回中, 各自的前', len(hot_words_1_40), '高频字重复的有', num, '个')
print('---------------------------------------------------------------------\n')
del num, item

# 提取数据: 给出前80回的所有单字使用次数的字典, 为后面提取高频用字提供基础
hlm_words_before80 = {}
for k in range(1, 81):
    for word in hlm['第' + str(k) + '回']:
        if word in hlm_words_before80.keys():
            hlm_words_before80[word] += 1
        else:
            hlm_words_before80[word] = 1
del k, word

# 提取数据: 考虑前k1个高频使用字符, 并且列成表格
k1 = k2 = 100
t1 = sorted(list(hlm_words_before80.values()), reverse=True)[0:k1]
t2 = hlm_words_before80.copy()
hot_words_before80 = []
table = PrettyTable(['高频使用字符', '使用次数'])
for value in t1:
    word = list(t2.keys())[list(t2.values()).index(value)]
    hot_words_before80.append(word)
    table.add_row([word, value])
    k2 -= 1
    del t2[word]
print(table)
print('这' + str(k1) + '个高频使用字占前80回篇幅的百分比为: ' + str(sum(t1) / sum(list(hlm_words_before80.values()))))
print('---------------------------------------------------------------------\n')

# 计算: 将1-40回和41-80回整合在一起时, 其前k高频字既是1-40回前k高频字, 又是41-80回前k高频字的个数
num = 0
for item in hot_words_before80:
    if item in hot_words_1_40 and item in hot_words_41_80:
        num += 1
print('前80回前' + str(len(hot_words_before80)) + '个高频字中, 既是1-40回前' +
      str(len(hot_words_before80)) + '高频字, 又是41-80回中前' +
      str(len(hot_words_before80)) + '高频字的字有 ' + str(num) + '个')
print('---------------------------------------------------------------------\n')
del num, item

# 分析: 对前k1个高频使用字符作箱线图和直方图, 以决定要使用哪些字
plt.figure()
plt.hist(t1, bins=55)
plt.show()

plt.figure()
plt.boxplot(t1)
plt.show()
del k1, k2, t1, t2, value, word, table

# 数据提取: 经过箱线图和直方图的分析, 决定使用的高频字
hot_words_before80_picked = hot_words_before80[
                            hot_words_before80.index('去'):
                            ]
hot_words_before80_picked.remove('？')
hot_words_before80_picked.remove('、')
hot_words_before80_picked.remove('！')

# 计算: 决定使用的高频字占前80回的篇幅
hot_words_before80_picked_count = 0
for word in hot_words_before80_picked:
    hot_words_before80_picked_count += hlm_words_before80[word]
table = PrettyTable()
table.add_row(['筛选得高频字总字数', hot_words_before80_picked_count])
t = sum(list(word_symbol_count.values()))
table.add_row(['前80回总字数', t])
table.add_row((['筛选所得高频字占前80回比重',
                hot_words_before80_picked_count / t]))
print(table)

plt.figure()
plt.pie([hot_words_before80_picked_count, t - hot_words_before80_picked_count],
        explode=[0.2, 0.05],
        labels=['Frequently Used Picked', 'The Rest'],
        colors=['lightblue', 'lightgray'],
        autopct='%1.2f%%',
        shadow=True,
        startangle=100
        )
plt.axis('equal')
plt.show()
del table, t, word

# 提取数据: 统计全120回中每一回各句长出现的频数（除去每一回的标题）
# 句子分割标准: 逗号, 句号, 问号, 感叹号, 冒号，顿号，分号，省略号, 其他标点符号算作字
sentence_len_120 = {}
t = 0  # 句长字数统计中间变量, 当进入下一句时变为0
for k in range(1, 121):
    sentence_len_120['第' + str(k) + '回'] = {}
    for word in hlm['第' + str(k) + '回'][19:]:
        if word in [',', '，', '.', '。', '?', '？', '!', '！', ':', '：', '、', '；', ';', '…']:
            if t == 0:
                continue
            elif t in sentence_len_120['第' + str(k) + '回'].keys():
                sentence_len_120['第' + str(k) + '回'][t] += 1
                t = 0
            else:
                sentence_len_120['第' + str(k) + '回'][t] = 1
                t = 0
        else:
            t += 1
del t, k, word

# 抽样: 在1到80回的每一回随机抽取一个长度为sampling_len的句段
# 统计其中拥有筛选得高频字的个数, 每一回抽取iter_times次
data1 = {}
sampling_len = 100
iter_times = 50
for k in range(1, 81):
    data1['第' + str(k) + '回'] = []
    for iter in range(iter_times):
        data1['第' + str(k) + '回'].append(0)
        pos = int(np.random.rand() *
                  (len(hlm['第' + str(k + 1) + '回']) - sampling_len))
        for item in hlm['第' + str(k + 1) + '回'][pos:pos + sampling_len]:
            if item in hot_words_before80_picked:
                data1['第' + str(k) + '回'][-1] += 1
del sampling_len, iter_times, k, iter, pos, item

# 打印出抽样的结果
for k in range(80):
    print(list(data1.keys())[k], list(data1.values())[k])
print('---------------------------------------------------------------------\n')

# 分析: 对data1中每一回对应的抽样来自的总体作K-S正态性检验, 然后可视化
not_norm = []
t1 = []
t2 = []
for k in range(1, 81):
    t = stats.kstest(data1['第' + str(k) + '回'], 'norm',
                     args=(
                         np.mean(data1['第' + str(k) + '回']),
                         np.std(data1['第' + str(k) + '回'], ddof=1)
                     ))
    t1.append(t[0])
    t2.append(t[1])
    if t1[-1] > t2[-1]:
        not_norm.append(k)

    print('第' + str(k) + '回: 统计量D = ' + str(t1[-1]) + ', pvalue = ' + str(t2[-1]))
print('抽样不来自正态总体的章节是:', not_norm, '章节数为:', len(not_norm))
print('---------------------------------------------------------------------\n')

plt.figure()
plt.bar(range(1, 81), t2)
plt.bar(range(1, 81), np.array(t1) * (-1))
plt.show()
del t1, t2, t, k

# 分析: 使用单因素方差分析对于前80回，回数（80个水平）对高频字的使用有无显著影响
t = data1.copy()
if not_norm != []:
    for k in not_norm:
        del t['第' + str(k) + '回']
table = my.anova1(t, alpha=0.05)
print('前80回对高频字使用密度的单因素方差分析')
print(table)
print('---------------------------------------------------------------------\n')
del table, not_norm, t
if 'k' in dir():
    del k

# 分析: 按顺序每5回作一个单因素方差分析, 看看回数对高频字的使用有无显著影响
r = []
for k in range(16):
    t = {}
    for s in range(1, 6):
        t['第' + str(5 * k + s) + '回'] = data1['第' + str(5 * k + s) + '回']
    table = my.anova1(t, alpha=0.05)
    print('第' + str(5 * k + 1) + '到' + str(5 * k + 5) + '回对高频字使用密度的单因素方差分析')
    print(table)
    r.append(table._rows[0][4])
    print('---------------------------------------------------------------------\n')
plt.figure()
plt.bar(range(1, len(r) + 1), r)
plt.plot([1, len(r) + 1], [table._rows[0][5], table._rows[0][5]], color='red', linewidth=1)
plt.show()
del k, t, s, table, r

# 数据提取: 计算120回一每回中高频使用字的频率
hot_words_120_fre = {}
for k in range(1, 121):
    hot_words_120_fre['第' + str(k) + '回'] = 0
    for word in hlm['第' + str(k) + '回']:
        if word in hot_words_before80_picked:
            hot_words_120_fre['第' + str(k) + '回'] += 1
    hot_words_120_fre['第' + str(k) + '回'] /= word_symbol_count['第' + str(k) + '回']
del k, word

plt.figure()
plt.bar(range(1, 121), hot_words_120_fre.values())
plt.show()

# 数据提取: 将120回分成三大块: 1-40回, 41-80回, 81-120回, 计算每一大块下的每一回的高频字频率作为样本观察值
data2 = {'1_40': [], '41_80': [], '81_120': []}
for k in range(1, 41):
    data2['1_40'].append(hot_words_120_fre['第' + str(k) + '回'])
    data2['41_80'].append(hot_words_120_fre['第' + str(k + 40) + '回'])
    data2['81_120'].append(hot_words_120_fre['第' + str(k + 80) + '回'])
del k

# 分析: 对data2中每一个水平对应的抽样来自的总体作K-S正态性检验, 然后可视化
not_norm = []
t1 = []
t2 = []
for k in range(3):
    t = stats.kstest(data2[str(40 * k + 1) + '_' + str(40 * k + 1 + 39)], 'norm',
                     args=(
                         np.mean(data2[str(40 * k + 1) + '_' + str(40 * k + 1 + 39)]),
                         np.std(data2[str(40 * k + 1) + '_' + str(40 * k + 1 + 39)], ddof=1)
                     ))
    t1.append(t[0])
    t2.append(t[1])
    if t1[-1] > t2[-1]:
        not_norm.append(k)
print('抽样不来自正态总体的水平是:', not_norm, '个数为:', len(not_norm))
print('---------------------------------------------------------------------\n')

plt.figure()
plt.bar(range(1, 4), t2)
plt.bar(range(1, 4), np.array(t1) * (-1))
plt.show()
del t1, t2, t, k, not_norm

# 分析: 使用单因素分析四个水平对: 1_40与41_80, 1_40与81_120, 41_80与81_120, 1_80与81_120
# 目的: 从高频字频率角度判断后40回(即81-120回)是否还是曹雪芹所著
print('通过1-40回与41-80回的单因素双水平方差分析验证前80回是曹雪芹所著')
t = data2.copy()
del t['81_120']
table = my.anova1(t, alpha=0.05)
print(table)
print('---------------------------------------------------------------------\n')

print('通过1-40回与81-120回的单因素双水平方差分析验证后40回非曹雪芹所著')
t = data2.copy()
del t['41_80']
table = my.anova1(t, alpha=0.05)
print(table)
r = my.effect_scale(t)
print('效应尺度 = ' + str(r))
print('---------------------------------------------------------------------\n')
del t, table, r

print('通过41-80回与81-120回的单因素双水平方差分析验证后40回非曹雪芹所著')
t = data2.copy()
del t['1_40']
table = my.anova1(t, alpha=0.05)
print(table)
r = my.effect_scale(t)
print('效应尺度 = ' + str(r))
print('---------------------------------------------------------------------\n')
del t, table, r

print('通过1-80回与81-120回的单因素双水平方差分析验证后40回非曹雪芹所著')
t = data2.copy()
t['1_80'] = t['1_40'] + t['41_80']
del t['1_40'], t['41_80']
table = my.anova1(t, alpha=0.05)
print(table)
r = my.effect_scale(t)
print('效应尺度 = ' + str(r))
print('---------------------------------------------------------------------\n')
del t, table, r

# 分析: 使用秩和检验分析四个总体对: 1_40与41_80, 1_40与81_120, 41_80与81_120, 1_80与81_120
# 目的: 从高频字频率角度判断后40回(即81-120回)是否还是曹雪芹所著
print('通过1-40回与41-80回的秩和检验验证前80回是曹雪芹所著')
t = data2.copy()
del t['81_120']
my.kruskal_wallis_bilateral_test_over10(t, alpha=0.05)
print('---------------------------------------------------------------------\n')

print('通过1-40回与81-120回的秩和检验验证后80回是曹雪芹所著')
t = data2.copy()
del t['41_80']
my.kruskal_wallis_bilateral_test_over10(t, alpha=0.05)
print('---------------------------------------------------------------------\n')

print('通过41-80回与81-120回的秩和检验验证后80回是曹雪芹所著')
t = data2.copy()
del t['1_40']
my.kruskal_wallis_bilateral_test_over10(t, alpha=0.05)
print('---------------------------------------------------------------------\n')

print('通过1-80回与81-120回的秩和检验验证后80回是曹雪芹所著')
t = data2.copy()
t['1_80'] = t['1_40'] + t['41_80']
del t['1_40'], t['41_80']
my.kruskal_wallis_bilateral_test_over10(t, alpha=0.05)
print('---------------------------------------------------------------------\n')

del t

# 整理: 分别统计1-40回, 41-80回, 81-120回中句长为1,2,...,20的句子个数, 并且可视化
sentence_len_1_40 = {}
sentence_len_41_80 = {}
sentence_len_81_120 = {}
for k in range(1, 21):
    sentence_len_1_40[k] = 0
    sentence_len_41_80[k] = 0
    sentence_len_81_120[k] = 0

    for chapter in range(1, 41):
        if k in sentence_len_120['第' + str(chapter) + '回'].keys():
            sentence_len_1_40[k] += sentence_len_120['第' + str(chapter) + '回'][k]
        if k in sentence_len_120['第' + str(chapter + 40) + '回'].keys():
            sentence_len_41_80[k] += sentence_len_120['第' + str(chapter + 40) + '回'][k]
        if k in sentence_len_120['第' + str(chapter + 80) + '回'].keys():
            sentence_len_81_120[k] += sentence_len_120['第' + str(chapter + 80) + '回'][k]

del k, chapter

plt.figure()
plt.bar(
    np.array(range(1, 21)),
    np.array(list(sentence_len_1_40.values()))
    + np.array(list(sentence_len_41_80.values()))
    + np.array(list(sentence_len_81_120.values())),
)
plt.show()

plt.figure()
plt.plot(np.array(range(1, 21)), sentence_len_1_40.values(), linewidth=1)
plt.plot(np.array(range(1, 21)), sentence_len_41_80.values(), linewidth=1)
plt.plot(np.array(range(1, 21)), sentence_len_81_120.values(), linewidth=1)
plt.show()

# 分析: 在前80回中使用多元线性回归分析句长在各回的规律
x = []
for chapter in range(1, 81):
    t = []
    for k in range(4, 17):
        t.append(sentence_len_120['第' + str(chapter) + '回'][k])
    x.append(t)
X = sm.add_constant(x)
model = sm.OLS(list(word_symbol_count.values())[0:80], X).fit()
print(model.summary())
del x, k, t, chapter, X, model

# # 清除所有的变量
# for var in dir():
#     if not var.startswith('__'):
#         globals().pop(var)
# del var
