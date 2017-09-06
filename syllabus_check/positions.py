# encoding: utf-8

#
# 科目データ
#
kamoku_info = {
    'year_p': (10, 5),
    'curr_p': (11, 5),
    'sem_p': (12, 5),
    'cred_p': (13, 5),
    'code_p': (14, 5),
    'name_p': (15, 5),
    'teach_p': (16, 5)
}

##
# Item maps for sheet traversal
# 必須項目の配置
##

# 30コマ科目
pos_30 = {  # row count 183
    '到達目標': (19, 5),
    '授業概要': (26, 5),
    '授業外学修': (155, 5),   # this is for 15x classes
    '評価方法': (162, 5),
    '教科書等': (169, 5)
}

# 15コマ科目
pos_15 = {  # row count 123
    '到達目標': (19, 5),
    '授業概要': (26, 5),
    '授業外学修': (95, 5),   # this is for 15x classes
    '評価方法': (102, 5),
    '教科書等': (109, 5)
}

# 8コマ科目
pos_8 = {  # row count 95
    '到達目標': (19, 5),
    '授業概要': (26, 5),
    '授業外学修': (67, 5),   # this is for 8x classes
    '評価方法': (74, 5),
    '教科書等': (81, 5)
}


##
# Item maps for sheet traversal: plan
# 受業計画：配置
##
plan_30 = [(i, 8) for i in range(34, 151, 4)]
plan_15 = [(i, 8) for i in range(34, 91, 4)]
plan_8 = [(i, 8) for i in range(34, 63, 4)]
