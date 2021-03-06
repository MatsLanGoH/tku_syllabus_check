# encoding: utf-8

# 点検対象となるWorkbookのセル配置
# CampusMateの出力形式がもし変われば、
# ここで実際の位置を反映させなければならない。

#
# 科目データ
#
kamoku_info = {
    'year': (10, 5),        # 年度
    'curriculum': (11, 5),  # 開講学科・専攻
    'semester': (12, 5),    # 講義期間
    'credits': (13, 5),     # 単位数
    'code': (14, 5),        # 講義コード
    'name': (15, 5),        # 授業科目名
    'teacher': (16, 5)      # 授業担当者氏名
}


# Item maps for sheet traversal
# 必須項目の配置
# 科目のコマ数によって異なるため、Excelで確認して以下のdictに定義づける。
kamoku_maps = dict()

# 30コマ科目
items_30 = {  # row count 183
    '到達目標': (19, 5),  # '到達目標'
    '授業概要': (26, 5),  # '授業概要'
    '授業外学修': (155, 5),  # '授業外学修'
    '評価方法': (162, 5),  # '評価方法'
    '教科書等': (169, 5),  # '教科書等'
    '受業計画': [(i, 8) for i in range(34, 151, 4)]  # '受業計画'
}


# 15コマ科目
items_15 = {  # row count 123
    '到達目標': (19, 5),  # '到達目標'
    '授業概要': (26, 5),  # '授業概要'
    '授業外学修': (95, 5),   # '授業外学修'
    '評価方法': (102, 5),  # '評価方法'
    '教科書等': (109, 5),  # '教科書等'
    '受業計画': [(i, 8) for i in range(34, 91, 4)]  # '受業計画'
}

# 8コマ科目
items_8 = {  # row count 95
    '到達目標': (19, 5),  # '到達目標'
    '授業概要': (26, 5),  # '授業概要'
    '授業外学修': (67, 5),   # '授業外学修'
    '評価方法': (74, 5),  # '評価方法'
    '教科書等': (81, 5),  # '教科書等'
    '受業計画': [(i, 8) for i in range(34, 63, 4)]  # '受業計画'
}


# Add item maps to dict
# シートの行数から引き出せるようにキーは行数に合わせる。
# Ex: 30コマの授業はsheet.nrowsが183⇒
kamoku_maps[183] = items_30
kamoku_maps[123] = items_15
kamoku_maps[95] = items_8
