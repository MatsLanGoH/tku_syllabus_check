# encoding: utf-8

from helpers import * 

##
# 点検項目を入れる行列
#
checks = []


################################################
# Check 1: 必須科目はすべて記入されている
################################################

# (1) 点検内容の説明（指示）
message = "必須科目をすべて記入してください。"


# (2) 点検条件の定義
def all_items_entered(items, sheet):
    error_msg = ""

    for label, description in items.items():
        try:
            cell = sheet.cell_value(*description)
            if is_empty(cell):
                error_msg += '\t({})\n'.format(label)
        except TypeError:
            # print("Invalid value: {} : {}".format(label, description))
            pass
    return error_msg


# (3) 点検行列に加えておく
checks.append([message, all_items_entered])


################################################
# Check : 受業計画は開講回数すべてを記載してください。
################################################

# (1) 点検内容の説明（指示）
message = "受業計画は開講回数すべてを記載してください。"


# (2) 点検条件の定義
def all_lessons_entered(items, sheet):
    error_msg = ""
    count = 1
    try:
        plan = items['受業計画']
        for lesson in plan:
            koma = sheet.cell_value(*lesson)
            if is_empty(koma):
                error_msg += '\t(第{}回)\n'.format(count)
            count += 1
    except KeyError:
        print("No lesson plan found")
        pass

    return error_msg


# (3) 点検行列に加えておく
checks.append([message, all_lessons_entered])


################################################
# Check : 受業計画は各回の違いが明確になるように記載してください。
################################################

# (1) 点検内容の説明（指示）
message = "受業計画は各回の違いが明確になるように記載してください。"


# (2) 点検条件の定義
def test_crit(items, sheet):
    error_msg = ""

    try:
        plan = items['受業計画']
        lessons = []
        for lesson in plan:
            koma = sheet.cell_value(*lesson)

            # Remove unwanted characters
            koma = u''.join(koma.strip().split())

            # Ignore empty fields
            if len(koma) > 0:
                lessons.append(koma)

        # Check similarity for lessons.
        # Ratioを低くすればするほど、引っかかってしまうコマが増える
        # 文字数も関係するため、短い説明文は特に怪しい。
        # 自動項目として入れるなら微調整が必要
        similars = contains_similar_items(lessons, 0.8)

        # Gather list of similar items.
        if len(similars) > 0:
            error_msg += '\t(以下のコマを確認してください。)'
            for similar in similars:
                error_msg += '\n\t{}'.format(similar)
            error_msg += '\n'

    except KeyError:
        print("No lesson plan found")
        pass

    return error_msg


# (3) 点検行列に加えておく
checks.append([message, test_crit])


################################################
# Check : 受業計画に「試験」「テスト」の表記を使わないでください。
################################################

# (1) 点検内容の説明（指示）
message = "受業計画に「試験」「テスト」の表記を使わないでください。"


# (2) 点検条件の定義
def test_crit(items, sheet):
    error_msg = ""

    count = 1
    try:
        plan = items['受業計画']
        for lesson in plan:
            koma = sheet.cell_value(*lesson)
            if has_test_in_lesson(koma):
                error_msg += '\t(第{}回)\t{}\n'.format(count, koma)
            count += 1
    except KeyError:
        print("No lesson plan found")
        pass

    return error_msg


# (3) 点検行列に加えておく
checks.append([message, test_crit])


################################################
# Check : 評価方法に出席点に関わる表記を使わないでください。 
################################################

# (1) 点検内容の説明（指示）
message = "評価方法に出席点に関わる表記を使わないでください。 "


# (2) 点検条件の定義
def test_crit(items, sheet):
    error_msg = ""
    
    try:
        grade_pos = items['評価方法']
        grade_method = sheet.cell_value(*grade_pos)
        if has_illegal_grading_method(grade_method):
            error_msg += '\t{}\n'.format(grade_method)
    except KeyError:
        print("Item not found")
        pass

    return error_msg


# (3) 点検行列に加えておく
checks.append([message, test_crit])


################################################
# Check :
################################################

# (1) 点検内容の説明（指示）
message = ""


# (2) 点検条件の定義
def test_crit(items, sheet):
    error_msg = ""

    return error_msg


# (3) 点検行列に加えておく
checks.append([message, test_crit])



