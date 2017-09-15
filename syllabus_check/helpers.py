# encoding: utf-8

from difflib import SequenceMatcher


##
# HELPER FUNCTIONS
#
def is_empty(s):
    """
    s : string with description of item
    returns True if string is too short
    else False
    """
    return len(s) <= 2


def has_test_in_lesson(s):
    """
    s : string with description of lesson
    returns True if forbidden word is contained in s
    else False
    """
    forbidden = ["試験", "テスト", "ﾃｽﾄ", "exam", "test"]
    for f in forbidden:
        if f in s.lower():
            return True
    return False


def has_illegal_grading_method(s):
    """
    s : string with description of grading method
    returns True if forbidden word contained in s
    else False
    """
    forbidden = ["出席", "出欠", "欠席", "参加度"]
    for f in forbidden:
        if f in s:
            return True
    return False


def contains_similar_items(items, ratio=0.5):
    """
    items: a list of strings (of lesson descriptions)
    returns: a list of similar strings,
             or an empty list if no similar strings are found.
    """
    similar_strings = []

    # Iterate over array and check if an item is similar to following strings
    for idx, item in enumerate(items):
        for partner in items[idx+1:]:
            s = SequenceMatcher(None, item, partner)

            # If similar, add compared strings into similar_strings array
            if s.quick_ratio() > ratio:
                # print('{} : {} / {}'.format(item, partner, idx))
                # Ensure this string is in the list as well!
                if item not in similar_strings:
                    similar_strings.append(item)
                if partner not in similar_strings:
                    similar_strings.append(partner)

    return similar_strings

