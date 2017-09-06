# encoding: utf-8

from difflib import SequenceMatcher


def find_similar_items(items):
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
            if s.quick_ration() > 0.5:
                print('{} : {} / {}'.format(item, partner, idx))
                # Ensure this string is in the list as well!
                if item not in similar_strings:
                    similar_strings.append(item)
                similar_strings.append(partner)

    return similar_strings

