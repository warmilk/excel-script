import re
import pandas
import os

'''
 python version = 3.9.6
 Make sure to close '/assets/compare.xlsx' when the script is running !!!
'''

# Get the path separator for the current system
os_sep = os.path.sep


match_rules_set = set()
not_match_rules_set = set()

def go():
    excel_file_path = '.' + os_sep + 'assets' + os_sep + 'compare.xlsx'
    sids_np_array_dataframe = pandas.read_excel(excel_file_path, sheet_name='Sheet1')
    rules_np_array_dataframe = pandas.read_excel(excel_file_path, index_col=0, sheet_name='Sheet2')
    sids_np_array = sids_np_array_dataframe.values
    rules_np_array = rules_np_array_dataframe.values
    print('passing rate :', len(sids_np_array)/len(rules_np_array))
    whole_rules_set = set()
    for sid in sids_np_array:
        for rule in rules_np_array:
            num = str(sid[0])
            regexp = re.compile(rf'sid\s*?:\s*?{num}\s*?;')
            if regexp.search(rule[0]):
                match_rules_set.add(rule[0])
            else:
                whole_rules_set.add(rule[0])
    not_match_rules_set = whole_rules_set.difference(match_rules_set)
    print('match :', len(match_rules_set))
    print('not match :', len(not_match_rules_set))
    print('all :', len(whole_rules_set))

    with pandas.ExcelWriter(excel_file_path, mode='a', if_sheet_exists='replace') as writer:
        pandas.DataFrame(list(match_rules_set)).to_excel(writer, sheet_name='match')
        pandas.DataFrame(list(not_match_rules_set)).to_excel(writer, sheet_name='not_match')


if __name__ == '__main__':
    go()

