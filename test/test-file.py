import sys
import os

sys.path.insert(1, "../")

from data2doc.data2doc import merge

test_data_folder="../test-data/test01/"
test_inpot_file_name1=f"{test_data_folder}test.docx"
test_inpot_file_name2=f"{test_data_folder}test.xlsx"
test_result_file_name="../test-result/test-result.docx"

result=merge(test_inpot_file_name1,test_inpot_file_name2,test_result_file_name)

if result and "win" in sys.platform:
    os.startfile(test_result_file_name.replace("/","\\"))