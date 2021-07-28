import sys
import os

sys.path.insert(1, "../")


from data2doc.data2doc import merge

test_result_file_name = "test-result/test-result.docx"
test_data_folder = "test-data/test01/"

if not os.path.exists("test-data"):
    test_result_file_name = f"../{test_result_file_name}"
    test_data_folder = f"../{test_data_folder}"

test_input_file_name1 = f"{test_data_folder}test.docx"
test_input_file_name2 = f"{test_data_folder}test.xlsx"


result = merge(test_input_file_name1,
               test_input_file_name2,
               test_result_file_name)

if result and "win" in sys.platform:
    os.startfile(test_result_file_name.replace("/", "\\"))
