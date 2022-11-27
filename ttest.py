

from difflib import SequenceMatcher, Differ
import openpyxl as xl
from openpyxl.styles import Color, PatternFill, Font
import xlsxwriter
import os

# https://xlsxwriter.readthedocs.io/example_rich_strings.html
# import difflib

class KoreanDiff:
    def __init__(self, source_path):
        self.source_excel_path = source_path
        temp_target_path_list = source_path.split("\\")
        temp_target_path_list[-1] = "compared_" + temp_target_path_list[-1]
        self.target_excel_path = "\\".join(temp_target_path_list)

        self.source_row_num = 0
        self.target_row_num = 2

        self.red_color = Font(color='FE2E2E')
        self.black_color = Font(color='000000')

        self.output_workbook = xl.Workbook()
        self.output_sheet = self.output_workbook.active
        # write_ws.append([1,2,3])

        print(f"source path: {self.source_excel_path}")
        print(f"target path: {self.target_excel_path}")
        return

    def run(self):
        # read source excel
        self.wb = xl.load_workbook(self.source_excel_path)
        self.first_sheetname = self.wb.sheetnames[0]
        # do  get matching blocks
        for row in self.wb[self.first_sheetname].iter_rows(min_row = 1):
        # for each line
            print(row[self.source_row_num].value)
            print(row[self.target_row_num].value)
            
            # fill red all
            row[self.source_row_num].fill = self.red_color
            row[self.target_row_num].fill = self.red_color

            if not row[self.source_row_num].value or not row[self.target_row_num].value:
                continue
            # same match -> black
            s = SequenceMatcher(None, row[self.source_row_num].value, row[self.target_row_num].value)
            # print(s.get_matching_blocks())
            for mb in s.get_matching_blocks():
                print(mb, row[self.source_row_num].value[mb.a:mb.a+mb.size])
                print(mb, row[self.target_row_num].value[mb.b:mb.b+mb.size])
            print("*"*30)
        return 


if __name__ == "__main__":
    # dir_path = "./sample/"
    source_path = "D:\\GitKoo\\pyKoreanDiff\\sample\\test.xlsx"
    # with open(source_path, 'rt', encoding='UTF8') as f:
    #     source = f.read()
    # with open(target_path, 'rt', encoding='UTF8') as f:
    #     target = f.read()
    # __run_diff(source, target)
    
    k_diff = KoreanDiff(source_path)
    k_diff.run()
    pass