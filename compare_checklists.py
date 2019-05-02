import re
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from textblob import TextBlob
import xlwt
import xlrd
from xlutils.copy import copy
from xlrd import open_workbook
from fuzzywuzzy import fuzz


def xls_compare(source_obj, compare_obj):
    OUTPUT_FILE = "output_file.xls"
    original_count =0 

    #Create a new sheet to add the back annotated checklist items to

    #Make a copy of the source obj
    output_obj = copy(source_obj)
    sh = output_obj.add_sheet("ANNOTATED_ITEMS")

    col1_name = 'NEW_ID'
    col2_name = 'NEW_DESCRIPTION'
    col3_name = 'NEW_PARAGRAPH'
    col4_name = 'NEW_CATEGORY'
    col5_name = 'ORIGINAL_ID'
    col6_name = 'ORIGINAL_DESCRIPTION'
    col7_name = 'ORIGINAL_PARAGRAPH'
    col8_name = 'ORIGINAL_CATEGORY'

    sh.write(0, 1, col1_name)
    sh.write(0, 2, col2_name)
    sh.write(0, 3, col3_name)
    sh.write(0, 4, col4_name)
    sh.write(0, 5, col5_name)
    sh.write(0, 6, col6_name)
    sh.write(0, 7, col7_name)
    sh.write(0, 8, col8_name)

    #1. Go through the source file and copy all items to destination sheet, offset by 4 columns
    src_sh = source_obj.sheet_by_index(0)
    cmp_sh = compare_obj.sheet_by_index(0)
    

    for i,cell in enumerate(cmp_sh.col(1)):
        found_exact = 0
        found_ratio = 0
        if not i:
            continue
        new_string = cmp_sh.cell(i,1).value
        new_id = cmp_sh.cell(i,0).value
        #Check if there is an exact match 
        for k,cell in enumerate(src_sh.col(1)):
            if not k:
                continue
            old_string = src_sh.cell(k,1).value
            old_id = src_sh.cell(k,0).value

            #First check if there is an exact match
            if new_string == old_string :
                found_exact = 1
                print("EXACT_MATCH", new_string, new_id, old_id)
                continue
            #Otherwise calculate the match ratio using FuzzyWuzzy package NOTE: use small caps to get more accurate ratio
            ratio = fuzz.ratio(new_string.lower(), old_string.lower() )
            token_ratio = fuzz.token_sort_ratio(new_string.lower(), old_string.lower())
            if ratio > 75 :
                found_ratio = 1
                print("RATIO_MATCH", new_string, new_id, old_string, old_id)
                continue
            elif token_ratio > 85 :
                found_ratio = 1
                print("TOKEN_RATIO_MATCH", new_string, new_id, old_string, old_id)
                continue
        if found_exact == 0 and found_ratio == 0:
            print("NEW_ASSERTION", new_string, new_id)

    #2. Go through each item in the compare file and find match in original file
    #3. If a match is found, annotate the item in the spreadsheet
    #4. If no match is found, add item at bottom of the list, not NO_MATCH in ORIGINAL_* columns
    #5. After done, check items which have contents in ORIGINAL_* columns, but left empty in NEW_* columns, add "REMOVED" in NEW_* columns

    output_obj.save(OUTPUT_FILE)



SOURCE_FILE = "source_file.xls"
COMPARE_FILE = "compare_file.xls"

ix_num = 0
source_book = open_workbook(SOURCE_FILE, formatting_info=True)
compare_book = open_workbook(COMPARE_FILE, formatting_info=True)

xls_compare(source_book, compare_book)

