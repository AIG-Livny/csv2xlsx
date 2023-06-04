#!/usr/bin/env python3

import pandas
import os
import sys
import xlwt
import json
import re
import dataclasses
import io

@dataclasses.dataclass
class Table:
    name: str
    data: str

@dataclasses.dataclass
class TablesFile:
    tables: list[Table]

def replace_new_line(matched_string):
    if matched_string:
        return matched_string.group(1)+re.sub(r'\n', r'\\n', 
                matched_string.group(2))+matched_string.group(3)
    else:
        return matched_string

def custom_parser(multiline_string):
    if isinstance(multiline_string, (bytes, bytearray)):
        multiline_string = multiline_string.decode()
    multiline_string = re.sub(r'\t', r' ', multiline_string)
    return re.sub(r'(\s*")(.*?)((?<!\\)")', replace_new_line, 
    multiline_string, flags=re.DOTALL)

def print_help():
    scriptname = os.path.basename(__file__)
    print(f'{scriptname} - csv to xlsx or xls converter v.0.4')
    print(f'    usage: {scriptname} [csv or esv path] [xlsx or xls path]')
    print(f'    esv - JSON formatted CSV file. This allow write multiple sheets')



def convert(if_path:str, of_path:str):
    # Create directories for path
    dirs = os.path.dirname(of_path)
    if dirs != '':
        os.makedirs(dirs, exist_ok=True)

    # Input processing
    try:
        with open(if_path,'r') as f:
            if if_path.lower().endswith('esv'):
                d = json.loads(custom_parser(f.read()))
                file = TablesFile([Table(**t) for t in d['tables']])
            else:
                file = TablesFile([Table('sheet1', f.read())])
    except Exception as e:
            print(e)
            exit(1)
    

    # Output processing
    if of_path.lower().endswith('.xlsx'):
        output = io.BytesIO()
        writer = pandas.ExcelWriter(output, engine='openpyxl')
        for table in file.tables:
            df = pandas.read_csv(io.StringIO(table.data), comment=';', na_filter = False)
            df.to_excel(writer, sheet_name=table.name, index=False)
        writer.close()
        with open(of_path,'wb+') as f:
            f.write(output.getbuffer())
        exit(0)

    if of_path.lower().endswith('.xls'):
        workbook = xlwt.Workbook(of_path)
        for table in file.tables:
            df = pandas.read_csv(io.StringIO(table.data), comment=';', na_filter = False)
            worksheet = workbook.add_sheet(table.name,cell_overwrite_ok=True)
            number_format = xlwt.easyxf(num_format_str='0.00')
            cols_to_format = {0:number_format}

            for z, value in enumerate(df.columns):
                worksheet.write(0, z, value)

            # Iterate over the data and write it out row by row
            for x, y in df.iterrows():
                for z, value in enumerate(y):
                    if z in cols_to_format.keys():
                        worksheet.write(x + 1, z, value, cols_to_format[z])
                    else:
                         worksheet.write(x + 1, z, value)

        workbook.save(of_path)
        exit(0)

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('Wrong arguments number\n')
        print_help()
        exit(1)

    if not sys.argv[1][-4:].lower() in ['.csv','.esv']:
        print('First argument must be .csv or .esv file\n')
        print_help()
        exit(1)

    if not (sys.argv[2].lower().endswith('.xlsx') or sys.argv[2].lower().endswith('.xls')):
        print('Second argument must be .xlsx or .xls file\n')
        print_help()
        exit(1)

    if not os.path.isfile(sys.argv[1]):
        print(sys.argv[1]+' is not exists or not file')
        exit(1)

    convert(sys.argv[1], sys.argv[2])








