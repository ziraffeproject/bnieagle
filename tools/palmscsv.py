import sys
import csv
import re
import xml.etree.ElementTree as ET
import os
import os.path
import glob


def replace_cell_text(text):
    if text is None:
        return None
    
    return re.sub(r'(\d+)-(\d+)-(\d+)T00:00:00', r'\1/\2/\3', text)

    

def convert_file(filename, output_filename):
    tree = ET.parse(filename)
    rows = tree.findall('.//{urn:schemas-microsoft-com:office:spreadsheet}Row')

    with open(output_filename, 'w', encoding='utf-8') as csvfile:
        csv_writer = csv.writer(csvfile)

        for row in rows:
            if row[0][0].text in ['BNI', 'ビジター', '合計']:
                continue

            csv_writer.writerow(map(replace_cell_text, [cell[0].text for cell in row]))

    print("{}に出力しました。".format(output_filename))


def main():
    dirname = os.path.dirname(sys.argv[0])
    if dirname != '':
        os.chdir(dirname)

    if len(sys.argv) == 1:
        filenames = glob.glob('_________PALMS_*.xls')
        if len(filenames) != 1:
            print("処理するファイルがないか、複数あるため、処理を中断しました。")
            return

        convert_file(filenames[0], 'palms.csv')

    else:
        for filename in sys.argv[1:]:
            base_filename, file_extension = os.path.splitext(filename)
            convert_file(filename, '{}.csv'.format(base_filename))
            


main()
print("Enterキーを押してください")
input()




