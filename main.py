from openpyxl import load_workbook
import pyperclip
FILEPATH = './data/lidar_20250604_181314.xlsx'
TXT_PATH = 'last_row.txt'
def road_txt():
    try:
        with open(TXT_PATH, 'r') as f:
            return int(f.read().strip())
    except(FileNotFoundError, ValueError):
        return 1
def save_txt(num):
    with open(TXT_PATH, 'w') as f:
        f.write(str(num))

def main():
    wb = load_workbook(filename=FILEPATH)
    ws = wb.active
    n = 2
    val = ws[f'A{n}'].value
    pyperclip.copy(val)
    print(f'A{n}:{val}をコピー')
if __name__ == '__main__':
    main()
