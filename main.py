import timeit
import os
import xlrd as xl
import pandas as pd
from openpyxl import load_workbook

def count():
    lines_for_each_file = []

    dir = r'C:/Users/Johnn/OneDrive/Desktop/Brain Corp/BJs/EOP Report/Exception Reports'

    files = [f for f in os.listdir(dir) if os.path.isfile(f) and f != 'main.py' and f!= ".gitignore"]

    count = 0
    for file in files:
        df = pd.read_excel(file)

        print(len(df.index))
        lines_for_each_file.append(len(df.index))
    
    print(sum(lines_for_each_file))

def main():
    if __name__ == "__main__":
        count()
        print(timeit.timeit('output = 10*5'))

main()