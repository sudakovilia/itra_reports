import pandas as pd

if __name__ == '__main__':
    source_1c_data = pd.read_excel('./data/1c_data.xlsx')
    print(len(source_1c_data.index))