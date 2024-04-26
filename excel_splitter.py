from bz2 import compress
from tkinter import N
import pandas as pd
import numpy as np
import xlwings as xw
import argparse as arg
import os
import shutil as sht
#use the data environment
#requires pyarrow package for parquets

class Excel_bloater:
    def __init__(self, path, password, csv):
        self.path = path
        self.password = password
        self.csv = csv
        self.exclusion_list  = ['Query']
        #everything up to the file name in the xlsx path
        self.save_root = os.path.split(path)[0]
        #the name with extension
        self.file_name = os.path.split(path)[1]
        #the name without extension
        self.file_stem = self.file_name.split('.')[0]
        self.sheet_names = []
        self.sheet_dfs = []

    def extract_sheet_dfs(self):
        app = xw.App(visible=False)
        wb = app.books.open(self.path, password=self.password)

        for s in wb.sheets:
            exclude = False
            for ex in self.exclusion_list:
                if ex in s.name:
                    exclude=True
                    print(f'{s.name} excluded for {ex} substring')
            if exclude==False:
                self.sheet_names.append(s.name)

        for sheet in self.sheet_names:
            current_sheet = wb.sheets[sheet]
            new_df = current_sheet.range('A1').expand().options(pd.DataFrame, index=False, header=True).value
            key_names=[]
            for key, value in new_df.items():
                key_names.append(key)
            self.sheet_dfs.append({'name': sheet,                                    
                                   'keys':key_names,
                                   'df': new_df})

            print(f'{sheet} loaded into memory')
        wb.close()
        app.kill()
      
        print(f'all data extracted from {self.file_name}')

    def merge_dfs(self):

        for n, dickt in enumerate(self.sheet_dfs):
            print(f"checking sheet {dickt['name']} for matches")
            match_indicies=[]
            start = n+1
            finish = len(self.sheet_dfs)-1
            for r in range(start, finish):
                if dickt['keys'] == self.sheet_dfs[r]['keys']:
                    match_indicies.append(r)
                    print(f"full column match found between {dickt['name']} and {self.sheet_dfs[r]['name']}, preparing to merge")
            if len(match_indicies)>0:
                df_container=[dickt['df']]
                for match in match_indicies:
                    df_container.append(self.sheet_dfs[match]['df'])
                dickt['df']=pd.concat(df_container, ignore_index=True)
                print(f"Dataframes merged into {dickt['name']}")
                shift=0
                for match in match_indicies:
                    print(f"popping {self.sheet_dfs[match-shift]['name']} to clean up redundancy")
                    self.sheet_dfs.pop(match-shift)
                    shift+=1
            else:
                print(f"no matches found for {dickt['name']}")                                     

    def write_sheet_dfs(self):
        print(f'csv status: {self.csv}')
        #write a folder for data
        folder_path=os.path.join(self.save_root, f'{self.file_stem}')
        print(f'Writing folder {self.file_stem} in parent directory {self.save_root}')        
        os.mkdir(folder_path)
        print(f'data folder {folder_path} written')
        for n, dickt in enumerate(self.sheet_dfs):
            #derive save_name
            #join the newly made folder to the file name
            #the bug is down here when we write things out
            #the bug was the character limit being hit
            save_path = f'''{folder_path}\\{self.file_stem}_{dickt['name']}.parquet.gzip'''
            print(f'preparing to save data at {save_path}')          
            #write each into its own data file
            if self.csv=='False':
                dickt['df'].to_parquet(save_path, compression='gzip')
                print(f'{save_path} written as parquet')
            elif self.csv=='True':
                csv_path = save_path.replace('parquet.gzip', 'csv')
                dickt['df'].to_csv(csv_path, index=False)
                print(f'{csv_path} written as csv')


def main():
    parser = arg.ArgumentParser()
    parser.add_argument("-x", "--xlsx", help="bloated xlsx to split", required=True)
    parser.add_argument("-p", "--password", help="password for xlsx", required=False)
    parser.add_argument("-c", "--csv", help="csv output option, True or False. False = parquet output", required=False)
    args = parser.parse_args()
    
    datapath = args.xlsx
    passw0rd = args.password
    csv_mode = args.csv
    if (csv_mode != 'True') and (csv_mode != 'False'):
        csv_mode='False'
    bloater = Excel_bloater(datapath, passw0rd, csv_mode)
    bloater.extract_sheet_dfs()
    bloater.merge_dfs()
    bloater.write_sheet_dfs()
    print(f'bloater {bloater.file_name} split successfully')

if __name__=="__main__":
    main()