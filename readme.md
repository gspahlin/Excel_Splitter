## Excel Splitter
**Intial Startup**

You need an Anaconda distribution for you computer. Get it here: https://www.anaconda.com/download

Install and open the anaconda prompt. 

**Installing the Virtual Environment**

Type the following commands and run them in order:

```shell
$conda create -n splitter
$conda activate splitter
$pip install numpy pandas pyarrow
$pip install xlwings==0.23.0

```
After setting this up once you can enter and exit with 

```shell
$conda activate splitter
$conda deactivate
```
Note on xlwings - This utility uses xlwings (a Python API for MS Excel) to extract data from bloated excel files. 
I have not run xlwings on a computer that has a version of MS Excel with python enabled, but I have heard that this causes conflicts. 
If you have python enabled in your excel troubleshooting may be required to recover compatibility. 

**Basic Function**

Excel_splitter is designed to extract data from Microsoft Excel files that are overly large. In cases where an Excel (.xlsx) sheet has over 1 million rows, they will be split into multiple sheets. 
Additionally, Excel is not well optimized for applications related to "big data". Excel files that go above 1 GB are generally very cumbersome to manipulate in excel, and are a pain to convert to 
other formats. This does not, however, mean that you will never encounter overly large excel files or smaller files that you would prefer to convert into a different format.

Excel splitter will scan the sheets present in an excel file and look for sheets with identical fields. It will pull them together and then write out the full sheets  as .csv or parquet files 
(parquets are default). These extracted files will be written into a folder in the same folder as the original excel file. Excel splitter is also compatable with xlsx files that are password 
protected. 


**Excel Splitter Run Command**

You can run excel splitter using the following command:

```shell
$python excel_splitter.py -x "<path to xlsx file>" 
```
for parquet outputs 

```shell
$python excel_splitter.py -x "<path to xlsx file>" -p "<password>"
```
for parquet outputs if the xlsx is protected with a password

```shell
$python excel_splitter.py -x "<path to xlsx file>"  -c True
```
If you would like outputs in the form of csv files. A -p "<password>" term
can be added to this if you want csv outputs and there is a password on the original xlsx.

note: double quotes may be used for the path and the password if you choose. If there are any spaces in the path 
double quotes MUST be used.