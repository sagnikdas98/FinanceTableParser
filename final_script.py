import ctypes
import csv
 
try:
    import pdfplumber
except:
    ctypes.windll.user32.MessageBoxW(0, "Please install pdf install using 'pip install pdfplumber' in cmd",
                                     "pdfplumber not installed!", 1)
    quit(0)
 
def pdftoTable(in_pdfName = None, in_pageNo = 0,symbol = "$"):
 
    if in_pdfName == None:
        ctypes.windll.user32.MessageBoxW(0, "Please Enter the pdf address",
                                         "Incorrect Address!", 1)
 
 
    lines = []
    temp_str = ''
 
 
    try:
        with pdfplumber.open(in_pdfName) as pdf:
            page = pdf.pages[in_pageNo]
            for char in page.extract_text():
                if char == '\n':
                    lines.append(temp_str)
                    temp_str = ''
                    continue
                temp_str += char
 
 
    except(IndexError,IOError) as e :
            ctypes.windll.user32.MessageBoxW(0, e, "Error!", 1)
            quit(0)
 
    frequency_dollar_dict = {}
 
    try:
 
        for i in range(1, 6):
            frequency_dollar_dict[i] = []
    except:
 
        ctypes.windll.user32.MessageBoxW(0, "More than 6 coloums",
                                         "Too many columns", 1)
        quit(0)
 
 
    for li in lines:
        k = li.count(symbol)
        if k == 0:
            continue
        frequency_dollar_dict[li.count(symbol)].append(li)
 
    maxLen = 0
    index = 0
 
 
 
    for i in range(1, 6):
        if len(frequency_dollar_dict[i]) > maxLen:
            maxLen, index = len(frequency_dollar_dict), i
 
    table_data = frequency_dollar_dict[index]
 
    json_list = {}
    dollar_index = 0
    int_list = []
    for line in table_data:
        dollar_index = line.index(symbol)
        key_string = line[:dollar_index]
        rem_string = line[dollar_index:].split(symbol)[1:]
        for num in rem_string:
            int_list.append(num.strip())
        json_list[key_string] = int_list
        int_list = []
 
    print(json_list)
 
    with open("test.csv", "w") as outfile:
        writer = csv.writer(outfile)
        writer.writerow(json_list.keys())
        writer.writerows(zip(*json_list.values()))
        
        
pdftoTable("NASDAQ_AMZN_2017.pdf", 28)