#!/usr/bin/env python
# coding: utf-8

# In[4]:


import xlrd
import xlwt
import os
import datetime

#read from database and store into a dictionary 
rb=xlrd.open_workbook("database.xls")
database=rb.sheet_by_index(0)
single_dictionary = {}
for i in range(database.nrows):
    key =database.cell(i, 0).value
    co =int(database.cell(i, 1).value)
    single_dictionary[key]=co
#os.remove("single_database.xls")

#rb=xlrd.open_workbook("double_database.xls")
database=rb.sheet_by_index(1)
double_dictionary = {}
for i in range(database.nrows):
    key =database.cell(i, 0).value
    co =int(database.cell(i, 1).value)
    double_dictionary[key]=co
#os.remove("double_database.xls")

#rb=xlrd.open_workbook("triple_database.xls")
database=rb.sheet_by_index(2)
triple_dictionary = {}
for i in range(database.nrows):
    key =database.cell(i, 0).value
    co =int(database.cell(i, 1).value)
    triple_dictionary[key]=co

# check already done news
database=rb.sheet_by_index(3)
already_done_dictionary = {}
for i in range(database.nrows):
    key =database.cell(i, 0).value
    co =int(database.cell(i, 1).value)
    already_done_dictionary[key]=co


# In[5]:


read=xlrd.open_workbook("Prothom_Alo.xls")
read_news=read.sheet_by_index(0)
import re
for i in range(read_news.nrows):
    news=read_news.cell(i,1).value
    check = read_news.cell(i,0).value
    #if news is already process then continue
    if check in already_done_dictionary:
        continue;
    already_done_dictionary[check]=1;
    #make three list with one, two and three words.
    list_with_single_word= re.split('\s|\)|\(|:|\।|\।|, ', news);
    list_with_single_word= list(filter(None, list_with_single_word))
    list_with_two_word=[]
    list_with_three_word=[]
    for j in range(1, len(list_with_single_word) ):
        list_with_two_word.append(list_with_single_word[j-1]+" "+list_with_single_word[j])
        
    for j in range(2, len(list_with_single_word) ):
        list_with_three_word.append(list_with_single_word[j-2]+" "+list_with_single_word[j-1]+" "+list_with_single_word[j])
        
    #update single dictionary  
    for wrd in (list_with_single_word):
        if wrd in single_dictionary:
            single_dictionary[wrd]=single_dictionary[wrd]+1;
        else:
            single_dictionary[wrd]=1;
            
    #update double dictionary  
    for wrd in (list_with_two_word):
        #print(wrd)
        if wrd in double_dictionary:
            double_dictionary[wrd]=double_dictionary[wrd]+1;
        else:
            double_dictionary[wrd]=1;
    
    #update triple dictionary  
    for wrd in (list_with_three_word):
        if wrd in triple_dictionary:
            triple_dictionary[wrd]=triple_dictionary[wrd]+1;
        else:
            triple_dictionary[wrd]=1;

        


# In[6]:


#store into a database
write=xlwt.Workbook()
write_single = write.add_sheet("single")
write_double = write.add_sheet("double")
write_triple = write.add_sheet("triple")
write_already_done = write.add_sheet("already_done")
row=0
column=0
for key, cnt in single_dictionary.items():
    if(row>65530):
        row=0
        column=column+2
    write_single.write(row, column, key)
    write_single.write(row, column+1, cnt)
    row=row+1
row=0
column=0
for key, cnt in double_dictionary.items():
    if(row>65530):
        row=0
        column=column+2
    write_double.write(row, column, key)
    write_double.write(row, column+1, cnt)
    row=row+1

row=0
column=0
for key, cnt in triple_dictionary.items():
    if(row>65530):
        row=0
        column=column+2
    write_triple.write(row, column, key)
    write_triple.write(row, column+1, cnt)
    row=row+1

row=0
column=0
for key, cnt in already_done_dictionary.items(): 
    if(row>65530):
        row=0
        column=column+2
    write_already_done.write(row, column, key)
    write_already_done.write(row, column+1, cnt)
    row=row+1
os.remove("database.xls")
write.save("database.xls")


# In[ ]:




