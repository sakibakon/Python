#!/usr/bin/env python
# coding: utf-8

# In[33]:


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

database=rb.sheet_by_index(1)
double_dictionary = {}
for i in range(database.nrows):
    key =database.cell(i, 0).value
    co =int(database.cell(i, 1).value)
    double_dictionary[key]=co

database=rb.sheet_by_index(2)
triple_dictionary = {}
for i in range(database.nrows):
    key =database.cell(i, 0).value
    co =int(database.cell(i, 1).value)
    triple_dictionary[key]=co

# store already processed news
database=rb.sheet_by_index(3)
already_done_dictionary = {}
for i in range(database.nrows):
    key =database.cell(i, 0).value
    co =int(database.cell(i, 1).value)
    already_done_dictionary[key]=co
os.remove("database.xls")


# In[34]:


read=xlrd.open_workbook("collected_news.xls")
read_news=read.sheet_by_name("first")
for i in range(read_news.nrows):
    news=read_news.cell(i,2).value
    check = read_news.cell(i,1).value
    #if news is already process then continue
    if check in already_done_dictionary:
        continue;
    already_done_dictionary[check]=1;
    #make three list with one, two and three words.
    list_with_single_word= news.split()
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

        


# In[35]:


#store into a database
write=xlwt.Workbook()
write_single = write.add_sheet("single")
write_double = write.add_sheet("double")
write_triple = write.add_sheet("triple")
write_already_done = write.add_sheet("already_done")
j=0
for key, cnt in single_dictionary.items(): 
    write_single.write(j, 0, key)
    write_single.write(j, 1, cnt)
    j=j+1
j=0
for key, cnt in double_dictionary.items(): 
    write_double.write(j, 0, key)
    write_double.write(j, 1, cnt)
    j=j+1
j=0
for key, cnt in triple_dictionary.items(): 
    write_triple.write(j, 0, key)
    write_triple.write(j, 1, cnt)
    j=j+1
j=0
for key, cnt in already_done_dictionary.items(): 
    write_already_done.write(j, 0, key)
    write_already_done.write(j, 1, cnt)
    j=j+1
write.save("database.xls")


# In[ ]:




