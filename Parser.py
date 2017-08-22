import xlwt
import xlrd
#import os.path
#import os
#print os.path.abspath(os.curdir) 
def headline(row):
    a ,b, c, d = row
    if b =='' and c == '' and d == '':
        return True
    else:
        return False
        
wbk = xlwt.Workbook()
Newsheet = wbk.add_sheet('Employment Report',cell_overwrite_ok=True)#name of the sheet in which new data is there
Newsheet.write(0,0,'Employer')

Newsheet.write(0,1,'Position')
Newsheet.write(0,2,'Location')
Newsheet.write(0,3,'State')
Newsheet.write(0,4,'School')
Newsheet.write(0,5,'Major')
Newsheet.write(0,6,'Report')
Newsheet.write(0,7,'Year')

Gradsheet = wbk.add_sheet('Graduate School Report',cell_overwrite_ok=True)

r = 1


f = xlrd.open_workbook('2015convert.xlsx')#workbook from where data is extracted
#g = xlwt.workbook('Data_2.xls')#workbook from where data is extracted

for x in range(f.nsheets):
    
    sheet = f. sheet_by_index(x)
    
   

    for i in range( 0 , sheet.nrows-1 ):
         
         row = sheet.row_values(i)
         
         if headline(row):
             if row[0]=="EMPLOYMENT DETAIL REPORT" or row[0] == "GRADUATE SCHOOL DETAIL REPORT":
                report = row[0]
                
                
             elif 'SCHOOL' in row[0] or 'COLLEGE' in row[0]:
                school = row[0]
                
                          
     
             elif sheet.row_values(i+1)[0] =="employer" or sheet.row_values(i+1)[0] =="institution" :
                  dept = row[0]
                  
                  
             else:
                 Newsheet.write(r-1,0,sheet.row_values(i-1)[0] + " " + sheet.row_values(i)[0])
     
         else:
             if row[0]=="employer" or row[0] == "institution":
                 continue;
                 
             elif row[0] =='':
                    print(row)
                    Newsheet.write(r-1,1,sheet.row_values(i-1)[1] + " " + sheet.row_values(i)[1])
                    continue;
             else:
                 col = 0
                 for element in row:                           
                        Newsheet.write(r,col,element)
                        #Employsheet.write(r,col,element)
                        Newsheet.write(r,4,school)
                        Newsheet.write(r,5,dept)
                        Newsheet.write(r,6,report)
                        Newsheet.write(r,7,2015)
                    
                        col += 1
                 r += 1
    
wbk.save('DATA2015.xls') #excel sheet name when saved in the folder
#%%
""" SEPARATING THE TWO REPORTS """

import xlwt
import xlrd
#import os.path
#import os
#print os.path.abspath(os.curdir) 
def headline(row):
    a ,b, c, d = row
    if b =='' and c == '' and d == '':
        return True
    else:
        return False
        
wbk = xlwt.Workbook()
Employsheet = wbk.add_sheet('Employment Report',cell_overwrite_ok=True)
Gradsheet = wbk.add_sheet('Graduate School Report',cell_overwrite_ok=True)
#name of the sheet in which new data is there
Employsheet.write(0,0,'Employer')
Employsheet.write(0,1,'Position')
Employsheet.write(0,2,'Location')
Employsheet.write(0,3,'State')
Employsheet.write(0,4,'School')
Employsheet.write(0,5,'Major')
Employsheet.write(0,6,'Report')
Employsheet.write(0,7,'Year')

Gradsheet.write(0,0,'Institution')
Gradsheet.write(0,1,'Program')
Gradsheet.write(0,2,'Location')
Gradsheet.write(0,3,'State')
Gradsheet.write(0,4,'Year')
Gradsheet.write(0,5,'Major')
Gradsheet.write(0,6,'Report')
Gradsheet.write(0,7,'Year')



er = 1
gr = 1


f = xlrd.open_workbook('DATA2013.xls')#workbook from where data is extracted
#g = xlwt.workbook('Data_2.xls')#workbook from where data is extracted


    
sheet = f. sheet_by_index(0)

   

for i in range( 0 , sheet.nrows ):
     
     row = sheet.row_values(i)
     
     if row[6]=="EMPLOYMENT DETAIL REPORT":
         print(row)
         c=0
         for element in row:
             
             Employsheet.write(er,c,element)
             c+= 1
         er = er + 1
     elif row[6] == "GRADUATE SCHOOL DETAIL REPORT":
          print(row)
          c=0
          for element in row:
              
              Gradsheet.write(gr,c,element)
              c+=1
          gr =gr + 1
              
     
     
    
wbk.save('2013.xls') #excel sheet name when saved in the folder
#%%


import xlwt
import xlrd
#import os.path
#import os
#print os.path.abspath(os.curdir) 
def headline(row):
    
    if row[1] =='' and row[2] == '' and row[3] == '':
        return True
    else:
        return False
        
wbk = xlwt.Workbook()
Newsheet = wbk.add_sheet('Reports',cell_overwrite_ok=True)#name of the sheet in which new data is there
Newsheet.write(0,0,'Employer')

Newsheet.write(0,1,'Position')
Newsheet.write(0,2,'Location')
Newsheet.write(0,3,'State')
Newsheet.write(0,4,'School')
Newsheet.write(0,5,'Major')
Newsheet.write(0,6,'Report')
Newsheet.write(0,7,'Year')

Gradsheet = wbk.add_sheet('Graduate School Report',cell_overwrite_ok=True)

r = 1


f = xlrd.open_workbook('2014convert.xlsx')#workbook from where data is extracted
#g = xlwt.workbook('Data_2.xls')#workbook from where data is extracted

for x in range(f.nsheets):
    
    sheet = f. sheet_by_index(x)
    
   

    for i in range( 0 , sheet.nrows-1 ):
         
         row = sheet.row_values(i)
         
         if headline(row):
             if row[0]=="EMPLOYMENT DETAIL REPORT" or row[0] == "GRADUATE SCHOOL DETAIL REPORT":
                report = row[0]
                
                
             elif 'SCHOOL' in row[0] or 'COLLEGE' in row[0]:
                school = row[0]
                
                          
     
             elif sheet.row_values(i+1)[0] =="employer" or sheet.row_values(i+1)[0] =="institution" :
                  dept = row[0]
                  
                  
             else:
                 Newsheet.write(r-1,0,sheet.row_values(i-1)[0] + " " + sheet.row_values(i)[0])
     
         else:
             if row[0]=="employer" or row[0] == "institution":
                 continue;
                 
             elif row[0] =='':
                    print(row)
                    Newsheet.write(r-1,1,sheet.row_values(i-1)[1] + " " + sheet.row_values(i)[1])
                    continue;
             else:
                 col = 0
                 for element in row:                           
                        Newsheet.write(r,col,element)
                        #Employsheet.write(r,col,element)
                        Newsheet.write(r,4,school)
                        Newsheet.write(r,5,dept)
                        Newsheet.write(r,6,report)
                        Newsheet.write(r,7,2014)
                    
                        col += 1
                 r += 1
    
wbk.save('DATA2014.xls') #excel sheet name when saved in the folder
#%%

import xlwt
import xlrd
#import os.path
#import os
#print os.path.abspath(os.curdir) 
def headline(row):
    
    if row[1] =='' and row[2] =='':
        return True
    else:
        return False
        
wbk = xlwt.Workbook()
Newsheet = wbk.add_sheet('Employment Report',cell_overwrite_ok=True)#name of the sheet in which new data is there
Newsheet.write(0,0,'Employer')

Newsheet.write(0,1,'Position')
Newsheet.write(0,2,'Location')
Newsheet.write(0,3,'State')
Newsheet.write(0,4,'School')
Newsheet.write(0,5,'Major')
Newsheet.write(0,6,'Report')
Newsheet.write(0,7,'Year')

Gradsheet = wbk.add_sheet('Reports',cell_overwrite_ok=True)

r = 1


f = xlrd.open_workbook('2013convert.xlsx')#workbook from where data is extracted
#g = xlwt.workbook('Data_2.xls')#workbook from where data is extracted

for x in range(f.nsheets):
    
    sheet = f. sheet_by_index(x)
    
   

    for i in range( 0 , sheet.nrows-1 ):
         
         row = sheet.row_values(i)
         
         if headline(row):
             if row[0]=="EMPLOYMENT DETAIL REPORT" or row[0] == "GRADUATE SCHOOL DETAIL REPORT":
                report = row[0]
                
                
             elif 'School' in row[0] or 'College' in row[0]:
                school = row[0]
                
                          
     
             elif sheet.row_values(i+1)[0] =="employer" or sheet.row_values(i+1)[0] =="institution" :
                  dept = row[0]
                  
                  
             else:
                 Newsheet.write(r-1,0,sheet.row_values(i-1)[0] + " " + sheet.row_values(i)[0])
     
         else:
             if row[0]=="employer" or row[0] == "institution":
                 continue;
                 
             elif row[0] =='':
                    print(row)
                    Newsheet.write(r-1,1,sheet.row_values(i-1)[1] + " " + sheet.row_values(i)[1])
                    continue;
             else:
                 col = 0
                 for element in row:
                        
                            Newsheet.write(r,col,element)
                        #Employsheet.write(r,col,element)
                            Newsheet.write(r,4,school)
                            Newsheet.write(r,5,dept)
                            Newsheet.write(r,6,report)
                            Newsheet.write(r,7,2013)
                        
                            col += 1
                 r += 1
    
wbk.save('DATA2013.xls') #excel sheet name when saved in the folder

#%%

import xlwt
import xlrd
#import os.path
#import os
#print os.path.abspath(os.curdir) 
def headline(row):
    
    if row[1] =='' and row[2] =='':
        return True
    else:
        return False
        
wbk = xlwt.Workbook()
Newsheet = wbk.add_sheet('Employment Report',cell_overwrite_ok=True)#name of the sheet in which new data is there
Newsheet.write(0,0,'Employer')

Newsheet.write(0,1,'Position')
Newsheet.write(0,2,'Location')
Newsheet.write(0,3,'State')
Newsheet.write(0,4,'School')
Newsheet.write(0,5,'Major')
Newsheet.write(0,6,'Report')
Newsheet.write(0,7,'Year')

Gradsheet = wbk.add_sheet('Reports',cell_overwrite_ok=True)

r = 1


f = xlrd.open_workbook('2012convert.xlsx')#workbook from where data is extracted
#g = xlwt.workbook('Data_2.xls')#workbook from where data is extracted

for x in range(f.nsheets):
    
    sheet = f. sheet_by_index(x)
    
   

    for i in range( 0 , sheet.nrows-1 ):
         
         row = sheet.row_values(i)
         
         if headline(row):
             if row[0]=="EMPLOYMENT DETAIL REPORT" or row[0] == "GRADUATE SCHOOL DETAIL REPORT":
                report = row[0]
                
                
             elif 'School' in row[0] or 'College' in row[0]:
                school = row[0]
                
                          
     
             elif sheet.row_values(i+1)[0] =="employer" or sheet.row_values(i+1)[0] =="institution" :
                  dept = row[0]
                  
                  
             else:
                 Newsheet.write(r-1,0,sheet.row_values(i-1)[0] + " " + sheet.row_values(i)[0])
     
         else:
             if row[0]=="employer" or row[0] == "institution":
                 continue;
                 
             elif row[0] =='':
                    print(row)
                    Newsheet.write(r-1,1,sheet.row_values(i-1)[1] + " " + sheet.row_values(i)[1])
                    continue;
             else:
                 col = 0
                 for element in row:
                        
                            Newsheet.write(r,col,element)
                        #Employsheet.write(r,col,element)
                            Newsheet.write(r,4,school)
                            Newsheet.write(r,5,dept)
                            Newsheet.write(r,6,report)
                            Newsheet.write(r,7,2012)
                        
                            col += 1
                 r += 1
    
wbk.save('DATA2012.xls') #excel sheet name when saved in the folder
#%%
import xlwt
import xlrd
#import os.path
#import os
#print os.path.abspath(os.curdir) 
def headline(row):
    a ,b, c, d = row
    if b =='' and c == '' and d == '':
        return True
    else:
        return False
        
wbk = xlwt.Workbook()
Newsheet = wbk.add_sheet('Employment Report',cell_overwrite_ok=True)#name of the sheet in which new data is there
Newsheet.write(0,0,'Employer')

Newsheet.write(0,1,'Position')
Newsheet.write(0,2,'Location')
Newsheet.write(0,3,'State')
Newsheet.write(0,4,'School')
Newsheet.write(0,5,'Major')
Newsheet.write(0,6,'Report')
Newsheet.write(0,7,'Year')

Gradsheet = wbk.add_sheet('Graduate School Report',cell_overwrite_ok=True)

r = 1


f = xlrd.open_workbook('2015convert.xlsx')#workbook from where data is extracted
#g = xlwt.workbook('Data_2.xls')#workbook from where data is extracted

for x in range(f.nsheets):
    
    sheet = f. sheet_by_index(x)
    
   

    for i in range( 0 , sheet.nrows-1 ):
         
         row = sheet.row_values(i)
         
         if headline(row):
             if row[0]=="EMPLOYMENT DETAIL REPORT" or row[0] == "GRADUATE SCHOOL DETAIL REPORT":
                report = row[0]
                
                
             elif 'SCHOOL' in row[0] or 'COLLEGE' in row[0]:
                school = row[0]
                
                          
     
             elif sheet.row_values(i+1)[0] =="employer" or sheet.row_values(i+1)[0] =="institution" :
                  dept = row[0]
                  
                  
             else:
                 Newsheet.write(r-1,0,sheet.row_values(i-1)[0] + " " + sheet.row_values(i)[0])
     
         else:
             if row[0]=="employer" or row[0] == "institution":
                 continue;
                 
             elif row[0] =='':
                    print(row)
                    Newsheet.write(r-1,1,sheet.row_values(i-1)[1] + " " + sheet.row_values(i)[1])
                    continue;
             else:
                 col = 0
                 for element in row:                           
                        Newsheet.write(r,col,element)
                        #Employsheet.write(r,col,element)
                        Newsheet.write(r,4,school)
                        Newsheet.write(r,5,dept)
                        Newsheet.write(r,6,report)
                        Newsheet.write(r,7,2015)
                    
                        col += 1
                 r += 1
    
wbk.save('DATA2014.xls') #excel sheet name when saved in the folder
#%%
""" SEPARATING THE TWO REPORTS """

import xlwt
import xlrd
#import os.path
#import os
#print os.path.abspath(os.curdir) 
def headline(row):
    a ,b, c, d = row
    if b =='' and c == '' and d == '':
        return True
    else:
        return False
        
wbk = xlwt.Workbook()
Employsheet = wbk.add_sheet('Employment Report',cell_overwrite_ok=True)
Gradsheet = wbk.add_sheet('Graduate School Report',cell_overwrite_ok=True)
#name of the sheet in which new data is there
Employsheet.write(0,0,'Employer')
Employsheet.write(0,1,'Position')
Employsheet.write(0,2,'Location')
Employsheet.write(0,3,'State')
Employsheet.write(0,4,'School')
Employsheet.write(0,5,'Major')
Employsheet.write(0,6,'Report')
Employsheet.write(0,7,'Year')

Gradsheet.write(0,0,'Institution')
Gradsheet.write(0,1,'Program')
Gradsheet.write(0,2,'Location')
Gradsheet.write(0,3,'State')
Gradsheet.write(0,4,'Year')
Gradsheet.write(0,5,'Major')
Gradsheet.write(0,6,'Report')
Gradsheet.write(0,7,'Year')



er = 1
gr = 1


f = xlrd.open_workbook('DATA2014.xls')#workbook from where data is extracted
#g = xlwt.workbook('Data_2.xls')#workbook from where data is extracted


    
sheet = f. sheet_by_index(0)

   

for i in range( 0 , sheet.nrows ):
     
     row = sheet.row_values(i)
     
     if row[6]=="EMPLOYMENT DETAIL REPORT":
         print(row)
         c=0
         for element in row:
             
             Employsheet.write(er,c,element)
             c+= 1
         er = er + 1
     elif row[6] == "GRADUATE SCHOOL DETAIL REPORT":
          print(row)
          c=0
          for element in row:
              
              Gradsheet.write(gr,c,element)
              c+=1
          gr =gr + 1
              
     
     
    
wbk.save('2014.xls') #excel sheet name when saved in the folder
#%%


import xlwt
import xlrd
#import os.path
#import os
#print os.path.abspath(os.curdir) 
def headline(row):
    
    if row[1] =='' and row[2] == '' and row[3] == '':
        return True
    else:
        return False
        
wbk = xlwt.Workbook()
Newsheet = wbk.add_sheet('Reports',cell_overwrite_ok=True)#name of the sheet in which new data is there
Newsheet.write(0,0,'Employer')

Newsheet.write(0,1,'Position')
Newsheet.write(0,2,'Location')
Newsheet.write(0,3,'State')
Newsheet.write(0,4,'School')
Newsheet.write(0,5,'Major')
Newsheet.write(0,6,'Report')
Newsheet.write(0,7,'Year')

Gradsheet = wbk.add_sheet('Graduate School Report',cell_overwrite_ok=True)

r = 1


f = xlrd.open_workbook('2014convert.xlsx')#workbook from where data is extracted
#g = xlwt.workbook('Data_2.xls')#workbook from where data is extracted

for x in range(f.nsheets):
    
    sheet = f. sheet_by_index(x)
    
   

    for i in range( 0 , sheet.nrows-1 ):
         
         row = sheet.row_values(i)
         
         if headline(row):
             if row[0]=="EMPLOYMENT DETAIL REPORT" or row[0] == "GRADUATE SCHOOL DETAIL REPORT":
                report = row[0]
                
                
             elif 'SCHOOL' in row[0] or 'COLLEGE' in row[0]:
                school = row[0]
                
                          
     
             elif sheet.row_values(i+1)[0] =="employer" or sheet.row_values(i+1)[0] =="institution" :
                  dept = row[0]
                  
                  
             else:
                 Newsheet.write(r-1,0,sheet.row_values(i-1)[0] + " " + sheet.row_values(i)[0])
     
         else:
             if row[0]=="employer" or row[0] == "institution":
                 continue;
                 
             elif row[0] =='':
                    print(row)
                    Newsheet.write(r-1,1,sheet.row_values(i-1)[1] + " " + sheet.row_values(i)[1])
                    continue;
             else:
                 col = 0
                 for element in row:                           
                        Newsheet.write(r,col,element)
                        #Employsheet.write(r,col,element)
                        Newsheet.write(r,4,school)
                        Newsheet.write(r,5,dept)
                        Newsheet.write(r,6,report)
                        Newsheet.write(r,7,2014)
                    
                        col += 1
                 r += 1
    
wbk.save('DATA2014.xls') #excel sheet name when saved in the folder
#%%

import xlwt
import xlrd
#import os.path
#import os
#print os.path.abspath(os.curdir) 
def headline(row):
    
    if row[1] =='' and row[2] =='':
        return True
    else:
        return False
        
wbk = xlwt.Workbook()
Newsheet = wbk.add_sheet('Employment Report',cell_overwrite_ok=True)#name of the sheet in which new data is there
Newsheet.write(0,0,'Employer')

Newsheet.write(0,1,'Position')
Newsheet.write(0,2,'Location')
Newsheet.write(0,3,'State')
Newsheet.write(0,4,'School')
Newsheet.write(0,5,'Major')
Newsheet.write(0,6,'Report')
Newsheet.write(0,7,'Year')

Gradsheet = wbk.add_sheet('Reports',cell_overwrite_ok=True)

r = 1


f = xlrd.open_workbook('2013convert.xlsx')#workbook from where data is extracted
#g = xlwt.workbook('Data_2.xls')#workbook from where data is extracted

for x in range(f.nsheets):
    
    sheet = f. sheet_by_index(x)
    
   

    for i in range( 0 , sheet.nrows-1 ):
         
         row = sheet.row_values(i)
         
         if headline(row):
             if row[0]=="EMPLOYMENT DETAIL REPORT" or row[0] == "GRADUATE SCHOOL DETAIL REPORT":
                report = row[0]
                
                
             elif 'School' in row[0] or 'College' in row[0]:
                school = row[0]
                
                          
     
             elif sheet.row_values(i+1)[0] =="employer" or sheet.row_values(i+1)[0] =="institution" :
                  dept = row[0]
                  
                  
             else:
                 Newsheet.write(r-1,0,sheet.row_values(i-1)[0] + " " + sheet.row_values(i)[0])
     
         else:
             if row[0]=="employer" or row[0] == "institution":
                 continue;
                 
             elif row[0] =='':
                    print(row)
                    Newsheet.write(r-1,1,sheet.row_values(i-1)[1] + " " + sheet.row_values(i)[1])
                    continue;
             else:
                 col = 0
                 for element in row:
                        
                            Newsheet.write(r,col,element)
                        #Employsheet.write(r,col,element)
                            Newsheet.write(r,4,school)
                            Newsheet.write(r,5,dept)
                            Newsheet.write(r,6,report)
                            Newsheet.write(r,7,2013)
                        
                            col += 1
                 r += 1
    
wbk.save('DATA2013.xls') #excel sheet name when saved in the folder

#%%

import xlwt
import xlrd
#import os.path
#import os
#print os.path.abspath(os.curdir) 
def headline(row):
    
    if row[1] =='' and row[2] =='':
        return True
    else:
        return False
        
wbk = xlwt.Workbook()
Newsheet = wbk.add_sheet('Employment Report',cell_overwrite_ok=True)#name of the sheet in which new data is there
Newsheet.write(0,0,'Employer')

Newsheet.write(0,1,'Position')
Newsheet.write(0,2,'Location')
Newsheet.write(0,3,'State')
Newsheet.write(0,4,'School')
Newsheet.write(0,5,'Major')
Newsheet.write(0,6,'Report')
Newsheet.write(0,7,'Year')



r = 1


f = xlrd.open_workbook('2011convert.xlsx')#workbook from where data is extracted
#g = xlwt.workbook('Data_2.xls')#workbook from where data is extracted

for x in range(f.nsheets):
    
    sheet = f. sheet_by_index(x)
    
   

    for i in range( 0 , sheet.nrows-1 ):
         
         row = sheet.row_values(i)
         
         if headline(row):
             if row[0]=="Employment Detail Report" or row[0] == "Graduate School Detail Report":
                report = row[0]
                
                
             elif 'School' in row[0] or 'College' in row[0]:
                school = row[0]
                
                          
     
             elif sheet.row_values(i+1)[0] =="employer" or sheet.row_values(i+1)[0] =="institution" :
                  dept = row[0]
                  
                  
             else:
                 Newsheet.write(r-1,0,sheet.row_values(i-1)[0] + " " + sheet.row_values(i)[0])
     
         else:
             if row[0]=="employer" or row[0] == "institution":
                 continue;
                 
             elif row[0] =='':
                    print(row)
                    print(x)
                    print(i)
                    Newsheet.write(r-1,1,sheet.row_values(i-1)[1] + " " + sheet.row_values(i)[1])
                    continue;
             else:
                 col = 0
                 for element in row:
                        
                            Newsheet.write(r,col,element)
                        #Employsheet.write(r,col,element)
                            Newsheet.write(r,4,school)
                            Newsheet.write(r,5,dept)
                            Newsheet.write(r,6,report)
                            Newsheet.write(r,7,2011)
                        
                            col += 1
                 r += 1
    
wbk.save('DATA2011.xls') #excel sheet name when saved in the folder

#%%
import xlwt
import xlrd
#import os.path
#import os
#print os.path.abspath(os.curdir) 
def headline(row):
    
    if row[1] =='' and row[2] =='':
        return True
    else:
        return False
        
wbk = xlwt.Workbook()
Newsheet = wbk.add_sheet('Employment Report',cell_overwrite_ok=True)#name of the sheet in which new data is there
Newsheet.write(0,0,'Employer')

Newsheet.write(0,1,'Position')
Newsheet.write(0,2,'Location')
Newsheet.write(0,3,'State')
Newsheet.write(0,4,'School')
Newsheet.write(0,5,'Major')
Newsheet.write(0,6,'Report')
Newsheet.write(0,7,'Year')

Gradsheet = wbk.add_sheet('Reports',cell_overwrite_ok=True)

r = 1


f = xlrd.open_workbook('2010convert.xlsx')#workbook from where data is extracted
#g = xlwt.workbook('Data_2.xls')#workbook from where data is extracted

for x in range(f.nsheets):
    
    sheet = f. sheet_by_index(x)
    
   

    for i in range( 0 , sheet.nrows-1 ):
         
         row = sheet.row_values(i)
         
         if headline(row):
             if row[0]=="Employment Detail Report" or row[0] == "Graduate School Detail Report":
                report = row[0]
                
                
             elif 'School' in row[0] or 'College' in row[0]:
                school = row[0]
                
                          
     
             elif sheet.row_values(i+1)[0] =="employer" or sheet.row_values(i+1)[0] =="institution" :
                  dept = row[0]
                  
                  
             else:
                 Newsheet.write(r-1,0,sheet.row_values(i-1)[0] + " " + sheet.row_values(i)[0])
     
         else:
             if row[0]=="employer" or row[0] == "institution":
                 continue;
                 
             elif row[0] =='':
                    print(row)
                    print(x)
                    print(i)
                    Newsheet.write(r-1,1,sheet.row_values(i-1)[1] + " " + sheet.row_values(i)[1])
                    continue;
             else:
                 col = 0
                 for element in row:
                        
                            Newsheet.write(r,col,element)
                        #Employsheet.write(r,col,element)
                            Newsheet.write(r,4,school)
                            Newsheet.write(r,5,dept)
                            Newsheet.write(r,6,report)
                            Newsheet.write(r,7,2010)
                        
                            col += 1
                 r += 1
    
wbk.save('DATA2010.xls') #excel sheet name when saved in the folder
#%%

import xlwt
import xlrd
#import os.path
#import os
#print os.path.abspath(os.curdir) 
def headline(row):
    a ,b, c, d = row
    if b =='' and c == '' and d == '':
        return True
    else:
        return False
        
wbk = xlwt.Workbook()
Employsheet = wbk.add_sheet('Employment Report',cell_overwrite_ok=True)
Gradsheet = wbk.add_sheet('Graduate School Report',cell_overwrite_ok=True)
#name of the sheet in which new data is there
Employsheet.write(0,0,'Employer')
Employsheet.write(0,1,'Position')
Employsheet.write(0,2,'Location')
Employsheet.write(0,3,'State')
Employsheet.write(0,4,'School')
Employsheet.write(0,5,'Major')
Employsheet.write(0,6,'Report')
Employsheet.write(0,7,'Year')

Gradsheet.write(0,0,'Institution')
Gradsheet.write(0,1,'Program')
Gradsheet.write(0,2,'Location')
Gradsheet.write(0,3,'State')
Gradsheet.write(0,4,'Year')
Gradsheet.write(0,5,'Major')
Gradsheet.write(0,6,'Report')
Gradsheet.write(0,7,'Year')



er = 1
gr = 1

for i in range(6):
    file = "Data201" + str(i)+".xls"
    f = xlrd.open_workbook(file) 
    sheet = f. sheet_by_index(0)

   

    for i in range( 0 , sheet.nrows ):
     
     row = sheet.row_values(i)
     
     if row[6]=="EMPLOYMENT DETAIL REPORT" or row[6]=="Employment Detail Report" :
         print(row)
         c=0
         for element in row:
             
             Employsheet.write(er,c,element)
             c+= 1
         er = er + 1
     elif row[6] == "GRADUATE SCHOOL DETAIL REPORT" or row[6] == "Graduate School Detail Report" :
          print(row)
          c=0
          for element in row:
              
              Gradsheet.write(gr,c,element)
              c+=1
          gr =gr + 1
              
     
     
    
wbk.save('Data.xls') #excel sheet name when saved in the folder
#%%