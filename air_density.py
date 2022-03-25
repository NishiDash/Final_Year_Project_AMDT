from matplotlib.pyplot import text
from pymysql import NULL
from datetime import date;
from colorama import Fore
import openpyxl
import pandas as pd
def air_density(file1,attribute):
    try:
        print(Fore.RESET)
        path = './excel files/'+file1
        
        df = pd.read_excel(path)
        n = df.count()[0]+2
        m =  df.count()[2]+2

        obj = openpyxl.load_workbook(path.strip())
        today = date.today()

        ws1 = obj.create_sheet("Sheet2")
        ws1.title= "output_"+str(today)+"_air_density"
        
        sheet_obj = obj["Sheet1"]
        sheet_obj1 = obj["output_"+str(today)+"_air_density"]


        sheet_obj['E1']='TRIM_id'
        sheet_obj['F1']=attribute+' wrt id'
        sheet_obj['G1']=attribute+'_discrepancy'
        sheet_obj['H1']='Turbine serial Number'
        sheet_obj['I1']='ID Present in AWS?'

        sheet_obj1['A1']='id'
        sheet_obj1['B1']=attribute+'(AWS)'
        sheet_obj1['C1']='Turbine serial Number'
        sheet_obj1['D1']=attribute+'(RAMP)'
        sheet_obj1['E1']=attribute+'_discrepancy'
    
        for j in range(2,n):
            index1='E'+str(j)
            index2='F'+str(j)
            index3='G'+str(j)
                
        
            formula1 = '=TRIM(A'+str(j)+')' 
            formula2 = '=if(ISNA(VLOOKUP(E'+str(j)+',C:D,2,FALSE)),"Id not in RAMP",if(LEN(VLOOKUP(E'+str(j)+',C:D,2,FALSE))=0,"",VLOOKUP(E'+str(j)+',C:D,2,FALSE)))'
            formula3 = '=if(F'+str(j)+'<>"Id not in RAMP",IF(AND(OR(B'+str(j)+'="NULL",B'+str(j)+'=""),OR(F'+str(j)+'="NULL",F'+str(j)+'="")),"matching",if(F'+str(j)+'=B'+str(j)+',"matching","not matching")),"Id not in RAMP")'
                
            sheet_obj[index1]= formula1 # =trim(A2) at D2  
            sheet_obj[index2]= formula2 # =VLOOKUP(D2,[AD_wmf.xlsx]Sheet1!$A:$B,2,FALSE)
            sheet_obj[index3]= formula3 # =if(C2<>"id not in ramp", if(c2=b2,"matching","not matching"),"")
        print(m)
        for j in range(2,m):
            index1 = 'H'+str(j)
            index2 = 'I'+str(j)

            formula1 = '=TRIM(C'+str(j)+')'
            formula2 = '=if(ISNA(VLOOKUP(H'+str(j)+',A:B,2,FALSE)),"Id not in AWS",if(LEN(VLOOKUP(H'+str(j)+',A:B,2,FALSE))=0,"",VLOOKUP(H'+str(j)+',A:B,2,FALSE)))'

            sheet_obj[index1]=formula1
            sheet_obj[index2]=formula2
            
        for j in range(2,n):
            ind1 = 'A'+str(j)
            ind2 = 'B'+str(j)
            ind3 = 'C'+str(j)
            ind4 = 'D'+str(j)
            ind5 = 'E'+str(j)

            frm1 = '=Sheet1!A'+str(j)
            frm2 = '=Sheet1!B'+str(j)
            frm3 = '=Sheet1!A'+str(j)
            frm4 = '=Sheet1!F'+str(j)
            frm5 = '=Sheet1!G'+str(j)

            sheet_obj1[ind1] = frm1
            sheet_obj1[ind2] = frm2
            sheet_obj1[ind3] = frm3
            sheet_obj1[ind4] = frm4
            sheet_obj1[ind5] = frm5
        i = 2
        for j in range(n,n+m):
            ind1 = 'C'+str(j)
            ind2 = 'D'+str(j)
            ind3 = 'E'+str(j)

            frm1 = '=if(Sheet1!I'+str(i)+'="Id not in AWS",Sheet1!H'+str(i)+',"")'
            frm2 = '=if(Sheet1!I'+str(i)+'="Id not in AWS","","")'
            frm3 = '=if(Sheet1!I'+str(i)+'="Id not in AWS","Id not in AWS","")'

            sheet_obj1[ind1]=frm1
            sheet_obj1[ind2]=frm2
            sheet_obj1[ind3]=frm3
            i=i+1
        
        obj.save(path)
    

    except Exception as e:
        print(e)
        print (Fore.RED + "Error : The file does not found")
        return ("Error : "+str(e)[11:28]+" to "+ file1)
    print(Fore.GREEN + "###################### Successfully! Excel file has been read/written. ##############################")
    return("Successfully the excel file has been read/written.")


