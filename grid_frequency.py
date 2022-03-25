from matplotlib.pyplot import text
from pymysql import NULL
from datetime import date;
from colorama import Fore
import openpyxl
import pandas as pd

def grid_frequency(file1, attribute):
    try:
        print(Fore.RESET)
        path = './excel files/'+file1
        df = pd.read_excel(path)
        m = df.count()[2]+2
        n = df.count()[0]+2
        
        obj = openpyxl.load_workbook(path.strip())
        today = date.today()

        ws1 = obj.create_sheet("Sheet2")
        ws1.title= "output_"+str(today)+"_GridFrequency"
        
        sheet_obj = obj["Sheet1"]
        sheet_obj1 = obj["output_"+str(today)+"_GridFrequency"]

        sheet_obj.insert_cols(0,amount=1)
        sheet_obj.insert_cols(4,amount=1)
        
        sheet_obj['A1']='TRIM_id'
        sheet_obj['D1']='Trim_Turbine Serial Number'
        sheet_obj['G1']=attribute+' wrt id'
        sheet_obj['H1']=attribute+'_Discrepancy'
        sheet_obj['I1']='ID in AWS?'
        

        sheet_obj1['A1']='Id'
        sheet_obj1['B1']=attribute+" (AWS)"
        sheet_obj1['C1']='Turbine Serial Number'
        sheet_obj1['D1']=attribute+" (RAMP)"
        sheet_obj1['E1']=attribute+'_Discrepancy'
            
        for j in range(2,n):
            index1 = 'A'+str(j)
            index2 = 'G'+str(j)
            
            formula1 = '=TRIM(B'+str(j)+')'
            formula2 = '=if(ISNA(VLOOKUP(A'+str(j)+',D:F,3,FALSE)),"Id not in ramp",if(len(VLOOKUP(A'+str(j)+',D:F,3,FALSE))=0,"",VLOOKUP(A'+str(j)+',D:F,3,FALSE)))'
            
            sheet_obj[index1]= formula1
            sheet_obj[index2]= formula2

        for j in range(2,n):
            index='H'+str(j)
            formula = '=IF(G'+str(j)+'="Id not in ramp","Id not in ramp",IF(AND(OR(C'+str(j)+'="NULL",C'+str(j)+'=""),OR(G'+str(j)+'="NULL",G'+str(j)+'="")),"matching",IF(G'+str(j)+'=C'+str(j)+',"matching","not matching")))'
            sheet_obj[index]= formula  
         
        for j in range(2,n):
            ind1 = 'A'+str(j)
            ind2 = 'B'+str(j)
            ind3 = 'C'+str(j)
            ind4 = 'D'+str(j)
            ind5 = 'E'+str(j)
            

            frm1 = '=Sheet1!B'+str(j)
            frm2 = '=Sheet1!C'+str(j)
            frm3 = '=Sheet1!B'+str(j)
            frm4 = '=Sheet1!G'+str(j)
            frm5 = '=Sheet1!H'+str(j)
            

            sheet_obj1[ind1] = frm1
            sheet_obj1[ind2] = frm2
            sheet_obj1[ind3] = frm3
            sheet_obj1[ind4] = frm4
            sheet_obj1[ind5] = frm5
            
        for j in range(2,m):
            index1 = 'D'+str(j)
            index2 = 'I'+str(j)

            formula1 = '=TRIM(E'+str(j)+')'
            formula2 = '=if(ISNA(vlookup(D'+str(j)+',A:C,3,false)),"Id not in AWS",if( len(vlookup(D'+str(j)+',A:C,3,false))=0,"",vlookup(D'+str(j)+',A:C,3,false) ))'
        
            sheet_obj[index1]=formula1
            sheet_obj[index2]=formula2
        i=2
        for j in range(n,m+n):
            ind1 = 'C'+str(j)
            ind2 = 'D'+str(j)
            ind3 = 'E'+str(j)

            frm1 = '=if(Sheet1!I'+str(i)+'="Id not in AWS",Sheet1!D'+str(i)+',"")'
            frm2 = '=if(Sheet1!I'+str(i)+'="Id not in AWS","Id not in AWS","")'
            frm3 = '=if(Sheet1!I'+str(i)+'="Id not in AWS","Id not in AWS","")'
            
            sheet_obj1[ind1]=frm1
            sheet_obj1[ind2]=frm2
            sheet_obj1[ind3]=frm3
            i=i+1

        obj.save(path)
        
            
    except Exception as e:
        print(e)
        print (Fore.RED + "Error : The file does not found")
        return ("An Error has occured, pls verify")
    print(Fore.GREEN + "###################### Successfully the excel file has been read/written. ##############################")
    return("Successfully the excel file has been read/written.")