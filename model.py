from matplotlib.pyplot import text
from pymysql import NULL
import xlsxwriter
from datetime import date;
from colorama import Fore
import openpyxl
import pandas as pd
from openpyxl.styles.fills import PatternFill
from openpyxl.styles import Font, colors


def model(file1,attribute):
    try:
        print(Fore.RESET)
        path = './excel files/'+file1
        df = pd.read_excel(path)
        n = df.count()[2]+2
        m = df.count()[0]+2
        p = df.count()[5]+3
        obj = openpyxl.load_workbook(path.strip())
        today = date.today()

        ws1 = obj.create_sheet("Sheet2")
        ws1.title= "output_"+str(today)+"_model"
        
        sheet_obj = obj["Sheet1"]
        sheet_obj1 = obj["output_"+str(today)+"_model"]

        
        sheet_obj['H1']='Current_model'
        sheet_obj['I1']='TRIM_id'
        sheet_obj['J1']='r_model wrt _id'
        sheet_obj['K1']='model wrt _id'
        sheet_obj['L1']='R_Model discrepancy'
        sheet_obj['M1']='Model discrepancy'
        sheet_obj['N1']='Turbine Serial Number'
        sheet_obj['O1']='ID Present in AWS?'

        sheet_obj1['A1']='Id'
        sheet_obj1['B1']=attribute+" (AWS)"
        sheet_obj1['C1']='Turbine Serial Number'
        sheet_obj1['D1']='R_'+attribute+" (RAMP)"
        sheet_obj1['E1']='R_'+attribute+"_Discrepancy"
    
        sheet_obj1['G1']='Id'
        sheet_obj1['H1']=attribute+" (AWS)"
        sheet_obj1['I1']='Turbine Serial Number'
        sheet_obj1['J1']=attribute+' (RAMP)'
        sheet_obj1['K1']=attribute+"_Discrepancy"
        
        for j in range(2,n):
            index='H'+str(j)
            formula = '=VLOOKUP(C'+str(j)+',D:E,2,FALSE)' 
            sheet_obj[index]= formula   
            
        for j in range(2,m):
            index1 = 'I'+str(j)
            index2 = 'J'+str(j)
            index3 = 'K'+str(j)
            
            formula1 = '=TRIM(A'+str(j)+')'
            formula2 = '=if(ISNA(VLOOKUP(I'+str(j)+',F:G,2,FALSE)),"Id not in ramp",if(len(VLOOKUP(I'+str(j)+',F:G,2,FALSE))=0,"",VLOOKUP(I'+str(j)+',F:G,2,FALSE)))'
            formula3 = '=if(ISNA(VLOOKUP(I'+str(j)+',F:H,3,false)),"Id not in ramp",if(len(VLOOKUP(I'+str(j)+',F:H,3,false))=0,"",VLOOKUP(I'+str(j)+',F:H,3,false)))'
            
            sheet_obj[index1]= formula1
            sheet_obj[index2]= formula2
            sheet_obj[index3]= formula3
        for j in range(2,m):
            index='L'+str(j)
            formula = '=IF(J'+str(j)+'="","",if(J'+str(j)+'="Id not in ramp","Id not in ramp",IF(J'+str(j)+'=B'+str(j)+',"matching","no matching")))'
            sheet_obj[index]= formula  
        for j in range(2,m):
            index='M'+str(j)
            formula = '=IF(L'+str(j)+'="matching","",IF(K'+str(j)+'="","",if(K'+str(j)+'="Id not in ramp","Id not in ramp",IF(K'+str(j)+'=B'+str(j)+',"matching","not matching"))))'
            sheet_obj[index]= formula  
        for j in range(2,m):
            ind1 = 'A'+str(j)
            ind2 = 'B'+str(j)
            ind3 = 'C'+str(j)
            ind4 = 'D'+str(j)
            ind5 = 'E'+str(j)
            
            ind6 = 'G'+str(j)
            ind7 = 'H'+str(j)
            ind8 = 'I'+str(j)
            ind9 = 'J'+str(j)
            ind10 = 'K'+str(j)

            frm1 = '=Sheet1!A'+str(j)
            frm2 = '=Sheet1!B'+str(j)
            frm3 = '=Sheet1!A'+str(j)
            frm4 = '=Sheet1!J'+str(j)
            frm5 = '=Sheet1!L'+str(j)
            frm6 = '=Sheet1!A'+str(j)
            frm7 = '=Sheet1!B'+str(j)
            frm8 = '=Sheet1!A'+str(j)
            frm9 = '=Sheet1!K'+str(j)
            frm10 = '=Sheet1!M'+str(j)

            sheet_obj1[ind1] = frm1
            sheet_obj1[ind2] = frm2
            sheet_obj1[ind3] = frm3
            sheet_obj1[ind4] = frm4
            sheet_obj1[ind5] = frm5
            sheet_obj1[ind6] = frm6
            sheet_obj1[ind7] = frm7
            sheet_obj1[ind8] = frm8
            sheet_obj1[ind9] = frm9
            sheet_obj1[ind10] = frm10
            
        for j in range(2,p):
            index1 = 'N'+str(j)
            index2 = 'O'+str(j)

            formula1 = '=TRIM(F'+str(j)+')'
            formula2 = '=if(ISNA(vlookup(N'+str(j)+',A:B,2,false)),"Id not in AWS",if( len(vlookup(N'+str(j)+',A:B,2,false))=0,"",vlookup(N'+str(j)+',A:B,2,false) ))'
        

            sheet_obj[index1]=formula1
            sheet_obj[index2]=formula2
        i=2
        for j in range(m,m+p):
            ind1 = 'C'+str(j)
            ind2 = 'D'+str(j)
            ind3 = 'E'+str(j)
            
            ind4 = 'I'+str(j)
            ind5 = 'J'+str(j)
            ind6 = 'K'+str(j)

            frm1 = '=if(Sheet1!O'+str(i)+'="Id not in AWS",Sheet1!N'+str(i)+',"")'
            frm2 = '=if(Sheet1!O'+str(i)+'="Id not in AWS","","")'
            frm3 = '=if(Sheet1!O'+str(i)+'="Id not in AWS","Id not in AWS","")'
                        
            sheet_obj1[ind1]=frm1
            sheet_obj1[ind2]=frm2
            sheet_obj1[ind3]=frm3

            sheet_obj1[ind4]=frm1
            sheet_obj1[ind5]=frm2
            sheet_obj1[ind6]=frm3

            i=i+1
        
        # #color the rows in sheet
        # redFill = PatternFill(patternType='solid', fgColor=colors.Color(rgb='00FFFF00'))
        # sheet_obj.cell(row=1,column=1).fill = redFill
        # sheet_obj1.cell(row=1,column=1).fill = redFill

        obj.save(path)
        
            
    except Exception as e:
        print(e)
        print (Fore.RED + "Error : The file does not found")
        return ("An Error has occured, pls verify")
    print(Fore.GREEN + "###################### Successfully the excel file has been read/written. ##############################")
    return("Successfully the excel file has been read/written.")