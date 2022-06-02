from matplotlib.pyplot import text
from datetime import date;
from colorama import Fore
import openpyxl
import pandas as pd
from openpyxl.styles import Font

def cod(file1,attribute):
    try:
        print(Fore.RESET)
        path = './excel files/'+file1
        df = pd.read_excel(path)
        m = df.count()[2]+2
        n = df.count()[0]+2
        
        obj = openpyxl.load_workbook(path.strip())
        today = date.today()

        ws1 = obj.create_sheet("Sheet2")
        ws1.title= "output_"+str(today)+"_COD"
        
        sheet_obj = obj["Sheet1"]
        sheet_obj1 = obj["output_"+str(today)+"_COD"]

        sheet_obj.insert_cols(0,amount=1)
        sheet_obj.insert_cols(4,amount=2)
        sheet_obj.insert_cols(8,amount=1)
        sheet_obj.insert_cols(9,amount=1) 

        
        sheet_obj['A1']='TRIM_id'
        sheet_obj['D1']='Appropriate format of '+attribute
        sheet_obj['E1']='Trim_Turbine Serial Number'
        sheet_obj['H1']='R_'+attribute+' wrt id'
        sheet_obj['K1']=attribute+' wrt id'
        sheet_obj['I1']='Appropriate format of R_'+attribute+' wrt id'
        sheet_obj['L1']='Appropriate format of '+attribute+' wrt id'
        sheet_obj['M1']='R_'+attribute+'_Discrepancy'
        sheet_obj['N1']=attribute+'_Discrepancy'
        sheet_obj['O1']='ID in AWS?'
        
        print("m,n",m,n)
        sheet_obj1['A1']='Id'
        sheet_obj1['B1']=attribute+" (AWS)"
        sheet_obj1['C1']='Turbine Serial Number'
        sheet_obj1['D1']='R_'+attribute+" (RAMP)"
        sheet_obj1['E1']='R_'+attribute+'_Discrepancy'

        sheet_obj1['G1']='Id'
        sheet_obj1['H1']=attribute+" (AWS)"
        sheet_obj1['I1']='Turbine Serial Number'
        sheet_obj1['J1']=attribute+" (RAMP)"
        sheet_obj1['K1']=attribute+'_Discrepancy'
        ##############################FILLING##############################
        f10 = openpyxl.styles.fills.PatternFill(start_color='FFFF00',end_color='FFFF00',fill_type='solid')
        for y in range(1,15+1):
            sheet_obj.cell(row=1,column=y).fill = f10
            sheet_obj.cell(row=1,column=y).font = Font(bold=True)
        for y in range(1,5+1):
            sheet_obj1.cell(row=1,column=y).fill = f10
            sheet_obj1.cell(row=1,column=y).font = Font(bold=True)
        for y in range(7,11+1):
            sheet_obj1.cell(row=1,column=y).fill = f10
            sheet_obj1.cell(row=1,column=y).font = Font(bold=True)
            
        for j in range(2,n):
            index1 = 'A'+str(j)
            index2 = 'D'+str(j)
            
            formula1 = '=TRIM(B'+str(j)+')'
            formula2 = '=IF(C'+str(j)+'="NULL","",IF(C'+str(j)+'="","",IF(ISERROR(TEXT(DATE(YEAR(C'+str(j)+'),MONTH(C'+str(j)+'),DAY(C'+str(j)+')),"m/d/yyyy")),C'+str(j)+',TEXT(DATE(YEAR(C'+str(j)+'),MONTH(C'+str(j)+'),DAY(C'+str(j)+')),"m/d/yyyy"))))'
            # =if(C'+str(j)+'="NULL","NULL",if(C'+str(j)+'="","",text(date(year(C'+str(j)+'),month(C'+str(j)+'),day(C'+str(j)+')),"m/d/yyyy")))'
            
            sheet_obj[index1]= formula1
            sheet_obj[index2]= formula2

        for j in range(2,m):
            index1 = 'E'+str(j)
            index2 = 'O'+str(j)

            formula1 = '=TRIM(F'+str(j)+')'
            formula2 = '=if(ISNA(vlookup(E'+str(j)+',A:D,4,false)),"RAMP id not in AWS",if( len(vlookup(E'+str(j)+',A:D,4,false))=0,"",vlookup(E'+str(j)+',A:D,4,false) ))'

            sheet_obj[index1]=formula1
            sheet_obj[index2]=formula2    
            
        for j in range(2,n):
            index1 = 'H'+str(j)
            index2 = 'K'+str(j)
            index3 = 'I'+str(j)
            index4 = 'L'+str(j)

            formula1 = '=if(ISNA(VLOOKUP(A'+str(j)+',E:G,3,FALSE)),"AWS id not in RAMP",if(len(VLOOKUP(A'+str(j)+',E:G,3,FALSE))=0,"",VLOOKUP(A'+str(j)+',E:G,3,FALSE)))'
            formula2 = '=if(ISNA(VLOOKUP(A'+str(j)+',E:J,6,FALSE)),"AWS id not in RAMP",if(len(VLOOKUP(A'+str(j)+',E:J,6,FALSE))=0,"",VLOOKUP(A'+str(j)+',E:J,6,FALSE)))'
            formula3 = '=if(H'+str(j)+'="NULL","NULL",if(H'+str(j)+'="","",if(H'+str(j)+'="AWS id not in RAMP","AWS id not in RAMP",text(date(year(H'+str(j)+'),month(H'+str(j)+'),day(H'+str(j)+')),"m/d/yyyy"))))'
            formula4 = '=if(K'+str(j)+'="NULL","NULL",if(K'+str(j)+'="","",if(K'+str(j)+'="AWS id not in RAMP","AWS id not in RAMP",text(date(year(K'+str(j)+'),month(K'+str(j)+'),day(K'+str(j)+')),"m/d/yyyy"))))'

            sheet_obj[index1]=formula1
            sheet_obj[index2]=formula2  
            sheet_obj[index3]=formula3
            sheet_obj[index4]=formula4

        for j in range(2,n):
            index1 = 'M'+str(j)
            index2 = 'N'+str(j)
            formula1 = '=if(I'+str(j)+'="AWS id not in RAMP","AWS id not in RAMP",if(AND(OR(D'+str(j)+'="",D'+str(j)+'="NULL"),OR(I'+str(j)+'="",I'+str(j)+'="NULL")),"matching",if(I'+str(j)+'="","",if(D'+str(j)+'=I'+str(j)+',"matching","not matching"))))'
            formula2 = '=if(L'+str(j)+'="AWS id not in RAMP","AWS id not in RAMP",if(M'+str(j)+'="matching","",if(AND(OR(D'+str(j)+'="",D'+str(j)+'="NULL"),OR(L'+str(j)+'="",L'+str(j)+'="NULL")),"matching",if(L'+str(j)+'="","",if(D'+str(j)+'=L'+str(j)+',"matching","not matching")))))'
            sheet_obj[index1]=formula1
            sheet_obj[index2]=formula2 
              
        for j in range(2,n):
            ind1 = 'A'+str(j)
            ind2 = 'B'+str(j)
            ind3 = 'C'+str(j)
            ind4 = 'D'+str(j)
            ind5 = 'E'+str(j)      

            frm1 = '=Sheet1!B'+str(j)
            frm2 = '=Sheet1!D'+str(j)
            frm3 = '=Sheet1!B'+str(j)
            frm4 = '=Sheet1!I'+str(j)
            frm5 = '=Sheet1!M'+str(j)

            sheet_obj1[ind1] = frm1
            sheet_obj1[ind2] = frm2
            sheet_obj1[ind3] = frm3
            sheet_obj1[ind4] = frm4
            sheet_obj1[ind5] = frm5

            ind1 = 'G'+str(j)
            ind2 = 'H'+str(j)
            ind3 = 'I'+str(j)
            ind4 = 'J'+str(j)
            ind5 = 'K'+str(j) 

            frm1 = '=Sheet1!B'+str(j)
            frm2 = '=Sheet1!D'+str(j)
            frm3 = '=Sheet1!B'+str(j)
            frm4 = '=Sheet1!L'+str(j)
            frm5 = '=Sheet1!N'+str(j)            

            sheet_obj1[ind1] = frm1
            sheet_obj1[ind2] = frm2
            sheet_obj1[ind3] = frm3
            sheet_obj1[ind4] = frm4
            sheet_obj1[ind5] = frm5
            
        
        i=2
        for j in range(n,m+n):
            ind1 = 'C'+str(j)
            ind2 = 'D'+str(j)
            ind3 = 'E'+str(j)

            frm1 = '=if(Sheet1!O'+str(i)+'="RAMP id not in AWS",Sheet1!F'+str(i)+',"")'
            frm2 = '=if(Sheet1!O'+str(i)+'="RAMP id not in AWS","RAMP id not in AWS","")'
            frm3 = '=if(Sheet1!O'+str(i)+'="RAMP id not in AWS","RAMP id not in AWS","")'
            
            sheet_obj1[ind1]=frm1
            sheet_obj1[ind2]=frm2
            sheet_obj1[ind3]=frm3

            ind1 = 'I'+str(j)
            ind2 = 'J'+str(j)
            ind3 = 'K'+str(j)
            
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