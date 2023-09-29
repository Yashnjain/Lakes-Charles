import os
import re
import time
import glob
import pandas as pd
import logging
import bu_alerts
import xlwings as xw
from tabula import read_pdf
import xlwings.constants as win32c
from datetime import date, datetime, timedelta


drive = r"J:\India"

def remove_existing_files(files_location):
    """_summary_

    Args:
        files_location (_type_): _description_

    Raises:
        e: _description_
    """           
    logging.info("Inside remove_existing_files function")
    try:
        files = os.listdir(files_location)
        if len(files) > 0:
            for file in files:
                os.remove(files_location + "\\" + file)
            logging.info("Existing files removed successfully")
        else:
            print("No existing files available to reomve")
        print("Pause")
    except Exception as e:
        logging.exception("Exception in: remove_existing_files()")
        logging.exception(e)
        raise e


def xlOpner(inputFile):
    try:
        retry = 0
        while retry<10:
            try:
                input_wb = xw.Book(inputFile, update_links=False)
                return input_wb
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
    except Exception as e:
        print(f"Exception caught in xlOpner :{e}")
        logging.info(f"Exception caught in xlOpner :{e}")
        raise e
    

def remove_special_characters(my_pdf,column_list):
    try:
        # column_list = list(my_pdf.columns[[-5,-4,-3,-2]])
        logging.info("inside remove special characters")
        for values in column_list:
            my_pdf[values] = my_pdf[values].astype(str)
            my_pdf[values]  = [x[values].replace('$', '') for i, x in my_pdf.iterrows()]
            my_pdf[values]  = [x[values].replace('(', '-') for i, x in my_pdf.iterrows()]
            my_pdf[values]  = [x[values].replace(')', '') for i, x in my_pdf.iterrows()]
            my_pdf[values]  = [x[values].replace(',', '') for i, x in my_pdf.iterrows()]
            # my_pdf[values]  = [x[values].replace('0.0', '0.00') for i, x in my_pdf.iterrows()]
            my_pdf[values] = my_pdf[values].astype(float)
            # my_pdf[values]  = [x[values].replace('0.00', '0') for i, x in my_pdf.iterrows()]
            
        return  my_pdf   
    except Exception as e:
        raise e  


def insert_top1_btm2_borders(cellrange:str,working_sheet,working_workbook):
    try:
        working_sheet.activate()
        working_sheet.api.Range(cellrange).Select()
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalDown).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalUp).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeLeft).LineStyle = win32c.Constants.xlNone
        # linestylevalues=[win32c.BordersIndex.xlEdgeLeft,win32c.BordersIndex.xlEdgeTop,win32c.BordersIndex.xlEdgeBottom,win32c.BordersIndex.xlEdgeRight,win32c.BordersIndex.xlInsideVertical,win32c.BordersIndex.xlInsideHorizontal]
        # for values in linestylevalues:
        a=working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeTop)
        a.LineStyle = win32c.LineStyle.xlContinuous
        a.ColorIndex = 0
        a.TintAndShade = 0
        a.Weight = win32c.BorderWeight.xlThin
        b=working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeBottom)
        b.LineStyle = win32c.LineStyle.xlDouble
        b.ColorIndex = 0
        b.TintAndShade = 0
        b.Weight = win32c.BorderWeight.xlThick
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeRight).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlInsideVertical).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlInsideHorizontal).LineStyle = win32c.Constants.xlNone
    except Exception as e:
        raise e
    

def working(inventory_wb,sales_wb):
     try:
        sales_ws = sales_wb.sheets("Sheet1")
        inv_working_ws = inventory_wb.sheets("Working")
        last_row = sales_ws.range(f'B'+ str(sales_ws.cells.last_cell.row)).end('up').row
        curr_col_list = sales_ws.range("B6").expand('right').value
        sales_terminal = curr_col_list.index('Terminal')
        sales_ws.api.AutoFilterMode= False
        sales_ws.api.Range(f"B6:AR{last_row}").AutoFilter(Field:=f"{sales_terminal+1}", Criteria1="LAKE CHARLES, LA - TANK")
        inv_working_ws.range('A1').expand('down').clear()
        inv_working_ws.range('B1').expand('down').clear()
        inv_working_ws.range('C1').expand('down').clear()
        inv_working_ws.range('D1').expand('down').clear()
        inv_working_ws.range('E1').expand('down').clear()
        inv_working_ws.api.Range("F:F").EntireColumn.Clear()
        inv_working_ws.range('G1').expand('down').clear()
        inv_working_ws.range('H1').expand('down').clear()
        ########## Particulars #########################
        sales_ws.range(f'B6:B{last_row}').copy()
        inv_working_ws.range('A1').paste()
        ########## Date ################################
        sales_ws.range(f'E6:E{last_row}').copy()
        inv_working_ws.range('B1').paste()
        ########## Customer Name #######################
        sales_ws.range(f'I6:I{last_row}').copy()
        inv_working_ws.range('C1').paste()
        ########## BOL Number ##########################
        sales_ws.range(f'L6:L{last_row}').copy()
        inv_working_ws.range('D1').paste()
        ########## BOL Date ############################
        sales_ws.range(f'M6:M{last_row}').copy()
        inv_working_ws.range('E1').paste()
        ########## Billed QTY ##########################
        sales_ws.range(f"AD6:AD{last_row}").copy()
        inv_working_ws.range('F1').paste()
        sales_ws.api.AutoFilterMode= False
        ########## applying formul for G and H column ############  
        lst_rw = inv_working_ws.range(f'A'+ str(inv_working_ws.cells.last_cell.row)).end('up').row
        inv_working_ws.range(f"G2:G{lst_rw}").api.Formula = f"=+VLOOKUP(D2,Outbound!B:E,4,0)"
        inv_working_ws.range(f"H2:H{lst_rw}").api.Formula ="=+G2+F2"
        working_total_rw = lst_rw+2
        inv_working_ws.range(f"F{lst_rw+2}").api.Formula =f"=SUBTOTAL(9,F2:F{lst_rw})"
        insert_top1_btm2_borders(cellrange=f"F{lst_rw+2}",working_sheet=inv_working_ws,working_workbook=inventory_wb)
        return working_total_rw
     except Exception as e:
          raise e
     

def mrn(inventory_wb,mrn_wb):
     try:
        inventory_wb.activate()  
        inventory_mrndetail_ws = inventory_wb.sheets["MRN Detail"]
        inventory_mrndetail_ws.range("A1:AR1").expand('down').delete()
        mrn_ws = mrn_wb.sheets[0]
        last_row = mrn_ws.range(f'B'+ str(mrn_ws.cells.last_cell.row)).end('up').row
        curr_col_list = mrn_ws.range("B6").expand('right').value
        arrival_date_col = curr_col_list.index('Arrival Date')
        prev = date.today().replace(day=1) - timedelta(days=1)
        mrn_ws.api.Range(f"B6:AR{last_row}").AutoFilter(Field:=arrival_date_col+1,Criteria1:=f">={prev}",Operator:=2,Criteria2:=f"=")
        terminal_col_no = curr_col_list.index('Terminal')
        mrn_ws.api.Range(f"B6:AR{last_row}").AutoFilter(Field:=f"{terminal_col_no+1}", Criteria1="LAKE CHARLES, LA - TANK")
        # sp_address = row_range_calc("B",mrn_ws)
        # mrn_ws.range(f"{sp_address}").copy()
        mrn_ws.range(f"B6:AR{last_row}").copy()
        inventory_mrndetail_ws.range("B1").paste()
        inventory_mrndetail_ws.range("I1").expand('down').copy()
        inventory_mrndetail_ws.range("A1").paste()
     except Exception as e:
          raise e


    
def in_out_inv(inv_path,inventory_wb,working_total_rw,check_sheet):
    try:  
        outbound_inv =  inventory_wb.sheets['Outbound']  
        try:
            if len(glob.glob(inv_path+"\\BioUrja Outbound Tickets*.xlsx"))>0:   
                for file in glob.glob(inv_path+"\\BioUrja Outbound Tickets*.xlsx"):
                    path, file_name = os.path.split(file)
                    inbound_file_name = file_name
                    try:
                        outbound_wb = xlOpner(file)
                    except Exception as e:
                        logging.info(f"could not open workbook: {file}")
                        raise e  
                    
                    outbound_sheet = outbound_wb.sheets[f'{today_date.strftime("%B")}']
                    outbound_sheet.activate()


                    bol_row = outbound_inv.range(f'B'+ str(outbound_inv.cells.last_cell.row)).end('up').row
                    if bol_row!=1:
                        bol_no = int(outbound_inv.range(f"B{bol_row}").value)
                        # bol_no = "3184"
                        outbound_sheet.api.Range("B1").Select()
                        try:
                            outbound_sheet.range('B:B').api.Cells.Find(What:=bol_no, After:=outbound_sheet.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,
                                                LookAt:=win32c.LookAt.xlWhole, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
                        except:
                            logging.info(f"file is not updated {path}:::{file_name}")
                        first_row = outbound_sheet.api.Application.ActiveCell.Row+1
                        end_row = outbound_sheet.range(f"B{int(first_row)}").expand('down').last_cell.row 

                        paste_row = outbound_inv.range(f'B'+ str(outbound_inv.cells.last_cell.row)).end('up').row

                        if bol_row==paste_row:
                            print("Entries already done")
                            break                

                    else:
                        print("first day of the month")
                        first_row = 5
                        end_row = outbound_sheet.range(f'B'+ str(outbound_sheet.cells.last_cell.row)).end('up').row
                        paste_row = 1

                    outbound_df= outbound_sheet.range(f"A{first_row}:J{end_row}").options(pd.DataFrame,header=False,index=False).value
                    req_data = outbound_df[[0,1,9]]
                    req_data["temp1"]= ""
                    req_data["temp2"]= ""
                    req_data = req_data.reindex(columns = [0,1,'temp1','temp2',9])

                    outbound_inv.range(f"A{paste_row+1}").options(header=False,index=False).value = req_data
        

                    outbound_inv.range(f"F{paste_row+1}").value = f"=+VLOOKUP(B{paste_row+1},Working!D:F,3,0)"
                    outbound_inv.range(f"G{paste_row+1}").value = f"=+F{paste_row+1}+E{paste_row+1}"

                    outbound_inv.range(f"F{paste_row+1}:G{paste_row+1}").copy(outbound_inv.range(f"F{paste_row+1}:G{paste_row+len(req_data)}"))
                    outbound_inv.range(f"C{paste_row+1}").value = f"=XLOOKUP(B{paste_row+1},Working!D:D,Working!C:C,0)"

                    outbound_inv.range(f"C{paste_row+1}").copy(outbound_inv.range(f"C{paste_row+1}:C{paste_row+len(req_data)}"))
                    
                    print(f"Outbound Completed")
            else:
                logging.info(f"No outbound reports found, please check ::: {inv_path}")
        except Exception as e:
            logging.exception(f"Check {path}:::::{file_name}")
            logging.exception(str(e))
            print("Error while generating outbound sheet")
            raise e
        
        check_row_E= outbound_inv.range(f'E'+ str(outbound_inv.cells.last_cell.row)).end('up').row
        check_row_A = outbound_inv.range(f'A'+ str(outbound_inv.cells.last_cell.row)).end('up').row
        if check_row_E>check_row_A:
            outbound_inv.range(f"A{check_row_A + 1}:G{check_row_E}").clear_contents()

        outbound_total_rw = check_row_A + 2
        outbound_inv.range(f"E{outbound_total_rw}").value = f"=SUM(E2:E{check_row_A})"
        insert_top1_btm2_borders(cellrange=f"E{outbound_total_rw}",working_sheet=outbound_inv,working_workbook=inventory_wb)
        outbound_inv.range(f"E{check_row_A + 3}").value = f"=Working!F{working_total_rw}"
        
        try:
            if len(glob.glob(inv_path+"\\BioUrja Daily*.xlsx"))>0:
                inbound_inv = inventory_wb.sheets['Inbound']
                ####################### template sheet ###########
                try:
                    lc_wb = xlOpner(check_sheet)
                except Exception as e:
                        logging.info(f"could not open workbook: {check_sheet}")
                        raise e  
                barge_inv = lc_wb.sheets[0]
                check = None
                ####################### barge no sheet ############
                for file in glob.glob(inv_path+"\\BioUrja Daily*.xlsx"):
                    path, file_name = os.path.split(file)
                    outbound_file_name = file_name
                    try:
                        inbound_wb = xlOpner(file)
                    except Exception as e:
                        logging.info(f"could not open workbook: {file}")
                        raise e                                 
                    print(f"inbound started")

                    inbound_sheet = inbound_wb.sheets[-1]
                    inbound_sheet.activate()

                    inbound_inv.range(f"A2:F200").clear_contents()

                    if inbound_sheet.range(f"J13").value==None:
                        logging.info(f"no value found in {inbound_sheet} to update {inbound_inv}")
                        print(f"please update the inventory to get results {inbound_sheet}")
                        break
                    else:
                        print(f"Found values in {inbound_sheet}")
                        gln_lst_rw =inbound_sheet.range(f"J12").end('down').row
                        initial_rw = 13
                        if gln_lst_rw ==initial_rw:
                            print("only one entry found")
                            inbound_sheet.range(f"J{initial_rw}").copy()
                            inbound_inv.api.Range(f"B2")._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)
                            inbound_sheet.range(f"D{initial_rw}").copy(inbound_inv.range(f"C2"))
                            check =True
                        else:
                            inbound_sheet.range(f"J{initial_rw}:J{gln_lst_rw}").copy()
                            inbound_inv.api.Range(f"B2")._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)
                            inbound_sheet.range(f"D{initial_rw}:D{gln_lst_rw}").copy(inbound_inv.range(f"C2")) 


                #######mataching process###########
                end_row = barge_inv.range(f'C'+ str(barge_inv.cells.last_cell.row)).end('up').row
                end_row_main = inbound_inv.range(f'C'+ str(inbound_inv.cells.last_cell.row)).end('up').row
                df = inbound_inv.range(f"A1:C{end_row_main}").options(pd.DataFrame,index=False).value
                df2  = barge_inv.range(f"A1:C{end_row}").options(pd.DataFrame,index=False).value
                merged_df = df.merge(df2, on=['Net Gallons inv report', 'Entry Date'], how='left')
                df['Rail Car/Truck #'] = merged_df['Rail Car/Truck #1']

                ###################################
                inbound_inv.range(f"A2").options(headers=False,index=False,transpose=True).value = df['Rail Car/Truck #'].values
                inbound_inv.range(f"F:F").number_format = 'General'
                inbound_inv.range(f"F2").value = f"=VLOOKUP(A2,'MRN Detail'!A:B,2)"
                inbound_inv.range(f"D:D").number_format = "0.00"
                inbound_inv.range(f"D2").value = f"=VLOOKUP(A2,'MRN Detail'!A:AF,31)"
                inbound_inv.range(f"E:E").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                inbound_inv.range(f"E2").value = f"=B2-D2"
                if not check:
                    inbound_inv.range(f"D2:F2").copy(inbound_inv.range(f"D2:F{end_row_main}"))

                inbound_total_rw = end_row_main + 2
                inbound_inv.range(f"B{inbound_total_rw}").value = f"=SUM(B2:B{end_row_main})"
                insert_top1_btm2_borders(cellrange=f"B{inbound_total_rw}",working_sheet=inbound_inv,working_workbook=inventory_wb)  
                inbound_inv.range(f"E{inbound_total_rw}").value = f"=SUM(E2:E{end_row_main})"
                insert_top1_btm2_borders(cellrange=f"E{inbound_total_rw}",working_sheet=inbound_inv,working_workbook=inventory_wb) 
            else:
                logging.info(f"No inbound reports found, please check ::: {inv_path}")
        except Exception as e:
            logging.exception(str(e))
            logging.exception(f"Check {path}:::::{file_name}")
            print("Error while generating outbound sheet")
            raise e
        
        ############### updating summary tab ##################
        summary_inv = inventory_wb.sheets['Summary']
        summary_inv.activate()
        summary_inv.range(f"B3").value =f"=+Inbound!B{inbound_total_rw}"
        summary_inv.range(f"B4").value =f"=-Outbound!E{outbound_total_rw}"
        
        op_bal_df = pd.read_excel(pre_month_sheet,sheet_name="Summary",header=None)

        if op_bal_df.iloc[:, 0].str.contains('Ending Inventory').any():
            id_index = op_bal_df.iloc[:, 0].str.contains('Ending Inventory').tolist().index(True)
            previous_ending_bal = op_bal_df.iloc[:, 1][id_index]
        else:
            print('could not find ::: "Ending Inventory" in the sheet')
            logging.info('please check last sheet if it has :::: "Ending Inventory" in the sheet in first column')

        summary_inv.range(f"B2").value =previous_ending_bal
        return inbound_file_name,outbound_file_name
    except Exception as e:
        raise e
    finally:
        try:
            inventory_wb.app.quit()
        except:
            pass


if __name__ == "__main__":
    try:

        job_name="Lakes_Charles_INV_AUTOMATION_P2"
        # log progress --
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
        # logfile = os.getcwd() +"\\logs\\"+'Enverus_Logfile'+str(today_date)+'.txt'

        logfile = os.getcwd() + '\\' + 'logs' + '\\' + f'{job_name}.txt'

        logging.basicConfig(
            level=logging.INFO, 
            format='%(asctime)s [%(levelname)s] - %(message)s',
            filename=logfile)

        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        logging.info("Execution Started")

        locations_list = []
        # logging.info('setting paTH TO download')
        receiver_email = "yashn.jain@biourja.com"
        # receiver_email = "yashn.jain@biourja.com,imam.khan@biourja.com,apoorva.kansara@biourja.com, accounts@biourja.com, rini.gohil@biourja.com,itdevsupport@biourja.com"


        time_start=time.time()
        today_date=date.today()
        inv_path = r'J:\India\Inv Rpt\IT_INVENTORY\flows\Lake Charles'
        if len(glob.glob(inv_path+"\\BioUrja Daily*.xlsx"))>0:
            inv_file = glob.glob(inv_path+"\\BioUrja Daily*.xlsx")[0]
            pathinv, file_name_inv = os.path.split(inv_file)
            year = int(re.findall("\d+",file_name_inv)[1])
            pre_month = int(re.findall("\d+",file_name_inv)[0]) - 1
            pre_date = today_date.replace(month=pre_month)
            today_date = today_date.replace(month=int(re.findall("\d+",file_name_inv)[0]))
            pre_date_fldr = pre_date.strftime("%m-%y")
            date_fldr = today_date.strftime("%m-%y")
        else:
            logging.info(f"inventort report not found ::: {inv_path}")   
            locations_list.append(logfile)
            bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{job_name} completed successfully,Inventory file not found here ::: {inv_path}',multiple_attachment_list = locations_list)
                 

        check_sheet = drive+r'\Inv Rpt\IT_INVENTORY\Input\Templates_IT\Lakes Charles'+f'\\LCTemplate.xlsx'


        if not os.path.exists(check_sheet):
            logging.info(f"{check_sheet} Excel file not present")

        inventory_sheet = drive+rf'\{year}\{date_fldr}'+f'\\Lake Charles Tank.xlsx'
        if not os.path.exists(inventory_sheet):
            logging.info(f"{inventory_sheet} Excel file not present")           

        mrn_sheet = drive+rf'\{year}\{date_fldr}'+f'\\MRN.xlsx'
        if not os.path.exists(mrn_sheet):
            logging.info(f"{mrn_sheet} Excel file not present")

        sales_sheet = drive+rf'\{year}\{date_fldr}'+f'\\Sales.xlsx'
        if not os.path.exists(sales_sheet):
            logging.info(f"{sales_sheet} Excel file not present")

        pre_month_sheet = drive+rf'\{year}\{pre_date_fldr}\Transfered'+f'\\Lake Charles Tank.xlsx'
        if not os.path.exists(pre_month_sheet):
            logging.info(f"{pre_month_sheet} Excel file not present")


        try:
            inventory_wb = xlOpner(inventory_sheet)
        except Exception as e:
            logging.info(f"could not open workbook: {inventory_sheet}")
            raise e
        
        try:
            mrn_wb = xlOpner(mrn_sheet)
        except Exception as e:
            logging.info(f"could not open workbook: {mrn_sheet}")
            raise e   

        try:
            sales_wb = xlOpner(sales_sheet)
        except Exception as e:
            logging.info(f"could not open workbook: {sales_sheet}")
            raise e 
            
        inventory_wb.api.AutoFilterMode=False
        inventory_wb.app.api.CutCopyMode=False
        sales_wb.api.AutoFilterMode=False
        sales_wb.app.api.CutCopyMode=False 

        try:
            working_total_rw = working(inventory_wb,sales_wb)
        except Exception as e:
            logging.info(f"Sales Tab Failure : {e}")
            raise e  

        sales_wb.api.AutoFilterMode=False
        sales_wb.app.api.CutCopyMode=False         

        inventory_wb.api.AutoFilterMode=False
        inventory_wb.app.api.CutCopyMode=False
        mrn_wb.api.AutoFilterMode=False
        mrn_wb.app.api.CutCopyMode=False  

        try:
            mrn(inventory_wb,mrn_wb)
        except Exception as e:
            logging.info(f"Mrn Tab Failure : {e}")
            raise e 
        inventory_wb.api.AutoFilterMode=False
        inventory_wb.app.api.CutCopyMode=False
        mrn_wb.api.AutoFilterMode=False
        mrn_wb.app.api.CutCopyMode=False        
        print("sales and mrn done")

        try:
            inbound_file_name,outbound_file_name = in_out_inv(inv_path,inventory_wb,working_total_rw,check_sheet)
        except Exception as e:
            logging.info(f"Inbound/Outbound Tab Failure : {e}")
            raise e        
        print("Done")
        
        output_location = rf'J:\India\Inv Rpt\IT_INVENTORY\Output\{year}\{date_fldr}\Lakes Charles'
        if not os.path.exists(output_location):
            os.makedirs(output_location)

        try:
            inventory_wb.save(f"{output_location}\\Lake Charles Tank.xlsx")
            print(f"inventory done and saved in {output_location}")
            inventory_wb.app.kill()
        except Exception as e:
            logging.info(f"could not save or kill ::: {output_location}")
            raise e 




        remove_existing_files(inv_path)
        logging.info(f"files succesfully removed from folder :::: {inv_path}")
        locations_list.append(logfile)
        locations_list.append(f"{output_location}\\Lake Charles Tank.xlsx")
        nl = '<br>'
        body = ''
        body = (f'{nl}<strong>{inventory_wb.name}</strong> {nl}{nl} <strong>{inventory_wb.name}</strong> successfully created from reports <strong>{inbound_file_name},{outbound_file_name}</strong>, {nl} Attached path for the excel=<u>{output_location}</u>{nl}')
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{body}{job_name} completed successfully, Attached Logs and Excel',multiple_attachment_list = locations_list)
        logging.info("Process completed")
        print("process completed")

    except Exception as e:
        logging.exception(str(e))
        try:
            inventory_wb.app.kill()
        except:
            pass    
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed in __main__, Attached logs',attachment_location = logfile)

