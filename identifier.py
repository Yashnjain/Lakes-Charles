import os
import sys
import time
import glob
import PyPDF2
import logging
import bu_alerts
import xlwings as xw
from tabula import read_pdf
from datetime import date, datetime


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
    
def bank_recons_rep(pdf_path,input_sheet):
    try:
        file_names = []
        if len(glob.glob(pdf_path+"\\*.pdf"))>0:
            for file in glob.glob(pdf_path+"\\*.pdf"):
                path, file_name = os.path.split(file)

                barge_area=["2.678,1.53,109.013,358.02"]
                df=read_pdf(file,stream=True, multiple_tables=True,pages=1,area=barge_area,silent=True,guess=False)
                barge_value=df[0].columns[1]
                # Required_date=text_value[text_value.find("BARGE"):].split()[1]         

                with open(file, 'rb') as f:
                            pdf = PyPDF2.PdfFileReader(f)
                            number_of_pages = pdf.getNumPages()
                            print(number_of_pages)

                for i in range(1,number_of_pages+1):
                    data_area=["3.44,2.295,785.27,607.41"]
                    random_df=read_pdf(file,stream=True, multiple_tables=False,pages=i,area=data_area,silent=True,guess=False)[0]
                    if random_df.iloc[:,0].str.contains('Certificate of Quantity').any():
                        print(f"page found {i}")
                        required_table =read_pdf(file,stream=True, multiple_tables=True,pages=i)[0]
                        if required_table.iloc[:,0].str.contains('US Gallons').any():
                            index_value = required_table.iloc[:,0].str.contains('US Gallons').tolist().index(True)
                            gallons_value = required_table.iloc[:,1][index_value]
                            gallons_value = gallons_value.replace(",","")
                            page_value = i
                            break
                    else:
                        print(f"Certificate of Quantity not found on page {i}")

                if page_value>3:       
                    date_area=["366.045,122.4,449.43,514.845"]
                    date_df=read_pdf(file,stream=True, multiple_tables=True,pages=1,area=date_area,silent=True,guess=False)
                    date_value=date_df[0].values[0][0]

                else:
                    date_area=["424.193,11.475,496.103,608.175"]
                    date_df=read_pdf(file,stream=True, multiple_tables=True,pages=1,area=date_area,silent=True,guess=False)
                    date_value=date_df[0].values[0][0]   

                # job_name = "BANK_RECONS_Automation"
                retry=0
                global wb
                while retry < 10:
                    try:
                        wb=xw.Book(input_sheet)
                        break
                    except Exception as e:
                        time.sleep(5)
                        retry+=1
                        if retry ==10:
                            raise e

                ws1=wb.sheets[0]
                row_to_insert = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row 
                barge_value = barge_value.split("-")[1].strip()
                barge_value = "Barge # " + barge_value
                try:
                    ws1.range(f"A{row_to_insert+1}").value=barge_value.split("(")[0].strip()
                except:
                    ws1.range(f"A{row_to_insert+1}").value=barge_value
                ws1.range(f"B{row_to_insert+1}").value=gallons_value
                try:
                    entry_date=datetime.strptime(date_value,"%m/%d/%y")
                except:  
                    entry_date=datetime.strptime(date_value,"%m/%d/%Y")  
                ws1.range(f"C{row_to_insert+1}").value=entry_date  
                file_names.append(file_name)  
                locations_list.append(file) 
            wb.save()
            print(f"Report Generated")
        else:
            logging.info(f"files not found in {pdf_path}")
            locations_list.append(logfile)
            bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name} with 0 discharge reports',mail_body = f'{job_name} completed successfully, <strong>No Reports found in ::: {pdf_path}</strong>',multiple_attachment_list = locations_list)
            sys.exit(0)           
        return file_names
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass

if __name__ == "__main__":
    try:
        logging.info("Execution Started")
        time_start=time.time()
        today_date=date.today()
        job_name="Lakes_Charles_INV_AUTOMATION_P1"
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
        locations_list = []
        # logging.info('setting paTH TO download')
        # receiver_email = "yashn.jain@biourja.com"
        receiver_email = "yashn.jain@biourja.com,imam.khan@biourja.com,apoorva.kansara@biourja.com, accounts@biourja.com, rini.gohil@biourja.com,itdevsupport@biourja.com"
        input_sheet = drive+r'\Inv Rpt\IT_INVENTORY\Input\Templates_IT\Lakes Charles'+f'\\LCTemplate.xlsx'
        if not os.path.exists(input_sheet):
            logging.info(f"{input_sheet} Excel file not present")
        
        pdf_path = r'J:\India\Inv Rpt\IT_INVENTORY\flows\Lake charles discharge reports'
        if not os.path.exists(pdf_path):
                logging.info(f"{pdf_path} Excel file not present")     

        try:
            file_names = bank_recons_rep(pdf_path,input_sheet)
        except Exception as e:
            logging.info("failed at processing and formatting pdf")
            raise e        
        print("Done")
        
            
        remove_existing_files(pdf_path)
        logging.info(f"files succesfully removed from folder :::: {pdf_path}")
        locations_list.append(logfile)
        locations_list.append(input_sheet)
        nl = '<br>'
        body = ''
        body = (f'{nl}<strong>LCTemplate.xlsx</strong> {nl}{nl} <strong>LCTemplate.xlsx</strong> successfully created from discharge reports <strong>{file_names}</strong>, {nl} Attached path for the excel=<u>{input_sheet}</u>{nl}')
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{body}{job_name} completed successfully, Attached Logs and Excel',multiple_attachment_list = locations_list)
        logging.info("Process completed")
        print("process completed")

    except Exception as e:
        logging.exception(str(e))
        try:
            wb.app.kill()
        except:
            pass    
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed in __main__, Attached logs',attachment_location = logfile)

