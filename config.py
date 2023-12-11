from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import pandas as pd
import requests
import os
from datetime import datetime
import pyodbc
from sharepy import SharePointSession
import datetime
from datetime import datetime, timedelta ,date
import re
import openpyxl
import numpy as np
import http.client
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import subprocess

log_path = 'D:/รวมงาน/FS-Data/BI_Audit/NC-BI/'#'C:/Python/FS_CONFIG/CONFIG/Log/'#
# f = open("C:/Python/U_P_sharepoint.txt", "r")
# textlogin = f.readlines(0)
# f.close()
# U_sharepoint = textlogin[0].replace("\n", "")
# P_sharepoint = textlogin[1]

def xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database):
    log_time = str(date.today().strftime("%Y%m%d"))
    run_time = datetime.now()#.strftime("%Y%m%d%H%M%S")
    if os.path.exists(f'{log_path}{lof_filename}_{log_time}.txt'):
        with open(f'{log_path}{lof_filename}_{log_time}.txt', 'a') as f:
            f.write(f'\n({run_time}) Info: Starting task FS_CONFIG')
    else:
        with open(f'{log_path}{lof_filename}_{log_time}.txt', 'w') as f:
            f.write(f'({run_time}) Info: Starting task FS_CONFIG')
    try:
        log_time_name = datetime.now().strftime("%Y%m%d%H%M%S")
        year = str(date.today().strftime("%Y"))
        site_url = "https://vschem365.sharepoint.com/sites/DataTeam"
        username = 'it@vschem.com'#U_sharepoint#
        password = '365Loading'#P_sharepoint#
        sender_email = 'service@vschem.com'
        sender_password = 'n2C4@R6m1#'
        receiver_email = ['tadson.s@vschem.com']#,'niti.k@vschem.com']
        cc_recipients = []

        def round_down(value):
            if isinstance(value, (int, float)):
                return int(value)
            else:
                return value    

        def move(df,key,sharepoint_view,fstype):    
            log_time_name = datetime.now().strftime("%Y%m%d%H%M%S")
            year = str(date.today().strftime("%Y"))
            df.to_csv(f'{log_path}{key}_{fstype}_{log_time_name}.csv', index=False)#df.to_excel(f'C:/Python/FX_CONVERT/log/{key}_{log_time_name}.xlsx', index=False) 
            web = ctx.web
            ctx.load(web)
            ctx.execute_query()
            print('Connected to SharePoint: ',web.properties['Title'])

            fileName = f'{log_path}{key}_{fstype}_{log_time_name}.csv'#f'C:/Python/FX_CONVERT/log/{key}_{log_time_name}.xlsx'

            with open(fileName, 'rb') as content_file:
                file_content = content_file.read()     

            name = os.path.basename(fileName)
            list_title = "Documents"
            target_list = ctx.web.lists.get_by_title(list_title)
            #info = FileCreationInformation()
            destination_folder_path = f"/sites/DataTeam/Shared Documents/Data Center/000-Pool/001-Account/{sharepoint_view}/IMPORT/{year}"
            libraryRoot = ctx.web.get_folder_by_server_relative_url(destination_folder_path)
    #             DeliveryRoot = ctx.web.get_folder_by_server_relative_url(source_folder_path)
            os.remove(f'{log_path}{key}_{fstype}_{log_time_name}.csv')#os.remove(f'C:/Python/FX_CONVERT/log/{key}_{log_time_name}.xlsx')
            target_file = libraryRoot.upload_file(name, file_content).execute_query()
            print("File has been uploaded to url: {0}".format(target_file.serverRelativeUrl))

    #         file_name_to_delete = f"{key}.xlsx"
    #         file_path_to_delete = os.path.join(source_folder_path, file_name_to_delete)

    #         file_to_delete = ctx.web.get_file_by_server_relative_url(file_path_to_delete)
    #         file_to_delete.delete_object().execute_query()

    #         print(f"File '{file_name_to_delete}' has been deleted from the source folder.")

            # Cleanup and execute the batch
            ctx.execute_batch()        

            a = 'move success'
            return a


        def del_move(df,key,view,fstype):    
            log_time_name = datetime.now().strftime("%Y%m%d%H%M%S")
            year = str(date.today().strftime("%Y"))
            df.to_csv(f'Downloads/{key}_{fstype}_{log_time_name}.csv', index=False)#df.to_excel(f'C:/Python/FX_CONVERT/log/{key}_{log_time_name}.xlsx', index=False) 
            web = ctx.web
            ctx.load(web)
            ctx.execute_query()
            print('Connected to SharePoint: ',web.properties['Title'])

            fileName = f'Downloads/{key}_{fstype}_{log_time_name}.csv'#f'C:/Python/FX_CONVERT/log/{key}_{log_time_name}.xlsx'

            with open(fileName, 'rb') as content_file:
                file_content = content_file.read()     

            name = os.path.basename(fileName)
            list_title = "Documents"
            target_list = ctx.web.lists.get_by_title(list_title)
            #info = FileCreationInformation()
            destination_folder_path = f"/sites/DataTeam/Shared Documents/Data Center/000-Pool/001-Account/{view}/IMPORT/{year}"
            libraryRoot = ctx.web.get_folder_by_server_relative_url(destination_folder_path)
    #             DeliveryRoot = ctx.web.get_folder_by_server_relative_url(source_folder_path)
            os.remove(f'Downloads/{key}_{fstype}_{log_time_name}.csv')#os.remove(f'C:/Python/FX_CONVERT/log/{key}_{log_time_name}.xlsx')
            target_file = libraryRoot.upload_file(name, file_content).execute_query()
            print("File has been uploaded to url: {0}".format(target_file.serverRelativeUrl))

            file_name_to_delete = f"{key}.xlsx"
            file_path_to_delete = os.path.join(source_folder_path, file_name_to_delete)

            file_to_delete = ctx.web.get_file_by_server_relative_url(file_path_to_delete)
            file_to_delete.delete_object().execute_query()

            print(f"File '{file_name_to_delete}' has been deleted from the source folder.")

            # Cleanup and execute the batch
            ctx.execute_batch()        

            a = 'delete move success'
            return a
        def logfile(text):
            run_time = datetime.now()
            with open(f'{log_path}{lof_filename}_{log_time}.txt', 'a') as f:
                f.write(f'\n({run_time}) Info: {text}')
                
        auth_context = AuthenticationContext(url=site_url)
        auth_context.acquire_token_for_user(username=username, password=password)
        ctx = ClientContext(site_url, auth_context)    
        source_folder_path = f"/sites/DataTeam/Shared Documents/Data Center/000-Pool/001-Account/{sharepoint_view}"
        file_collection = ctx.web.get_folder_by_server_relative_url(source_folder_path).files
        ctx.load(file_collection)
        ctx.execute_query()

        logfile('Connected to Sharepoint')

        df_all_cur1 = []
        sorted_files = sorted(file_collection, key=lambda file: file.name)
        for file in sorted_files:
            if file.name == xlsx_name:
                print(file.name)
                response = file.read()
                key = f"{file.name[:-5]}"
                with pd.ExcelFile(response) as xls:
                    if 'BS' in xls.sheet_names:
#                         df = pd.read_excel(xls, sheet_name='BS')
                        column_data_types = {
                                                        'SEQID': int,
                                                        'LAYER1': str,
                                                        'LAYER2': str,
                                                        'LAYER3': str,
                                                        'FORMULA': str,
                                                        'GLCODE': str,
                                                        'DIMCODE1': str,
                                                        'DIMCODE2': str,
                                                        'DIMCODE3': str,
                                                        'DIMCODE4': str,
                                                        'DIMCODE5': str,
                                                        'INVERT': str,
                                                        'SHOWDETAIL': str,
                                                        'UOM': str,
                                                        'MODELYEAR': str,
                                                        'COMMENT': str
                                                        }
                        df = pd.read_excel(xls, sheet_name='BS', dtype=column_data_types)
                        df['COMPANY'] =  company
                        df['FSTYPE'] = 'BS'
                        df['FSVIEW'] = FSVIEW
                        move(df,key,sharepoint_view,'BS')
    #                     move(df,key,'FS_Audit_Config','BS')
                if not df.empty:
                    df_all_cur1.append(df)
        AD_BS = []
        if df_all_cur1:
            AD_BS = pd.concat(df_all_cur1, ignore_index=True)
            AD_BS = AD_BS.dropna(subset=['LAYER1'])
            AD_BS['GLCODE'] = AD_BS['GLCODE'].fillna('')
            AD_BS['GLCODE'] = AD_BS['GLCODE'].apply(round_down) 
            fill_values = {'DIMCODE1': '', 'DIMCODE2': '', 'DIMCODE3': '', 'DIMCODE4': '', 'DIMCODE5': '','SHOWDETAIL':0}
            AD_BS.fillna(fill_values, inplace=True)
            AD_BS['DIMCODE1'] = AD_BS['DIMCODE1'].apply(round_down)
            AD_BS['DIMCODE2'] = AD_BS['DIMCODE2'].apply(round_down)
            AD_BS['DIMCODE3'] = AD_BS['DIMCODE3'].apply(round_down)
            AD_BS['DIMCODE4'] = AD_BS['DIMCODE4'].apply(round_down)
            AD_BS['DIMCODE5'] = AD_BS['DIMCODE5'].apply(round_down)  
    #         mask = AD_BS['FORMULA'].str.startswith('=')
    #         AD_BS.loc[mask, 'FORMULA'] = " ' " + AD_BS.loc[mask, 'FORMULA']
            AD_BS['FORMULA'] = AD_BS['FORMULA'].str.replace("'", "\'\'")
            logfile(f'Get {FSVIEW} BS data')
        else:
            pass


        df_all_cur2 = []
        sorted_files = sorted(file_collection, key=lambda file: file.name)
        for file in sorted_files:
            if file.name == xlsx_name:
                print(file.name)
                response = file.read()
                key = f"{file.name[:-5]}"
                with pd.ExcelFile(response) as xls:
                    if 'PL' in xls.sheet_names:
#                         df = pd.read_excel(xls, sheet_name='PL')
                        column_data_types = {
                                                        'SEQID': int,
                                                        'LAYER1': str,
                                                        'LAYER2': str,
                                                        'LAYER3': str,
                                                        'FORMULA': str,
                                                        'GLCODE': str,
                                                        'DIMCODE1': str,
                                                        'DIMCODE2': str,
                                                        'DIMCODE3': str,
                                                        'DIMCODE4': str,
                                                        'DIMCODE5': str,
                                                        'INVERT': str,
                                                        'SHOWDETAIL': str,
                                                        'UOM': str,
                                                        'MODELYEAR': str,
                                                        'COMMENT': str
                                                        }
                        df = pd.read_excel(xls, sheet_name='PL', dtype=column_data_types)
                        df['COMPANY'] =  company
                        df['FSTYPE'] = 'PL'
                        df['FSVIEW'] = FSVIEW
                        move(df,key,sharepoint_view,'PL')
    #             del_move(df,key,'FS_Audit_Config','PL')
                if not df.empty:
                    df_all_cur2.append(df)
        AD_PL = []
        if df_all_cur2:
            AD_PL = pd.concat(df_all_cur2, ignore_index=True)
            AD_PL = AD_PL.dropna(subset=['LAYER1'])
            AD_PL['GLCODE'] = AD_PL['GLCODE'].fillna('')
            AD_PL['GLCODE'] = AD_PL['GLCODE'].apply(round_down)
            fill_values = {'DIMCODE1': '', 'DIMCODE2': '', 'DIMCODE3': '', 'DIMCODE4': '', 'DIMCODE5': '', 'SHOWDETAIL':0}
            AD_PL.fillna(fill_values, inplace=True)
            AD_PL['DIMCODE1'] = AD_PL['DIMCODE1'].apply(round_down)
            AD_PL['DIMCODE2'] = AD_PL['DIMCODE2'].apply(round_down)
            AD_PL['DIMCODE3'] = AD_PL['DIMCODE3'].apply(round_down)
            AD_PL['DIMCODE4'] = AD_PL['DIMCODE4'].apply(round_down)
            AD_PL['DIMCODE5'] = AD_PL['DIMCODE5'].apply(round_down) 
            logfile(f'Get {FSVIEW} PL data')

        final_df = [AD_BS, AD_PL]
        non_empty_tables = []

        for table in final_df:
            if isinstance(table, pd.DataFrame) and not table.empty:
                non_empty_tables.append(table)

        if non_empty_tables:
            concatenated_df = pd.concat(non_empty_tables, ignore_index=True)
            concatenated_df.fillna('', inplace=True)
            concatenated_df = concatenated_df[['COMPANY','FSTYPE','FSVIEW','SEQID','LAYER1','LAYER2','LAYER3','FORMULA','GLCODE','DIMCODE1','DIMCODE2','DIMCODE3','DIMCODE4','DIMCODE5','UOM','INVERT','SHOWDETAIL','MODELYEAR','COMMENT']]
            concatenated_df.fillna('', inplace=True)
            concatenated_df.replace("'", "''", regex=True, inplace=True)
            concatenated_df.loc[concatenated_df['FORMULA'] != 'Member','INVERT'] = 0
            concatenated_df.loc[concatenated_df['INVERT'].astype(str).str.len() == 0,'INVERT'] = 1
            ########################################
            concatenated_df['LAYER2'] = concatenated_df['LAYER2'].apply(lambda x: "" if len(str(x)) == 1 else x)
            concatenated_df['LAYER3'] = concatenated_df['LAYER3'].apply(lambda x: "" if len(str(x)) == 1 else x)
            concatenated_df['GLCODE'] = concatenated_df['GLCODE'].apply(lambda x: "" if len(str(x)) == 1 else x)
            concatenated_df['DIMCODE1'] = concatenated_df['DIMCODE1'].apply(lambda x: "" if len(str(x)) == 1 else x)
            concatenated_df['DIMCODE2'] = concatenated_df['DIMCODE2'].apply(lambda x: "" if len(str(x)) == 1 else x)
            concatenated_df['DIMCODE3'] = concatenated_df['DIMCODE3'].apply(lambda x: "" if len(str(x)) == 1 else x)
            concatenated_df['DIMCODE4'] = concatenated_df['DIMCODE4'].apply(lambda x: "" if len(str(x)) == 1 else x)
            concatenated_df['DIMCODE5'] = concatenated_df['DIMCODE5'].apply(lambda x: "" if len(str(x)) == 1 else x)
            concatenated_df['UOM'] = concatenated_df['UOM'].apply(lambda x: "" if len(str(x)) == 1 and '%' not in str(x) else x)
            # concatenated_df['UOM'].str.replace(' ', '')
            # concatenated_df['UOM'].str.replace('\s', '', regex=True)
            concatenated_df['MODELYEAR'] = concatenated_df['MODELYEAR'].apply(lambda x: "" if len(str(x)) == 1 else x)
            concatenated_df['COMMENT'] = concatenated_df['COMMENT'].apply(lambda x: "" if len(str(x)) == 1 else x)

            ########################################
    #         concatenated_df.loc[concatenated_df['INVERT'].str.len() == 0] = 1
            fill_values = {'INVERT':1}
            concatenated_df.fillna(fill_values, inplace=True)
            a = 'has data'

        else:
            a = 'no data'
            logfile('not have data')

        if a == 'has data':
            conn_str = (
                r'DRIVER={SQL Server};'
                r'SERVER=COWSQL04;'
                f'DATABASE={database};'
                r'UID=addon;'
                r'PWD=addon;'
            )
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()

            df_delete = concatenated_df[['COMPANY','FSTYPE','FSVIEW']]
            df_delete.drop_duplicates(inplace=True)
            df_delete = df_delete.reset_index(drop = True)
            for index, row in df_delete.iterrows():
                COMPANY = row['COMPANY']
                FSTYPE = row['FSTYPE']
                FSVIEW = row['FSVIEW']
                delete_statement = f"DELETE FROM XLS_FSCONFIG WHERE COMPANY = '{COMPANY}' and FSTYPE = '{FSTYPE}' and FSVIEW = '{FSVIEW}'"
                cursor.execute(delete_statement)
                conn.commit()    
            print('delete')
            logfile('Delete old FS_CONFIG')

            inserted_rows = 0
            for index, row in concatenated_df.iterrows():
                COMPANY = row['COMPANY']
                FSTYPE = row['FSTYPE']
                FSVIEW = row['FSVIEW']
                SEQID = row['SEQID']
                LAYER1 = row['LAYER1']
                LAYER2  = row['LAYER2']
                LAYER3 = row['LAYER3']
                FORMULA = row['FORMULA']
                GLCODE = row['GLCODE']
                DIMCODE1 =row['DIMCODE1']
                DIMCODE2 = row['DIMCODE2']
                DIMCODE3 = row['DIMCODE3']
                DIMCODE4 = row['DIMCODE4']
                DIMCODE5 = row['DIMCODE5']
                UOM = row['UOM']
                INVERT = row['INVERT']
                SHOWDETAIL = row['SHOWDETAIL']
                MODELYEAR = row['MODELYEAR']
                COMMENT = row['COMMENT']
                insert_statement = f"INSERT INTO XLS_FSCONFIG (COMPANY,FSTYPE,FSVIEW,SEQID,LAYER1,LAYER2,LAYER3,FORMULA,GLCODE,DIMCODE1,DIMCODE2,DIMCODE3,DIMCODE4,DIMCODE5,UOM,INVERT,SHOWDETAIL,MODELYEAR,COMMENT,SYNCDATE,DESCRIPTION) VALUES ('{COMPANY}','{FSTYPE}','{FSVIEW}',{SEQID},'{LAYER1}','{LAYER2}','{LAYER3}','{FORMULA}','{GLCODE}','{DIMCODE1}','{DIMCODE2}','{DIMCODE3}','{DIMCODE4}','{DIMCODE5}','{UOM}',{INVERT},{SHOWDETAIL},'{MODELYEAR}','{COMMENT}',GETDATE(),'')"
                cursor.execute(insert_statement)
                conn.commit()
                inserted_rows += 1
                print(f"Insert row: {inserted_rows}", end='\r')
            cursor.close()
            conn.close()
            logfile('Insert new FS_CONFIG')
            print('insert')
        today_dates =  date.today() 
        today_date =  str(today_dates.strftime("%Y%m%d"))
        if a == 'has data':
            msg = MIMEMultipart()
            today = date.today()
            msg['Subject'] = f"FS [{company}]: Individual Config File"
            msg['From'] = sender_email
            msg['To'] = ', '.join(receiver_email)
            msg['Cc'] = ', '.join(cc_recipients) 

            df_to_email = concatenated_df.loc[concatenated_df['FORMULA'] != 'Member']#concatenated_df.loc[:999]
            pd.set_option('display.max_rows', 500)
            pd.set_option('display.max_columns', 500)
            pd.set_option('display.width', 1000)


            html = """\
            <html>
              <head>The code ran successfully</head>
              <body>
                {0}
              </body>
            </html>
            """.format(df_to_email.to_html(index=False))

            part1 = MIMEText(html, 'html')
            msg.attach(part1)

            server = smtplib.SMTP('smtp.office365.com', 587)
            server.starttls()
            server.login(sender_email,sender_password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
            server.close()


    except Exception as e:
        print(f'Error: {str(e)}')
        today_dates =  date.today() 
        today_date =  str(today_dates.strftime("%Y%m%d"))
        msg = MIMEMultipart()
        today = date.today()
        msg['Subject'] = f"FS [{company}]: Individual Config File"
        msg['From'] = sender_email
        msg['To'] = ', '.join(receiver_email)
        msg['Cc'] = ', '.join(cc_recipients) 

        html = f'An error occurred: {str(e)}'

    #         msg.attach(MIMEText(message, 'plain'))
        part1 = MIMEText(html, 'plain')
        msg.attach(part1)

        server = smtplib.SMTP('smtp.office365.com', 587)
        #server.ehlo()#NOT NECESSARY
        server.starttls()
        #server.ehlo()#NOT NECESSARY
        server.login(sender_email,sender_password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.close() 
    return
    'process success'
    

try:
    site_url = "https://vschem365.sharepoint.com/sites/DataTeam"
    username = 'it@vschem.com'#U_sharepoint#
    password = '365Loading'#P_sharepoint#
    
    auth_context = AuthenticationContext(url=site_url)
    auth_context.acquire_token_for_user(username=username, password=password)
    ctx = ClientContext(site_url, auth_context)    
    source_folder_path = f"/sites/DataTeam/Shared Documents/Data Center/000-Pool/001-Account/FS_Audit_Config"
    file_collection = ctx.web.get_folder_by_server_relative_url(source_folder_path).files
    ctx.load(file_collection)
    ctx.execute_query()
        
    
    def run_script(path):
        python_script = path
        process = subprocess.Popen(['python', python_script], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, stderr = process.communicate()
        print(f'run {path} success')    
    
    
    def move(df,key):    
        log_time_name = datetime.now().strftime("%Y%m%d%H%M%S")
        year = str(date.today().strftime("%Y"))
        df.to_excel(f'{log_path}{key}.xlsx', index=False)#df.to_excel(f'C:/Python/FX_CONVERT/log/{key}_{log_time_name}.xlsx', index=False) 
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        print('Connected to SharePoint: ',web.properties['Title'])
        
        file_name_to_delete = f"{key}.xlsx"
        file_path_to_delete = os.path.join(source_folder_path, file_name_to_delete)

        file_to_delete = ctx.web.get_file_by_server_relative_url(file_path_to_delete)
        file_to_delete.delete_object().execute_query()

        print(f"File '{file_name_to_delete}' has been deleted from the source folder.")

        fileName = f'{log_path}{key}.xlsx'#f'C:/Python/FX_CONVERT/log/{key}_{log_time_name}.xlsx'

        with open(fileName, 'rb') as content_file:
            file_content = content_file.read()     

        name = os.path.basename(fileName)
        list_title = "Documents"
        target_list = ctx.web.lists.get_by_title(list_title)
        #info = FileCreationInformation()
        destination_folder_path = f"/sites/DataTeam/Shared Documents/Data Center/000-Pool/001-Account/FS_Audit_Config"
        libraryRoot = ctx.web.get_folder_by_server_relative_url(destination_folder_path)
#             DeliveryRoot = ctx.web.get_folder_by_server_relative_url(source_folder_path)
        os.remove(f'{log_path}{key}.xlsx')#os.remove(f'C:/Python/FX_CONVERT/log/{key}_{log_time_name}.xlsx')
        target_file = libraryRoot.upload_file(name, file_content).execute_query()
        print("File has been uploaded to url: {0}".format(target_file.serverRelativeUrl))



        # Cleanup and execute the batch
        ctx.execute_batch()        

        a = 'delete and upload success'
        return a 
    
    df_all_cur1 = []
    sorted_files = sorted(file_collection, key=lambda file: file.name)
    for file in sorted_files:
        if file.name == 'Reload_AD.xlsx':
            print(file.name)
            response = file.read()
            key = f"{file.name[:-5]}"
            with pd.ExcelFile(response) as xls:
                df = pd.read_excel(xls)
#                     move(df,key,'FS_Audit_Config','BS')
            if not df.empty:
                df_all_cur1.append(df)    
    if df_all_cur1:
        if (df.loc[df['FileName'] == 'FS_Audit_FBP.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_Audit_FBP.xlsx', 'Reload'] = 'N'
#             FS_Audit_FBP = run_script(f'{file_location}FS_Audit_FBP.py')

            lof_filename = ' FS_Audit_FBP'
            sharepoint_view = 'FS_Audit_Config'
            xlsx_name = 'FS_Audit_FBP.xlsx'
            company = 'FBP'
            FSVIEW = 'AD'
            database = 'DW_NCG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)   
    
            print(f'run {lof_filename} succes')
        else:
            pass
        if (df.loc[df['FileName'] == 'FS_Audit_NCI.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_Audit_NCI.xlsx', 'Reload'] = 'N'
#             FS_Audit_NCI = run_script(f'{file_location}FS_Audit_NCI.py')

            lof_filename = ' FS_Audit_NCI'
            sharepoint_view = 'FS_Audit_Config'
            xlsx_name = 'FS_Audit_NCI.xlsx'
            company = 'NCI'
            FSVIEW = 'AD'
            database = 'DW_NCG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)

            print(f'run {lof_filename} succes')
        else:
            pass
        if (df.loc[df['FileName'] == 'FS_Audit_NNC.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_Audit_NNC.xlsx', 'Reload'] = 'N'
#             FS_Audit_NNC = run_script(f'{file_location}FS_Audit_NNC.py')

            lof_filename = ' FS_Audit_NNC'
            sharepoint_view = 'FS_Audit_Config'
            xlsx_name = 'FS_Audit_NNC.xlsx'
            company = 'NNC'
            FSVIEW = 'AD'
            database = 'DW_NCG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
            
        else:
            pass
        if (df.loc[df['FileName'] == 'FS_Audit_NNCExNT.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_Audit_NNCExNT.xlsx', 'Reload'] = 'N'
#             FS_Audit_NNCExNT = run_script(f'{file_location}FS_Audit_NNCExNT.py')

            lof_filename = ' FS_Audit_NNCExNT'
            sharepoint_view = 'FS_Audit_Config'
            xlsx_name = 'FS_Audit_NNCExNT.xlsx'
            company = 'NNCExNT'
            FSVIEW = 'AD'
            database = 'DW_NCG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
#             GLCODE_FX_CONFIG = run_script(f'{file_location}GLCODE_FX_CONFIG.py')
#             print('run GLCODE_FX_CONFIG succes')
        else:
            pass
        
        if (df.loc[df['FileName'] == 'FS_Audit_RCI.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_Audit_RCI.xlsx', 'Reload'] = 'N'

            lof_filename = ' FS_Audit_RCI'
            sharepoint_view = 'FS_Audit_Config'
            xlsx_name = 'FS_Audit_RCI.xlsx'
            company = 'RCI'
            FSVIEW = 'AD'
            database = 'DW_RCG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
            
        else:
            pass
        
        if (df.loc[df['FileName'] == 'FS_Audit_EOS.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_Audit_EOS.xlsx', 'Reload'] = 'N'

            lof_filename = ' FS_Audit_EOS'
            sharepoint_view = 'FS_Audit_Config'
            xlsx_name = 'FS_Audit_EOS.xlsx'
            company = 'EOS'
            FSVIEW = 'AD'
            database = 'DW_RCG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
            
        else:
            pass   
        
        if (df.loc[df['FileName'] == 'FS_Audit_RSAC.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_Audit_RSAC.xlsx', 'Reload'] = 'N'

            lof_filename = ' FS_Audit_RSAC'
            sharepoint_view = 'FS_Audit_Config'
            xlsx_name = 'FS_Audit_RSAC.xlsx'
            company = 'RSAC'
            FSVIEW = 'AD'
            database = 'DW_RCG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
            
        else:
            pass   
        
        if (df.loc[df['FileName'] == 'FS_Audit_VSC.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_Audit_VSC.xlsx', 'Reload'] = 'N'

            lof_filename = ' FS_Audit_VSC'
            sharepoint_view = 'FS_Audit_Config'
            xlsx_name = 'FS_Audit_VSC.xlsx'
            company = 'VSC'
            FSVIEW = 'AD'
            database = 'DW_VSG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
            
        else:
            pass   
        
        if (df.loc[df['FileName'] == 'FS_Audit_VSST.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_Audit_VSST.xlsx', 'Reload'] = 'N'

            lof_filename = ' FS_Audit_VSST'
            sharepoint_view = 'FS_Audit_Config'
            xlsx_name = 'FS_Audit_VSST.xlsx'
            company = 'VSST'
            FSVIEW = 'AD'
            database = 'DW_VSG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
            
        else:
            pass   
        
        if (df.loc[df['FileName'] == 'FS_Audit_VSI.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_Audit_VSI.xlsx', 'Reload'] = 'N'

            lof_filename = ' FS_Audit_VSI'
            sharepoint_view = 'FS_Audit_Config'
            xlsx_name = 'FS_Audit_VSI.xlsx'
            company = 'VSI'
            FSVIEW = 'AD'
            database = 'DW_VSG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
            
        else:
            pass  
        
        move(df,key)
        #ctx.authentication.logout()
        ctx = None
    else:
        pass
    
    #############MM###############
#     site_url = "https://vschem365.sharepoint.com/sites/DataTeam"
#     username = "tadson.s@vschem.com"
#     password = "Kenyut@588"
    
    auth_context = AuthenticationContext(url=site_url)
    auth_context.acquire_token_for_user(username=username, password=password)
    ctx = ClientContext(site_url, auth_context)    
    source_folder_path = f"/sites/DataTeam/Shared Documents/Data Center/000-Pool/001-Account/FS_MM_Config"
    file_collection = ctx.web.get_folder_by_server_relative_url(source_folder_path).files
    ctx.load(file_collection)
    ctx.execute_query()
        
    
    def run_script(path):
        python_script = path
        process = subprocess.Popen(['python', python_script], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, stderr = process.communicate()
        print(f'run {path} success')    
    
    
    def move(df,key):    
        log_time_name = datetime.now().strftime("%Y%m%d%H%M%S")
        year = str(date.today().strftime("%Y"))
        df.to_excel(f'{log_path}{key}.xlsx', index=False)#df.to_excel(f'C:/Python/FX_CONVERT/log/{key}_{log_time_name}.xlsx', index=False) 
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        print('Connected to SharePoint: ',web.properties['Title'])
        
        file_name_to_delete = f"{key}.xlsx"
        file_path_to_delete = os.path.join(source_folder_path, file_name_to_delete)

        file_to_delete = ctx.web.get_file_by_server_relative_url(file_path_to_delete)
        file_to_delete.delete_object().execute_query()

        print(f"File '{file_name_to_delete}' has been deleted from the source folder.")

        fileName = f'{log_path}{key}.xlsx'#f'C:/Python/FX_CONVERT/log/{key}_{log_time_name}.xlsx'

        with open(fileName, 'rb') as content_file:
            file_content = content_file.read()     

        name = os.path.basename(fileName)
        list_title = "Documents"
        target_list = ctx.web.lists.get_by_title(list_title)
        #info = FileCreationInformation()
        destination_folder_path = f"/sites/DataTeam/Shared Documents/Data Center/000-Pool/001-Account/FS_MM_Config"
        libraryRoot = ctx.web.get_folder_by_server_relative_url(destination_folder_path)
#             DeliveryRoot = ctx.web.get_folder_by_server_relative_url(source_folder_path)
        os.remove(f'{log_path}{key}.xlsx')#os.remove(f'C:/Python/FX_CONVERT/log/{key}_{log_time_name}.xlsx')
        target_file = libraryRoot.upload_file(name, file_content).execute_query()
        print("File has been uploaded to url: {0}".format(target_file.serverRelativeUrl))



        # Cleanup and execute the batch
        ctx.execute_batch()        

        a = 'delete and upload success'
        return a 
    
    df_all_cur1 = []
    sorted_files = sorted(file_collection, key=lambda file: file.name)
    for file in sorted_files:
        if file.name == 'Reload_MM.xlsx':
            print(file.name)
            response = file.read()
            key = f"{file.name[:-5]}"
            with pd.ExcelFile(response) as xls:
                df = pd.read_excel(xls)
#                     move(df,key,'FS_Audit_Config','BS')
            if not df.empty:
                df_all_cur1.append(df)    
    if df_all_cur1:
        if (df.loc[df['FileName'] == 'FS_MM_FBP.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_MM_FBP.xlsx', 'Reload'] = 'N'
#             FS_Audit_FBP = run_script(f'{file_location}FS_Audit_FBP.py')

            lof_filename = ' FS_MM_FBP'
            sharepoint_view = 'FS_MM_Config'
            xlsx_name = 'FS_MM_FBP.xlsx'
            company = 'FBP'
            FSVIEW = 'MM'
            database = 'DW_NCG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  
    
            print(f'run {lof_filename} succes')
        else:
            pass
        if (df.loc[df['FileName'] == 'FS_MM_NCI.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_MM_NCI.xlsx', 'Reload'] = 'N'
#             FS_Audit_NCI = run_script(f'{file_location}FS_Audit_NCI.py')

            lof_filename = ' FS_MM_NCI'
            sharepoint_view = 'FS_MM_Config'
            xlsx_name = 'FS_MM_NCI.xlsx'
            company = 'NCI'
            FSVIEW = 'MM'
            database = 'DW_NCG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
        else:
            pass
        if (df.loc[df['FileName'] == 'FS_MM_NNC.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_MM_NNC.xlsx', 'Reload'] = 'N'
#             FS_Audit_NNC = run_script(f'{file_location}FS_Audit_NNC.py')

            lof_filename = ' FS_MM_NNC'
            sharepoint_view = 'FS_MM_Config'
            xlsx_name = 'FS_MM_NNC.xlsx'
            company = 'NNC'
            FSVIEW = 'MM'
            database = 'DW_NCG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
            
        else:
            pass
        if (df.loc[df['FileName'] == 'FS_MM_NNCExNT.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_MM_NNCExNT.xlsx', 'Reload'] = 'N'
#             FS_Audit_NNCExNT = run_script(f'{file_location}FS_Audit_NNCExNT.py')

            lof_filename = ' FS_MM_NNCExNT'
            sharepoint_view = 'FS_MM_Config'
            xlsx_name = 'FS_MM_NNCExNT.xlsx'
            company = 'NNCExNT'
            FSVIEW = 'MM'
            database = 'DW_NCG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
#             GLCODE_FX_CONFIG = run_script(f'{file_location}GLCODE_FX_CONFIG.py')
#             print('run GLCODE_FX_CONFIG succes')
        else:
            pass

        if (df.loc[df['FileName'] == 'FS_MM_RCI.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_MM_RCI.xlsx', 'Reload'] = 'N'

            lof_filename = ' FS_MM_RCI'
            sharepoint_view = 'FS_MM_Config'
            xlsx_name = 'FS_MM_RCI.xlsx'
            company = 'RCI'
            FSVIEW = 'MM'
            database = 'DW_RCG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
            
        else:
            pass 
        
        if (df.loc[df['FileName'] == 'FS_MM_EOS.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_MM_EOS.xlsx', 'Reload'] = 'N'

            lof_filename = ' FS_MM_EOS'
            sharepoint_view = 'FS_MM_Config'
            xlsx_name = 'FS_MM_EOS.xlsx'
            company = 'EOS'
            FSVIEW = 'MM'
            database = 'DW_RCG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
            
        else:
            pass 

        if (df.loc[df['FileName'] == 'FS_MM_RSAC.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_MM_RSAC.xlsx', 'Reload'] = 'N'

            lof_filename = ' FS_MM_RSAC'
            sharepoint_view = 'FS_MM_Config'
            xlsx_name = 'FS_MM_RSAC.xlsx'
            company = 'RSAC'
            FSVIEW = 'MM'
            database = 'DW_RCG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
            
        else:
            pass 
        
        if (df.loc[df['FileName'] == 'FS_MM_VSC.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_MM_VSC.xlsx', 'Reload'] = 'N'

            lof_filename = ' FS_MM_VSC'
            sharepoint_view = 'FS_MM_Config'
            xlsx_name = 'FS_MM_VSC.xlsx'
            company = 'VSC'
            FSVIEW = 'MM'
            database = 'DW_VSG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
            
        else:
            pass 
        
        if (df.loc[df['FileName'] == 'FS_MM_VSST.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_MM_VSST.xlsx', 'Reload'] = 'N'

            lof_filename = ' FS_MM_VSST'
            sharepoint_view = 'FS_MM_Config'
            xlsx_name = 'FS_MM_VSST.xlsx'
            company = 'VSST'
            FSVIEW = 'MM'
            database = 'DW_VSG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
            
        else:
            pass 
        
        if (df.loc[df['FileName'] == 'FS_MM_VSI.xlsx', 'Reload'] == 'Y').all(): 
            df.loc[df['FileName'] == 'FS_MM_VSI.xlsx', 'Reload'] = 'N'

            lof_filename = ' FS_MM_VSI'
            sharepoint_view = 'FS_MM_Config'
            xlsx_name = 'FS_MM_VSI.xlsx'
            company = 'VSI'
            FSVIEW = 'MM'
            database = 'DW_VSG'
            xls_config(log_path,lof_filename,sharepoint_view,xlsx_name,company,FSVIEW,database)  

            print(f'run {lof_filename} succes')
            
        else:
            pass     
    
        move(df,key)
        #ctx.authentication.logout()
        ctx = None
    else:
        pass
            
except Exception as e:
    print(f'Error: {str(e)}')