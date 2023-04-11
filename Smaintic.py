import tkinter as tk
from tkinter import *
from tkinter import ttk, filedialog
import pandas as pd
import os
import numpy as np
from PIL import ImageTk, Image
from pathlib import Path


class NewprojectApp:
    def __init__(self, master=None):

        self.Name = tk.Tk() if master is None else tk.Toplevel(master)
        self.Name.title('Smaintic')
        self.Name.wm_geometry("850x630")
        self.Name.resizable(False, False)
        self.Name.wm_iconbitmap('images/Haeco_logo.ico')


        self.Search = ttk.Frame(self.Name)
        self.Search.configure(height=500, width=400)
        self.Done_image = Image.open('images/Green_Tick_logo.png')
        self.resized2 = self.Done_image.resize((30, 30), Image.ANTIALIAS)
        self.Done_image = ImageTk.PhotoImage(self.resized2)

        def import_PNfile():
            file = filedialog.askopenfile(mode='r', filetypes=[('Excel Files', '*.xlsx'), ('CSV Files', '*.csv')])
            if file:
                global InputListFile
                InputListFile = os.path.abspath(file.name)
                self.B_Import.configure(image=self.Done_image)

        def clearentry():
            self.Entry1.delete(0, END)
            self.Entry2.delete(0, END)
            self.Entry3.delete(0, END)
            self.Entry4.delete(0, END)
            self.Entry5.delete(0, END)
            self.B_Import.configure(image=self.Import_File_image)

        def searchdata():
            selected = self.drop.get()
            if selected == 'Please Select':
                return

            if selected == 'Aircraft Type/Check Type/Engine Type':
                Aircraft_Type = self.Entry1.get()
                Check_Type = self.Entry2.get()
                Engine_Type = self.Entry3.get()
                End_Date = self.Entry5.get()
                Input_Date = pd.Timestamp(End_Date)
                readAITAR = pd.read_excel(AITAR, sheet_name='AMOS')
                readTool_Inventory = pd.read_excel(Tool_Inventory, sheet_name='Sheet1')
                readPE_AV_TASK_MONITORING = pd.read_excel(PE_AV_TASK_MONITORING, sheet_name='WORKPAD')
                readAircraft_Registration_Table_Full = pd.read_excel(Aircraft_Registration_Table_Full)
                readAMM_MPD_Data = pd.read_excel(AMM_MPD_Data, sheet_name='Tool List', header=2)
                readTooling_Load_History = pd.read_csv(Tooling_Load_History)

                if Engine_Type == '':
                    Sort1 = pd.DataFrame(readAircraft_Registration_Table_Full, columns=['A/C Register', 'Engine'])
                else:
                    Sort1 = pd.DataFrame(readAircraft_Registration_Table_Full, columns=['A/C Register', 'Engine'])
                    Sort1 = Sort1[Sort1['Engine'].str.contains(Engine_Type, na=False)]

                AR = Sort1['A/C Register'].values.tolist()
                strARlist = [str(x) for x in AR]
                ARlist = '|'.join(strARlist)

                Sort2 = pd.DataFrame(readPE_AV_TASK_MONITORING,
                                     columns=['REF', 'REG w/o "-"', 'A/C TYPE', 'CHECK TYPE'])
                Sort2 = Sort2[Sort2['REG w/o "-"'].str.contains(ARlist, na=False)]

                if Aircraft_Type == '':
                    Sort3 = pd.DataFrame(Sort2)
                else:
                    Sort3 = pd.DataFrame(Sort2)
                    Sort3 = Sort3[Sort3['A/C TYPE'].str.contains(Aircraft_Type, na=False)]

                if Check_Type == '':
                    Sort4 = pd.DataFrame(Sort3)
                    pattern = '|'.join(['CHG', 'SWAP', 'check', 'CHK'])
                    Sort4['CHECK TYPE'] = Sort4['CHECK TYPE'].str.replace(pattern, '', regex=True)
                else:
                    Sort4 = pd.DataFrame(Sort3)
                    pattern = '|'.join(['CHG', 'SWAP', 'check', 'CHK'])
                    Sort4['CHECK TYPE'] = Sort4['CHECK TYPE'].str.replace(pattern, '', regex=True)
                    Sort4 = Sort4[Sort4['CHECK TYPE'].str.contains(Check_Type, na=False)]

                if Aircraft_Type == '':
                    Sheet3Sort1 = pd.DataFrame(readAMM_MPD_Data,
                                               columns=['AMM Task', 'Part Number', 'Qty Required', 'Effectivity',
                                                        'Engine'])
                    Sheet3Sort1['Check Type'] = readAMM_MPD_Data.iloc[:, [14]]
                else:
                    Sheet3Sort1 = pd.DataFrame(readAMM_MPD_Data,
                                               columns=['AMM Task', 'Part Number', 'Qty Required', 'Effectivity',
                                                        'Engine'])
                    Sheet3Sort1['Check Type'] = readAMM_MPD_Data.iloc[:, [14]]
                    Sheet3Sort1 = Sheet3Sort1[Sheet3Sort1['AMM Task'].str.contains(Aircraft_Type, na=False)]

                if Check_Type == '':
                    Sheet3Sort2 = pd.DataFrame(Sheet3Sort1)
                    pattern2 = '|'.join(['Chk', 'Chg'])
                    Sheet3Sort2['Check Type'] = Sheet3Sort2['Check Type'].str.replace(pattern2, '')
                else:
                    Sheet3Sort2 = pd.DataFrame(Sheet3Sort1)
                    pattern2 = '|'.join(['Chk', 'Chg'])
                    Sheet3Sort2['Check Type'] = Sheet3Sort2['Check Type'].str.replace(pattern2, '')
                    Sheet3Sort2 = Sheet3Sort2[Sheet3Sort2['Check Type'].str.contains(Check_Type, na=False)]

                if Engine_Type == '':
                    Sheet3Sort3 = pd.DataFrame(Sheet3Sort2)
                else:
                    Sheet3Sort3 = pd.DataFrame(Sheet3Sort2)
                    Sheet3Sort3 = Sheet3Sort3[Sheet3Sort3['Engine'].str.contains(Engine_Type, na=False)]
                Sheet3Sort3['Part Number'] = Sheet3Sort3['Part Number'].str.replace('T00L', '', regex=True)
                delete_row = Sheet3Sort3[Sheet3Sort3['Part Number'] == 'No Specific'].index
                Sheet3Sort3 = Sheet3Sort3.drop(delete_row)

                MPDPNUsage = Sheet3Sort3['Part Number'].value_counts().reset_index()
                MPDPNUsage.columns = ['Part Number', 'Usage(MPD)']

                REF = Sort4['REF'].values.tolist()
                strREFlist = [str(x) for x in REF]
                REFlist = '|'.join(strREFlist)

                df = readAITAR[readAITAR['AITAR #'].str.contains(REFlist).fillna(False)]
                df['Part Number'] = df['Part Number'].str.replace('T00L', '', regex=True)
                Remarks = pd.DataFrame(df, columns=['Part Number', 'Remarks'])
                Remarks = Remarks.dropna().astype(str).sort_values(['Part Number', 'Remarks'],
                                                                   ascending=[True, False]).drop_duplicates(
                    subset=['Part Number', 'Remarks'])
                Remarks = Remarks.groupby('Part Number')['Remarks'].apply(' / '.join).reset_index()

                PNUsage = df['Part Number'].value_counts().reset_index()
                PNUsage.columns = ['Part Number', 'Usage(AITAR)']
                QOH1 = MPDPNUsage.merge(PNUsage, how='outer', on='Part Number')
                '''delete_row = QOH1[QOH1['Part Number'] == 'No Specific'].index
                QOH1 = QOH1.drop(delete_row)'''

                LoanPNUsage = pd.DataFrame(readTooling_Load_History, columns=['Partno'])
                LoanPNUsage['Partno'] = LoanPNUsage['Partno'].str.replace('T00L', '', regex=True)
                LoanPNUsage = LoanPNUsage['Partno'].value_counts().reset_index()
                LoanPNUsage.columns = ['Part Number', 'Usage(Loan & Return)']
                PNUsage1 = QOH1.merge(LoanPNUsage, how='left', on='Part Number')

                PNrequired = pd.DataFrame(df, columns=['Part Number', 'Req. Qty.'])
                PNrequired['Req. Qty.'] = PNrequired['Req. Qty.'].replace([np.nan, '-'], 0)
                PNrequired['Req. Qty.'] = PNrequired['Req. Qty.'].astype(float).astype(int)
                PNrequired = PNrequired.groupby(['Part Number'])['Req. Qty.'].max().reset_index()
                PNrequired.columns = ['Part Number', 'Required(AITAR)']

                '''RequiredPN = PNrequired['Part Number'].values.tolist()
                strRequiredPNlist = [str(x) for x in RequiredPN]
                RequiredPNlist = '|'.join(strRequiredPNlist)

                Sheet1Required = readAMM_MPD_Data[
                    readAMM_MPD_Data['Part Number'].str.contains(RequiredPNlist).fillna(False)]
                Sheet1Required['Part Number'] = Sheet1Required['Part Number'].str.replace('T00L', '', regex=True)
                Sheet1Required = pd.DataFrame(Sheet1Required, columns=['Part Number', 'Qty Required'])
                Sheet1Required['Qty Required'] = Sheet1Required['Qty Required'].replace(
                    ['Already Installed', 'As required'], 0)
                Sheet1Required['Qty Required'] = Sheet1Required['Qty Required'].astype(float).astype(int)
                Sheet1Required = Sheet1Required.groupby(['Part Number'])['Qty Required'].max().reset_index()
                Sheet1Calculation = PNrequired.merge(Sheet1Required, how='outer', on='Part Number')
                Sheet1Calculation['Required'] = Sheet1Calculation[['Required(AITAR)', 'Qty Required']].max(axis=1)
                Sheet1Calculation = pd.DataFrame(Sheet1Calculation, columns=['Part Number', 'Required'])'''

                Sheet1Required = pd.DataFrame(Sheet3Sort3, columns=['Part Number', 'Qty Required'])
                Sheet1Required['Qty Required'] = Sheet1Required['Qty Required'].replace(['Already Installed', 'As required'], 0)
                Sheet1Required['Qty Required'] = Sheet1Required['Qty Required'].astype(float).astype(int)
                Sheet1Required = Sheet1Required.groupby(['Part Number'])['Qty Required'].max().reset_index()
                Sheet1Calculation = PNrequired.merge(Sheet1Required, how='outer', on='Part Number')
                Sheet1Calculation['Required'] = Sheet1Calculation[['Required(AITAR)', 'Qty Required']].max(axis=1)
                Sheet1Calculation = pd.DataFrame(Sheet1Calculation, columns=['Part Number', 'Required'])
                PN1 = PNUsage1.merge(Sheet1Calculation, how='outer', on='Part Number')

                df2 = pd.DataFrame(readTool_Inventory)
                df2['partno'] = df2['partno'].str.replace('T00L', '', regex=True)
                df3 = df2.groupby(['partno'])['qty'].sum().reset_index()
                df3.columns = ['Part Number', 'QOH']
                QOH = PN1.merge(df3, how='left', on='Part Number')
                QOH[['Required', 'QOH']] = QOH[['Required', 'QOH']].fillna(0).astype(int)

                PN = QOH['Part Number'].values.tolist()
                strPNlist = [str(x) for x in PN]
                PNlist = '|'.join(strPNlist)

                Calibration = pd.DataFrame(df2, columns=['partno', 'sn_or_bn', 'condition', 'event_description',
                                                         'expiry_date'])
                Calibration = (Calibration[Calibration['partno'].str.contains(PNlist).fillna(False)])
                Calibration.rename(columns={'partno': 'Part Number'}, inplace=True)
                Calibration['within 30 days by now'] = Calibration['expiry_date'].apply(lambda x: 'True' if (
                                                                                                                        x - pd.Timestamp.now() <= pd.Timedelta(
                                                                                                                    '30 days')) and x >= pd.Timestamp.now() else (
                    'Expired' if x < pd.Timestamp.now() else ''))
                Calibration['within 30 days by input'] = Calibration['expiry_date'].apply(
                    lambda x: 'True' if (x - Input_Date <= pd.Timedelta('30 days')) and x >= Input_Date else (
                        'Expired' if x < Input_Date else ''))
                Calibration['expiry_date'] = Calibration['expiry_date'].dt.date
                if End_Date == '':
                    Calibration['within 30 days by input'] = Calibration['within 30 days by now']

                newCalibration = pd.DataFrame(Calibration, columns=['Part Number', 'within 30 days by input'])
                filter = newCalibration['within 30 days by input'] == 'Expired'
                newCalibration = newCalibration[filter]
                unservable = newCalibration.groupby(['Part Number'])['within 30 days by input'].count().reset_index()
                calculation1 = QOH.merge(unservable, how='outer', on='Part Number')

                newCalibration1 = pd.DataFrame(Calibration,
                                               columns=['Part Number', 'condition', 'within 30 days by input'])
                filter1 = newCalibration1['within 30 days by input'] != 'Expired'
                newCalibration1 = newCalibration1[filter1]
                filter2 = newCalibration1['condition'] == 'US'
                newCalibration1 = newCalibration1[filter2]
                unservable1 = newCalibration1.groupby(['Part Number'])['condition'].count().reset_index()
                calculation1 = calculation1.merge(unservable1, how='outer', on='Part Number')

                calculation1[['condition', 'within 30 days by input']] = calculation1[
                    ['condition', 'within 30 days by input']].fillna(0).astype(int)
                QOH['Qty Available'] = calculation1['QOH'] - calculation1['within 30 days by input'] - calculation1[
                    'condition']
                calculation1['QTY Readiness'] = QOH['Qty Available'] - QOH['Required']
                QOH['QTY Readiness'] = calculation1['QTY Readiness'].apply(lambda x: 'Unready' if x < 0 else '')
                QOH = QOH.merge(Remarks, how='left', on='Part Number')

                """
                Sheet4Sort1 = Sheet3Sort3['AMM Task'].values.tolist()
                strSheet4Sort1list = [str(x) for x in Sheet4Sort1]
                Sheet4Sort1list = '|'.join(strSheet4Sort1list)

                Sheet4Sort2 = df[df['Event'].str.contains(Sheet4Sort1list, na=False)]
                Sheet4Sort2['Req. Qty.'] = Sheet4Sort2['Req. Qty.'].replace([np.nan, '-'], 0)
                Sheet4Sort2['Req. Qty.'] = Sheet4Sort2['Req. Qty.'].astype(float).astype(int)

                Sheet4Remarks = pd.DataFrame(Sheet4Sort2, columns=['Part Number', 'Remarks'])
                Sheet4Remarks = Sheet4Remarks.dropna().astype(str).sort_values(['Part Number', 'Remarks'],
                                                                               ascending=[True, False]).drop_duplicates(
                    subset=['Part Number', 'Remarks'])
                Sheet4Remarks = Sheet4Remarks.groupby('Part Number')['Remarks'].apply(' / '.join).reset_index()

                Sheet4AITARPNUsage = Sheet4Sort2.groupby(['Part Number'])['Req. Qty.'].sum().reset_index()
                Sheet4AITARPNUsage.columns = ['Part Number', 'Usage (AITAR)']

                Sheet4Req = pd.DataFrame(Sheet3Sort3, columns=['Part Number', 'Qty Required'])
                Sheet4Req['Qty Required'] = Sheet4Req['Qty Required'].replace(['Already Installed', 'As required'], 0)
                Sheet4Req['Qty Required'] = Sheet4Req['Qty Required'].astype(float).astype(int)
                Sheet4AMMPNUsage = Sheet4Req.groupby(['Part Number'])['Qty Required'].sum().reset_index()
                Sheet4AMMPNUsage.columns = ['Part Number', 'Usage (AMM)']
                Sheet4Output = Sheet4AITARPNUsage.merge(Sheet4AMMPNUsage, how='left', on='Part Number')

                Sheet4Req = Sheet4Req.groupby(['Part Number'])['Qty Required'].max().reset_index()

                Sheet4Calculation = PNrequired.merge(Sheet4Req, how='left', on='Part Number')
                Sheet4Output['Qty Required'] = Sheet4Calculation[['Required', 'Qty Required']].max(axis=1)

                Sheet4PN = Sheet4Sort2['Part Number'].values.tolist()
                strSheet4PNlist = [str(x) for x in Sheet4PN]
                Sheet4PNlist = '|'.join(strSheet4PNlist)

                Sheet4df2 = (readTool_Inventory[readTool_Inventory['partno'].str.contains(Sheet4PNlist).fillna(False)])
                Sheet4df2['partno'] = Sheet4df2['partno'].str.replace('T00L', '', regex=True)
                Sheet4df3 = Sheet4df2.groupby(['partno'])['qty'].sum().reset_index()
                Sheet4df3.columns = ['Part Number', 'QOH']
                Sheet4Output2 = Sheet4Output.merge(Sheet4df3, how='left', on='Part Number')

                Sheet4calculation1 = Sheet4Calculation.merge(unservable, how='left', on='Part Number')
                Sheet4calculation1 = Sheet4calculation1.merge(unservable1, how='left', on='Part Number')
                Sheet4calculation1 = Sheet4calculation1.merge(Sheet4df3, how='left', on='Part Number')
                Sheet4calculation1[['condition', 'within 30 days by input', 'QOH', 'Qty Required']] = \
                Sheet4calculation1[['condition', 'within 30 days by input', 'QOH', 'Qty Required']].fillna(0).astype(
                    int)
                Sheet4calculation1['QTY Readiness'] = Sheet4calculation1['QOH'] - Sheet4calculation1[
                    'within 30 days by input'] - Sheet4calculation1['condition'] - Sheet4calculation1['Qty Required']
                Sheet4Output2['QTY Readiness'] = Sheet4calculation1['QTY Readiness'].apply(
                    lambda x: 'Unready' if x < 0 else '')
                Sheet4Output3 = Sheet4Output2.merge(Sheet4Remarks, how='left', on='Part Number')
                """
                Item1_Output_File = filedialog.asksaveasfilename(defaultextension='.xlsx', initialfile='Item1 Output')
                with pd.ExcelWriter(Item1_Output_File) as writer:
                    QOH.to_excel(writer, sheet_name='Summary List', index=False)
                    Calibration.to_excel(writer, sheet_name='Detail List', index=False)
                    Sheet3Sort3.to_excel(writer, sheet_name='AMM Task List', index=False)
                    #Sheet4Output3.to_excel(writer, sheet_name='Common Task List', index=False)
                os.startfile(Item1_Output_File)

            if selected == 'Tool PN':
                Tool_PN = self.Entry1.get()
                End_Date = self.Entry3.get()
                Input_Date = pd.Timestamp(End_Date)
                readAITAR = pd.read_excel(AITAR, sheet_name='AMOS')
                readTool_Inventory = pd.read_excel(Tool_Inventory, sheet_name='Sheet1')
                readPE_AV_TASK_MONITORING = pd.read_excel(PE_AV_TASK_MONITORING, sheet_name='WORKPAD')

                df = (readTool_Inventory[readTool_Inventory['partno'].str.contains(Tool_PN).fillna(False)])
                df['partno'] = df['partno'].str.replace('T00L', '', regex=True)
                QOH = df.groupby(['partno'])['qty'].sum().reset_index()
                QOH.columns = ['Part Number', 'QOH']

                df2 = pd.DataFrame(readAITAR)
                df2['Part Number'] = df2['Part Number'].str.replace('T00L', '', regex=True)
                df2 = df2[df2['Part Number'].str.contains(Tool_PN, na=False)]

                PNUsage = df2['Part Number'].value_counts().reset_index()
                PNUsage.columns = ['Part Number', 'Usage(AITAR)']
                Output = QOH.merge(PNUsage, how='outer', on='Part Number')

                df21 = pd.DataFrame(df2, columns=['AITAR #', 'Part Number'])
                df21.columns = ['REF', 'Part Number']
                REF = df21['REF'].values.tolist()
                strREFlist = [str(x) for x in REF]
                REFlist = '|'.join(strREFlist)

                df3 = pd.DataFrame(readPE_AV_TASK_MONITORING, columns=['REF', 'A/C TYPE'])
                df3 = df3[df3['REF'].str.contains(REFlist, na=False)]
                df4 = df21.merge(df3, how='left', on='REF')
                Applicable = pd.DataFrame(df4, columns=['Part Number', 'A/C TYPE'])
                Applicable = Applicable.dropna().astype(str).sort_values(['Part Number', 'A/C TYPE'],
                                                                         ascending=[True, False]).drop_duplicates(
                    subset=['Part Number', 'A/C TYPE'])
                Applicable = Applicable.groupby('Part Number')['A/C TYPE'].apply(' / '.join).reset_index()
                Output2 = Output.merge(Applicable, how='outer', on='Part Number')

                Remarks = pd.DataFrame(readAITAR, columns=['Part Number', 'Remarks'])
                Remarks = Remarks[Remarks['Part Number'].str.contains(Tool_PN, na=False)]
                Remarks = Remarks.dropna().astype(str).sort_values(['Part Number', 'Remarks'],
                                                                   ascending=[True, False]).drop_duplicates(
                    subset=['Part Number', 'Remarks'])
                Remarks = Remarks.groupby('Part Number')['Remarks'].apply(' / '.join).reset_index()
                Output3 = Output2.merge(Remarks, how='left', on='Part Number')

                Sheet2SN = (readTool_Inventory[readTool_Inventory['partno'].str.contains(Tool_PN).fillna(False)])
                Sheet2SN['partno'] = Sheet2SN['partno'].str.replace('T00L', '', regex=True)
                Calibration = pd.DataFrame(Sheet2SN, columns=['partno', 'sn_or_bn', 'condition', 'event_description',
                                                         'expiry_date'])
                Calibration.rename(columns={'partno': 'Part Number'}, inplace=True)
                Calibration['within 30 days by now'] = Calibration['expiry_date'].apply(lambda x: 'True' if (
                                                                                                                    x - pd.Timestamp.now() <= pd.Timedelta(
                                                                                                                '30 days')) and x >= pd.Timestamp.now() else (
                    'Expired' if x < pd.Timestamp.now() else ''))
                Calibration['within 30 days by input'] = Calibration['expiry_date'].apply(
                    lambda x: 'True' if (x - Input_Date <= pd.Timedelta('30 days')) and x >= Input_Date else (
                        'Expired' if x < Input_Date else ''))
                Calibration['expiry_date'] = Calibration['expiry_date'].dt.date

                Item2_Output_File = filedialog.asksaveasfilename(defaultextension='.xlsx', initialfile='Item2 Output')
                with pd.ExcelWriter(Item2_Output_File) as writer:
                    Output3.to_excel(writer, sheet_name='Summary List', index=False)
                    Calibration.to_excel(writer, sheet_name='Detail List', index=False)
                os.startfile(Item2_Output_File)

            if selected == 'Tool PN List':
                End_Date = self.Entry2.get()
                Input_Date = pd.Timestamp(End_Date)
                readToolPNList = pd.read_excel(InputListFile)
                readAITAR = pd.read_excel(AITAR, sheet_name='AMOS')
                readTool_Inventory = pd.read_excel(Tool_Inventory, sheet_name='Sheet1')
                readPE_AV_TASK_MONITORING = pd.read_excel(PE_AV_TASK_MONITORING, sheet_name='WORKPAD')

                readToolPNList = readToolPNList.drop_duplicates()
                item3ToolPN = readToolPNList['Part Number'].values.tolist()
                stritem3ToolPNlist = [str(x) for x in item3ToolPN]
                item3ToolPNlist = '|'.join(stritem3ToolPNlist)

                df = readTool_Inventory[readTool_Inventory['partno'].str.contains(item3ToolPNlist).fillna(False)]
                df['partno'] = df['partno'].str.replace('T00L', '', regex=True)
                QOH = df.groupby(['partno'])['qty'].sum().reset_index()
                QOH.columns = ['Part Number', 'QOH']

                df2 = pd.DataFrame(readAITAR, columns=['AITAR #', 'Part Number'])
                df2['Part Number'] = df2['Part Number'].str.replace('T00L', '', regex=True)
                df2 = df2[df2['Part Number'].str.contains(item3ToolPNlist, na=False)]
                df2.columns = ['REF', 'Part Number']
                REF = df2['REF'].values.tolist()
                strREFlist = [str(x) for x in REF]
                REFlist = '|'.join(strREFlist)
                df3 = pd.DataFrame(readPE_AV_TASK_MONITORING, columns=['REF', 'A/C TYPE'])
                df3 = df3[df3['REF'].str.contains(REFlist, na=False)]
                df4 = df2.merge(df3, how='left', on='REF')
                Applicable = pd.DataFrame(df4, columns=['Part Number', 'A/C TYPE'])
                Applicable = Applicable.dropna().astype(str).sort_values(['Part Number', 'A/C TYPE'],
                                                                         ascending=[True, False]).drop_duplicates(
                    subset=['Part Number', 'A/C TYPE'])
                Applicable = Applicable.groupby('Part Number')['A/C TYPE'].apply(' / '.join).reset_index()
                Output = QOH.merge(Applicable, how='outer', on='Part Number')

                Remarks = pd.DataFrame(readAITAR, columns=['Part Number', 'Remarks'])
                Remarks['Part Number'] = Remarks['Part Number'].str.replace('T00L', '', regex=True)
                Remarks = Remarks[Remarks['Part Number'].str.contains(item3ToolPNlist, na=False)]
                Remarks = Remarks.dropna().astype(str).sort_values(['Part Number', 'Remarks'],
                                                                   ascending=[True, False]).drop_duplicates(
                    subset=['Part Number', 'Remarks'])
                Remarks = Remarks.groupby('Part Number')['Remarks'].apply(' / '.join).reset_index()
                Output2 = Output.merge(Remarks, how='outer', on='Part Number')

                Calibration = pd.DataFrame(df, columns=['partno', 'sn_or_bn', 'condition', 'event_description',
                                                         'expiry_date'])
                Calibration.rename(columns={'partno': 'Part Number'}, inplace=True)
                Calibration['within 30 days by now'] = Calibration['expiry_date'].apply(lambda x: 'True' if (
                                                                                                                    x - pd.Timestamp.now() <= pd.Timedelta(
                                                                                                                '30 days')) and x >= pd.Timestamp.now() else (
                    'Expired' if x < pd.Timestamp.now() else ''))
                Calibration['within 30 days by input'] = Calibration['expiry_date'].apply(
                    lambda x: 'True' if (x - Input_Date <= pd.Timedelta('30 days')) and x >= Input_Date else (
                        'Expired' if x < Input_Date else ''))
                Calibration['expiry_date'] = Calibration['expiry_date'].dt.date
                if End_Date == '':
                    Calibration['within 30 days by input'] = Calibration['within 30 days by now']


                Item3_Output_File = filedialog.asksaveasfilename(defaultextension='.xlsx', initialfile='Item3 Output')
                with pd.ExcelWriter(Item3_Output_File) as writer:
                    Output2.to_excel(writer, sheet_name='Summary List', index=False)
                    Calibration.to_excel(writer, sheet_name='Detail List', index=False)
                os.startfile(Item3_Output_File)

            if selected == 'Aircraft Input Workpack':
                End_Date = self.Entry2.get()
                Input_Date = pd.Timestamp(End_Date)
                readTaskList = pd.read_excel(InputListFile)
                readAITAR = pd.read_excel(AITAR, sheet_name='AMOS')
                readTool_Inventory = pd.read_excel(Tool_Inventory, sheet_name='Sheet1')
                readTooling_Load_History = pd.read_csv(Tooling_Load_History)
                readAMM_MPD_Data = pd.read_excel(AMM_MPD_Data, sheet_name='Tool List', header=2)

                Event = readTaskList['Event'].values.tolist()
                strEventlist = [str(x) for x in Event]
                Eventlist = '|'.join(strEventlist)

                df = readAITAR[readAITAR['Event'].str.contains(Eventlist).fillna(False)]
                df['Part Number'] = df['Part Number'].str.replace('T00L', '', regex=True)
                Sheet1Output = pd.DataFrame(df, columns=['Event', 'Part Number'])
                Sheet1Output = Sheet1Output.groupby(['Event', 'Part Number']).first()
                Sheet1Output = Sheet1Output.reset_index()

                MPD = readAMM_MPD_Data[readAMM_MPD_Data['AMM Task'].str.contains(Eventlist).fillna(False)]
                MPD['Part Number'] = MPD['Part Number'].str.replace('T00L', '', regex=True)
                delete_row = MPD[MPD['Part Number'] == 'No Specific'].index
                MPD = MPD.drop(delete_row)
                Sheet1Output1 = pd.DataFrame(MPD, columns=['AMM Task', 'Part Number'])
                Sheet1Output1 = Sheet1Output1.groupby(['AMM Task', 'Part Number']).first()
                Sheet1Output1 = Sheet1Output1.reset_index()
                Sheet1Output1.columns = ['Event', 'Part Number']
                Sheet1Output2 = pd.merge(Sheet1Output, Sheet1Output1, how='outer')

                MPDPNUsage = pd.DataFrame(MPD, columns=['Part Number'])
                MPDPNUsage = MPDPNUsage['Part Number'].value_counts().reset_index()
                MPDPNUsage.columns = ['Part Number', 'Usage(MPD)']

                PNUsage = df['Part Number'].value_counts().reset_index()
                PNUsage.columns = ['Part Number', 'Usage(AITAR)']
                PNUsage1 = MPDPNUsage.merge(PNUsage, how='outer', on='Part Number')

                LoanPNUsage = pd.DataFrame(readTooling_Load_History, columns=['Partno'])
                LoanPNUsage['Partno'] = LoanPNUsage['Partno'].str.replace('T00L', '', regex=True)
                LoanPNUsage = LoanPNUsage['Partno'].value_counts().reset_index()
                LoanPNUsage.columns = ['Part Number', 'Usage(Loan & Return)']
                PNUsage1 = PNUsage1.merge(LoanPNUsage, how='outer', on='Part Number')

                PNrequired = pd.DataFrame(df, columns=['Part Number', 'Req. Qty.'])
                PNrequired['Req. Qty.'] = PNrequired['Req. Qty.'].replace([np.nan, '-'], 0)
                PNrequired['Req. Qty.'] = PNrequired['Req. Qty.'].astype(float).astype(int)
                PNrequired = PNrequired.groupby(['Part Number'])['Req. Qty.'].max().reset_index()
                PNrequired.columns = ['Part Number', 'Required(AITAR)']

                Sheet1Required = pd.DataFrame(MPD, columns=['Part Number', 'Qty Required'])
                Sheet1Required['Qty Required'] = Sheet1Required['Qty Required'].replace(['Already Installed', 'As required'], 0)
                Sheet1Required['Qty Required'] = Sheet1Required['Qty Required'].astype(float).astype(int)
                Sheet1Required = Sheet1Required.groupby(['Part Number'])['Qty Required'].max().reset_index()
                Sheet1Calculation = PNrequired.merge(Sheet1Required, how='outer', on='Part Number')
                Sheet1Calculation['Required'] = Sheet1Calculation[['Required(AITAR)', 'Qty Required']].max(axis=1)
                Sheet1Calculation = pd.DataFrame(Sheet1Calculation, columns=['Part Number', 'Required'])
                PN1 = PNUsage1.merge(Sheet1Calculation, how='outer', on='Part Number')

                df2 = pd.DataFrame(readTool_Inventory)
                df2['partno'] = df2['partno'].str.replace('T00L', '', regex=True)
                df3 = df2.groupby(['partno'])['qty'].sum().reset_index()
                df3.columns = ['Part Number', 'QOH']
                QOH = PN1.merge(df3, how='outer', on='Part Number')
                QOH[['Required', 'QOH']] = QOH[['Required', 'QOH']].fillna(0).astype(int)

                PN = Sheet1Output2['Part Number'].values.tolist()
                strPNlist = [str(x) for x in PN]
                PNlist = '|'.join(strPNlist)

                Calibration = pd.DataFrame(df2, columns=['partno', 'sn_or_bn', 'condition', 'event_description',
                                                         'expiry_date'])
                Calibration = (Calibration[Calibration['partno'].str.contains(PNlist).fillna(False)])
                Calibration.rename(columns={'partno': 'Part Number'}, inplace=True)
                Calibration['within 30 days by now'] = Calibration['expiry_date'].apply(lambda x: 'True' if (
                                                                                                                    x - pd.Timestamp.now() <= pd.Timedelta(
                                                                                                                '30 days')) and x >= pd.Timestamp.now() else (
                    'Expired' if x < pd.Timestamp.now() else ''))
                Calibration['within 30 days by input'] = Calibration['expiry_date'].apply(
                    lambda x: 'True' if (x - Input_Date <= pd.Timedelta('30 days')) and x >= Input_Date else (
                        'Expired' if x < Input_Date else ''))
                Calibration['expiry_date'] = Calibration['expiry_date'].dt.date
                if End_Date == '':
                    Calibration['within 30 days by input'] = Calibration['within 30 days by now']

                newCalibration = pd.DataFrame(Calibration, columns=['Part Number', 'within 30 days by input'])
                filter = newCalibration['within 30 days by input'] == 'Expired'
                newCalibration = newCalibration[filter]
                unservable = newCalibration.groupby(['Part Number'])['within 30 days by input'].count()
                calculation1 = QOH.merge(unservable, how='outer', on='Part Number')

                newCalibration1 = pd.DataFrame(Calibration,
                                               columns=['Part Number', 'condition', 'within 30 days by input'])
                filter1 = newCalibration1['within 30 days by input'] != 'Expired'
                newCalibration1 = newCalibration1[filter1]
                filter2 = newCalibration1['condition'] == 'US'
                newCalibration1 = newCalibration1[filter2]
                unservable1 = newCalibration1.groupby(['Part Number'])['condition'].count()
                calculation1 = calculation1.merge(unservable1, how='outer', on='Part Number')

                calculation1[['condition', 'within 30 days by input']] = calculation1[
                    ['condition', 'within 30 days by input']].fillna(0).astype(int)
                calculation1['QTY Readiness'] = calculation1['QOH'] - calculation1['within 30 days by input'] - \
                                                  calculation1['condition'] - calculation1['Required']
                QOH['Qty Available'] = calculation1['QOH'] - calculation1['within 30 days by input'] - calculation1[
                    'condition']
                QOH['QTY Readiness'] = calculation1['QTY Readiness'].apply(
                    lambda x: 'Unready' if x < 0 else '')

                Remarks = pd.DataFrame(df, columns=['Part Number', 'Remarks'])
                Remarks = Remarks.dropna().astype(str).sort_values(['Part Number', 'Remarks'],
                                                                   ascending=[True, False]).drop_duplicates(
                    subset=['Part Number', 'Remarks'])
                Remarks = Remarks.groupby('Part Number')['Remarks'].apply(' / '.join).reset_index()
                QOH = QOH.merge(Remarks, how='outer', on='Part Number')
                QOH = Sheet1Output2.merge(QOH, how='outer', on='Part Number')

                Item4_Output_File = filedialog.asksaveasfilename(defaultextension='.xlsx', initialfile='Item4 Output')
                with pd.ExcelWriter(Item4_Output_File) as writer:
                    QOH.to_excel(writer, sheet_name='Summary List', index=False)
                    Calibration.to_excel(writer, sheet_name='Detail List', index=False)
                os.startfile(Item4_Output_File)

        self.Search_By = ttk.Label(self.Search)
        self.Search_By.configure(text='Search By:')
        self.Search_By.grid(column=0, padx=40, pady=20, row=0)

        self.B_Search = ttk.Button(self.Search, command=searchdata)
        self.B_Search.configure(text='Search')
        self.B_Search.grid(column=0, columnspan=2, ipadx=10, pady=10, row=7)

        self.B_Clear = ttk.Button(self.Search, command=clearentry)
        self.B_Clear.configure(text='Clear')
        self.B_Clear.grid(column=0, columnspan=2, ipadx=10, pady=10, row=8)

        self.drop = tk.StringVar(value='Please Select')
        combobox = ttk.Combobox(self.Search, textvariable=self.drop)
        combobox['values'] = ('Please Select', 'Aircraft Type/Check Type/Engine Type', 'Tool PN', 'Tool PN List', 'Aircraft Input Workpack')
        combobox.grid(column=1, ipadx=35, padx=10, row=0, sticky="w")

        self.Title1 = ttk.Label(self.Search)
        self.Title1.configure(text='')
        self.Title1.grid(column=0, pady=20, row=1)
        self.Title2 = ttk.Label(self.Search)
        self.Title2.configure(text='')
        self.Title2.grid(column=0, pady=20, row=2)
        self.Title3 = ttk.Label(self.Search)
        self.Title3.configure(text='')
        self.Title3.grid(column=0, pady=20, row=3)
        self.Title4 = ttk.Label(self.Search)
        self.Title4.configure(text='')
        self.Title4.grid(column=0, pady=20, row=4)
        self.Title5 = ttk.Label(self.Search)
        self.Title5.configure(text='')
        self.Title5.grid(column=0, pady=20, row=5)
        self.Title6 = ttk.Label(self.Search)
        self.Title6.configure(text='')
        self.Title6.grid(column=0, pady=20, row=6)

        self.Entry1 = ttk.Entry(self.Search)
        self.Entry2 = ttk.Entry(self.Search)
        self.Entry3 = ttk.Entry(self.Search)
        self.Entry4 = ttk.Entry(self.Search)
        self.Entry5 = ttk.Entry(self.Search)
        self.Import_File_image = Image.open('images/Import_File_logo.png')
        self.resized3 = self.Import_File_image.resize((30, 30), Image.ANTIALIAS)
        self.Import_File_image = ImageTk.PhotoImage(self.resized3)
        self.B_Import = ttk.Button(self.Search, command=import_PNfile)
        self.B_Import.configure(image=self.Import_File_image)

        combobox['state'] = 'readonly'
        combobox.grid(column=1, ipadx=35, padx=10, row=0, sticky="w")

        def value_changed(event):
            if self.drop.get() == 'Please Select':
                self.Title1.configure(text='')
                self.Title1.grid(column=0, pady=20, row=1)
                self.Title2.configure(text='')
                self.Title2.grid(column=0, pady=20, row=2)
                self.Title3.configure(text='')
                self.Title3.grid(column=0, pady=20, row=3)
                self.Title4.configure(text='')
                self.Title4.grid(column=0, pady=20, row=4)
                self.Title5.configure(text='')
                self.Title5.grid(column=0, pady=20, row=5)
                self.Title6.configure(text='')
                self.Title6.grid(column=0, pady=20, row=6)

                self.Entry1.grid_remove()
                self.Entry2.grid_remove()
                self.Entry3.grid_remove()
                self.Entry4.grid_remove()
                self.Entry5.grid_remove()
                self.B_Import.grid_remove()

            elif self.drop.get() == 'Aircraft Type/Check Type/Engine Type':
                self.Title1.configure(text='A/C Type')
                self.Title1.grid(column=0, pady=20, row=1)
                self.Title2.configure(text='Chk Type')
                self.Title2.grid(column=0, pady=20, row=2)
                self.Title3.configure(text='Engine Type')
                self.Title3.grid(column=0, pady=20, row=3)
                self.Title4.configure(text='Start Date')
                self.Title4.grid(column=0, pady=20, row=4)
                self.Title5.configure(text='End Date')
                self.Title5.grid(column=0, pady=20, row=5)
                self.Title6.configure(text='')
                self.Title6.grid(column=0, pady=20, row=6)

                self.Entry1.grid(column=1, ipadx=20, ipady=3, padx=5, row=1, sticky="w")
                self.Entry2.grid(column=1, ipadx=20, ipady=3, padx=5, row=2, sticky="w")
                self.Entry3.grid(column=1, ipadx=20, ipady=3, padx=5, row=3, sticky="w")
                self.Entry4.grid(column=1, ipadx=20, ipady=3, padx=5, row=4, sticky="w")
                self.Entry5.grid(column=1, ipadx=20, ipady=3, padx=5, row=5, sticky="w")
                self.B_Import.grid_remove()

            elif self.drop.get() == 'Tool PN':
                self.Title1.configure(text='Tool PN')
                self.Title1.grid(column=0, pady=20, row=1)
                self.Title2.configure(text='Start Date')
                self.Title2.grid(column=0, pady=20, row=2)
                self.Title3.configure(text='End Date')
                self.Title3.grid(column=0, pady=20, row=3)
                self.Title4.configure(text='')
                self.Title4.grid(column=0, pady=20, row=4)
                self.Title5.configure(text='')
                self.Title5.grid(column=0, pady=20, row=5)
                self.Title6.configure(text='')
                self.Title6.grid(column=0, pady=20, row=6)

                self.Entry1.grid(column=1, ipadx=20, ipady=3, padx=5, row=1, sticky="w")
                self.Entry2.grid(column=1, ipadx=20, ipady=3, padx=5, row=2, sticky="w")
                self.Entry3.grid(column=1, ipadx=20, ipady=3, padx=5, row=3, sticky="w")
                self.Entry4.grid_remove()
                self.Entry5.grid_remove()
                self.B_Import.grid_remove()

            elif self.drop.get() == 'Tool PN List':
                self.Title1.configure(text='Start Date')
                self.Title1.grid(column=0, pady=20, row=1)
                self.Title2.configure(text='End Date')
                self.Title2.grid(column=0, pady=20, row=2)
                self.Title3.configure(text='')
                self.Title3.grid(column=0, pady=20, row=3)
                self.Title4.configure(text='')
                self.Title4.grid(column=0, pady=20, row=4)
                self.Title5.configure(text='')
                self.Title5.grid(column=0, pady=20, row=5)
                self.Title6.configure(text='Import File:')
                self.Title6.grid(column=0, pady=20, row=6)

                self.Entry1.grid(column=1, ipadx=20, ipady=3, padx=5, row=1, sticky="w")
                self.Entry2.grid(column=1, ipadx=20, ipady=3, padx=5, row=2, sticky="w")
                self.Entry3.grid_remove()
                self.Entry4.grid_remove()
                self.Entry5.grid_remove()
                self.B_Import.grid(column=1, ipadx=40, ipady=10, padx=10, pady=10, row=6, sticky="w")

            elif self.drop.get() == 'Aircraft Input Workpack':
                self.Title1.configure(text='Start Date')
                self.Title1.grid(column=0, pady=20, row=1)
                self.Title2.configure(text='End Date')
                self.Title2.grid(column=0, pady=20, row=2)
                self.Title3.configure(text='')
                self.Title3.grid(column=0, pady=20, row=3)
                self.Title4.configure(text='')
                self.Title4.grid(column=0, pady=20, row=4)
                self.Title5.configure(text='')
                self.Title5.grid(column=0, pady=20, row=5)
                self.Title6.configure(text='Import File:')
                self.Title6.grid(column=0, pady=20, row=6)

                self.Entry1.grid(column=1, ipadx=20, ipady=3, padx=5, row=1, sticky="w")
                self.Entry2.grid(column=1, ipadx=20, ipady=3, padx=5, row=2, sticky="w")
                self.Entry3.grid_remove()
                self.Entry4.grid_remove()
                self.Entry5.grid_remove()
                self.B_Import.grid(column=1, ipadx=40, ipady=10, padx=10, pady=10, row=6, sticky="w")

        combobox.bind('<<ComboboxSelected>>', value_changed)
        self.Search.grid(column=0, row=0)



        self.Import = ttk.Frame(self.Name)
        self.Import.configure(height=500, width=400)

        def import_AITAR():
            file = filedialog.askopenfile(mode='r', filetypes=[('Excel Files', '*.xlsx'), ('CSV Files', '*.csv')])
            if file:
                global AITAR
                AITAR = os.path.abspath(file.name)
                self.B_AITAR.configure(text=Path(AITAR).stem, image=self.Done_image, compound=LEFT, width=5)
                #self.L_AITAR = Label(self.Import, text=Path(AITAR).stem)
                #self.L_AITAR.grid(column=3, padx=50, pady=20, row=1)

        def import_Tooling_Load_History():
            file = filedialog.askopenfile(mode='r', filetypes=[('CSV Files', '*.csv'), ('Excel Files', '*.xlsx')])
            if file:
                global Tooling_Load_History
                Tooling_Load_History = os.path.abspath(file.name)
                self.B_Tooling_Load_History.configure(text=Path(Tooling_Load_History).stem, image=self.Done_image, compound=LEFT, width=5)
                #self.L_Tooling_Load_History = Label(self.Import, text=Path(Tooling_Load_History).stem)
                #self.L_Tooling_Load_History.grid(column=3, padx=50, pady=20, row=2)

        def import_AMM_MPD_Data():
            file = filedialog.askopenfile(mode='r', filetypes=[('Excel Files', '*.xlsx'), ('CSV Files', '*.csv')])
            if file:
                global AMM_MPD_Data
                AMM_MPD_Data = os.path.abspath(file.name)
                self.B_AMM_MPD_Data.configure(text=Path(AMM_MPD_Data).stem, image=self.Done_image, compound=LEFT, width=5)
                #self.L_AMM_MPD_Data = Label(self.Import, text=Path(AMM_MPD_Data).stem)
                #self.L_AMM_MPD_Data.grid(column=3, padx=50, pady=20, row=3)

        def import_Tool_Inventory():
            file = filedialog.askopenfile(mode='r', filetypes=[('Excel Files', '*.xlsx'), ('CSV Files', '*.csv')])
            if file:
                global Tool_Inventory
                Tool_Inventory = os.path.abspath(file.name)
                self.B_Tool_Inventory.configure(text=Path(Tool_Inventory).stem, image=self.Done_image, compound=LEFT, width=5)
                #self.L_Tool_Inventory = Label(self.Import, text=Path(Tool_Inventory).stem)
                #self.L_Tool_Inventory.grid(column=3, padx=50, pady=20, row=4)

        def import_Calibration_Control_Data():
            file = filedialog.askopenfile(mode='r', filetypes=[('Excel Files', '*.xlsx'), ('CSV Files', '*.csv')])
            if file:
                global Calibration_Control_Data
                Calibration_Control_Data = os.path.abspath(file.name)
                self.B_Calibration_Control_Data.configure(text=Path(Calibration_Control_Data).stem, image=self.Done_image, compound=LEFT, width=5)
                #self.L_Calibration_Control_Data = Label(self.Import, text=Path(Calibration_Control_Data).stem)
                #self.L_Calibration_Control_Data.grid(column=3, padx=50, pady=20, row=5)

        def import_PE_AV_TASK_MONITORING():
            file = filedialog.askopenfile(mode='r', filetypes=[('Excel Files', '*.xlsx'), ('CSV Files', '*.csv')])
            if file:
                global PE_AV_TASK_MONITORING
                PE_AV_TASK_MONITORING = os.path.abspath(file.name)
                self.B_PE_AV_TASK_MONITORING.configure(text=Path(PE_AV_TASK_MONITORING).stem, image=self.Done_image, compound=LEFT, width=5)
                #self.L_PE_AV_TASK_MONITORING = Label(self.Import, text=Path(PE_AV_TASK_MONITORING).stem)
                #self.L_PE_AV_TASK_MONITORING.grid(column=3, padx=50, pady=20, row=6)

        def import_Aircraft_Registration_Table_Full():
            file = filedialog.askopenfile(mode='r', filetypes=[('Excel Files', '*.xlsx'), ('CSV Files', '*.csv')])
            if file:
                global Aircraft_Registration_Table_Full
                Aircraft_Registration_Table_Full = os.path.abspath(file.name)
                self.B_Aircraft_Registration_Table_Full.configure(text=Path(Aircraft_Registration_Table_Full).stem, image=self.Done_image, compound=LEFT, width=5)
                #self.L_Aircraft_Registration_Table_Full = Label(self.Import, text=Path(Aircraft_Registration_Table_Full).stem)
                #self.L_Aircraft_Registration_Table_Full.grid(column=3, padx=50, pady=20, row=7)

        def clearimported():
            self.B_AITAR.configure(text='', image=self.Import_File_image, width=5)
            self.B_Tooling_Load_History.configure(text='', image=self.Import_File_image, width=5)
            self.B_AMM_MPD_Data.configure(text='', image=self.Import_File_image, width=5)
            self.B_Tool_Inventory.configure(text='', image=self.Import_File_image, width=5)
            self.B_Calibration_Control_Data.configure(text='', image=self.Import_File_image, width=5)
            self.B_PE_AV_TASK_MONITORING.configure(text='', image=self.Import_File_image, width=5)
            self.B_Aircraft_Registration_Table_Full.configure(text='', image=self.Import_File_image, width=5)

        self.B_ClearImported = ttk.Button(self.Import, command=clearimported)
        self.B_ClearImported.configure(text='Clear')
        self.B_ClearImported.grid(column=0, columnspan=2, ipadx=10, pady=10, row=8)

        self.Label8 = ttk.Label(self.Import)
        self.Label8.configure(text='AITAR')
        self.Label8.grid(column=0, padx=50, pady=20, row=1)
        self.Label9 = ttk.Label(self.Import)
        self.Label9.configure(text='Tooling Load History')
        self.Label9.grid(column=0, pady=20, row=2)
        self.Label10 = ttk.Label(self.Import)
        self.Label10.configure(text='AMM_MPD Data')
        self.Label10.grid(column=0, pady=20, row=3)
        self.Label11 = ttk.Label(self.Import)
        self.Label11.configure(text='Tool Inventory')
        self.Label11.grid(column=0, pady=20, row=4)
        self.Label12 = ttk.Label(self.Import)
        self.Label12.configure(text='Calibration Control Data')
        self.Label12.grid(column=0, pady=20, row=5)
        self.Label13 = ttk.Label(self.Import)
        self.Label13.configure(text='PE AV Task Monitoring')
        self.Label13.grid(column=0, pady=20, row=6)
        self.Label14 = ttk.Label(self.Import)
        self.Label14.configure(text='Aircraft Registration Table Full')
        self.Label14.grid(column=0, ipadx=10, padx=10, pady=20, row=7)

        self.B_AITAR = ttk.Button(self.Import, command=import_AITAR)
        self.B_AITAR.configure(image=self.Import_File_image, width=5)
        self.B_AITAR.grid(column=1, ipadx=70, ipady=10, padx=10, pady=10, row=1, sticky="w")
        self.B_Tooling_Load_History = ttk.Button(self.Import, command=import_Tooling_Load_History)
        self.B_Tooling_Load_History.configure(image=self.Import_File_image, width=5)
        self.B_Tooling_Load_History.grid(column=1, ipadx=70, ipady=10, padx=10, pady=10, row=2, sticky="w")
        self.B_AMM_MPD_Data = ttk.Button(self.Import, command=import_AMM_MPD_Data)
        self.B_AMM_MPD_Data.configure(image=self.Import_File_image, width=5)
        self.B_AMM_MPD_Data.grid(column=1, ipadx=70, ipady=10, padx=10, pady=10, row=3, sticky="w")
        self.B_Tool_Inventory = ttk.Button(self.Import, command=import_Tool_Inventory)
        self.B_Tool_Inventory.configure(image=self.Import_File_image, width=5)
        self.B_Tool_Inventory.grid(column=1, ipadx=70, ipady=10, padx=10, pady=10, row=4, sticky="w")
        self.B_Calibration_Control_Data = ttk.Button(self.Import, command=import_Calibration_Control_Data)
        self.B_Calibration_Control_Data.configure(image=self.Import_File_image, width=5)
        self.B_Calibration_Control_Data.grid(column=1, ipadx=70, ipady=10, padx=10, pady=10, row=5, sticky="w")
        self.B_PE_AV_TASK_MONITORING = ttk.Button(self.Import, command=import_PE_AV_TASK_MONITORING)
        self.B_PE_AV_TASK_MONITORING.configure(image=self.Import_File_image, width=5)
        self.B_PE_AV_TASK_MONITORING.grid(column=1, ipadx=70, ipady=10, padx=10, pady=10, row=6, sticky="w")
        self.B_Aircraft_Registration_Table_Full = ttk.Button(self.Import, command=import_Aircraft_Registration_Table_Full)
        self.B_Aircraft_Registration_Table_Full.configure(image=self.Import_File_image, width=5)
        self.B_Aircraft_Registration_Table_Full.grid(column=1, ipadx=70, ipady=10, padx=10, pady=10, row=7, sticky="w")

        label17 = ttk.Label(self.Import)
        label17.configure(takefocus=False, text='Import Files :')
        label17.grid(column=0, row=0)
        self.Import.grid(column=1, row=0)

        self.mainwindow = self.Name

    def run(self):
        self.mainwindow.mainloop()


if __name__ == "__main__":
    app = NewprojectApp()
    app.run()
