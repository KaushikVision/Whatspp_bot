import tkinter as tk
from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox, filedialog
from datetime import datetime
import time
import os
import wmi
import threading
import pythoncom
import logging
from constant import Constant

class Sequencer(tk.LabelFrame, logging.Handler):
    running = True

    def current_time(self):
        currentTime = datetime.now()
        return currentTime
    
    def set_check_val_in_log(self, data, line, field, checkVal):
        if(line.__contains__(field)):
            if(data.strip() == 'True'):
                checkVal.set(1)
            else:
                checkVal.set(0)

    def write_check_val_in_log(self, checkVal, field):
        with open("keepAlive.log", "r") as f:
            lines = f.readlines()
        if(checkVal == 1):
            with open("keepAlive.log", "w") as f:
                for line in lines:
                    if not line.startswith(field):
                        f.write(line)
                f.write(f"{field}=True")
                f.write("\n")
        elif(checkVal == 0):
            with open("keepAlive.log", "w") as f:
                for line in lines:
                    if not line.startswith(field):
                        f.write(line)
                f.write(f"{field}=False")
                f.write("\n")

    def checkDatabaseUpdate(self, checkvalue):
        self.write_check_val_in_log(self.firstCheck.get(), 'whatsappCheck')

    def write_entry_data_in_log(self, entryval, fieldName):
        if(entryval):
            with open("keepAlive.log", "r") as f:
                lines = f.readlines()
            with open("keepAlive.log", "w") as f:
                for line in lines:
                    if not line.startswith(fieldName):
                        f.write(line)
                f.write(f"{fieldName}={entryval}")
                f.write("\n")

    def callbackFocus(self, waitField):
        self.write_entry_data_in_log(self.firstEntry.get(), 'whatsappEntry')

    def on_closing(self):
        root.destroy()
        pythoncom.CoInitialize()
        w = wmi.WMI()
        mainWindowStatus = False
        processName = 'whatsapp_seq.exe'
        for process in w.Win32_process():
            if(processName == str(process.Name)):
                mainWindowStatus = True
                break
        if(mainWindowStatus):
            os.system(f'taskkill /f /im whatsapp_seq.exe')

    

    def __init__(self, *args, **kwargs):
        self.start_running = False
        checkLogFile = os.path.exists('sequencer.log')
        if(checkLogFile):
            os.remove('sequencer.log')

        logging.Handler.__init__(self)
        logging.basicConfig(
            filename='sequencer.log',
            level=logging.INFO, 
            format='%(asctime)s - %(levelname)s - %(message)s')
        # logging.getLogger().setLevel(logging.DEBUG)

        root.protocol("WM_DELETE_WINDOW", self.on_closing)
        tk.LabelFrame.__init__(self, *args, **kwargs)
        try:
            self.firstCheck = tk.IntVar()

            self.CURRENT_RUNNING_PROCESS = tk.StringVar()
            self.CURRENT_RUNNING_PROCESS.set('')

            self.NEXT_PROCESS = tk.StringVar()
            self.NEXT_PROCESS.set('')

            self.FIRST_TIME_INTERVAL = tk.BooleanVar()
            self.FIRST_TIME_INTERVAL.set(False)

            self.FUNCTION01_LAST_UPDATE_TIME = tk.IntVar()
            self.FUNCTION01_LAST_UPDATE_TIME.set(0)

            self.firstEntry = tk.StringVar()

            self.whatsapp_index_edit01 = tk.StringVar()
            self.whatsapp_index_edit01.set('')

            self.whatsapp_final_path_edit01 = tk.StringVar()
            self.whatsapp_final_path_edit01.set('')

            self.firstProcessCon = False

            checkFile = os.path.exists('keepAlive.log')
            if(not checkFile):
                with open("keepAlive.log", "w") as f:
                    f.close()
            else:
                with open("keepAlive.log", "r") as f:
                    lines = f.readlines()
                    for line in lines:
                        if(line):
                            data = line.split("=")[1]
                            self.set_check_val_in_log(data, line, 'whatsappCheck', self.firstCheck)
                            
                            if(line.__contains__("whatsappEntry")):
                                self.firstEntry.set(data.strip())
                            elif(line.__contains__("whatsappRootPath")):
                                Constant.Whatsapp_excel_file = data.strip()
                                self.whatsapp_index_edit01.set(data.strip())
                            elif(line.__contains__("whatsappFinalPath")):
                                Constant.Whatsapp_final_path = data.strip()
                                self.whatsapp_final_path_edit01.set(data.strip())
                                
                

            self.name_label1 = Label(self, text='Whatsapp Send Messages', anchor="w", font=('calibri', 15))
            self.active_cb1 = Checkbutton(self, takefocus = 0,onvalue=True, offvalue=False, variable=self.firstCheck, command = lambda check_data = "whatsappCheck" :  self.checkDatabaseUpdate(check_data))

            waitField1 = Entry(self, textvariable=self.firstEntry, font = ('calibri', 15, ''), width=8)
            waitField1.bind('<FocusOut>', lambda x='waitField1': self.callbackFocus('waitField1'))

            self.editButton1 = Button(self, text="Edit", style='W.TButton', command=self.edit01)

            self.runButton1 = Button(self, text="Run", style='W.TButton', command=self.run01)

            self.startButton = Button(self, text="START",style='W.TButton', command=self.start)
            self.startButton.place(x=320, y=120)

            self.stopButton = Button(self, text="STOP",style='W.TButton', command=self.stop)
            self.stopButton.place(x=430, y=120)

            self.resumeButton = Button(self, text="RESUME",style='W.TButton', command=self.resume)
            self.resumeButton.place(x=540, y=120)

            self.data = [
                    [1,   self.name_label1, self.active_cb1, waitField1, self.editButton1, self.runButton1],
            ]

            self.grid_columnconfigure(2, weight=3)
            Label(self, text="Sr No.", anchor="w", font=('calibri', 16, ''), foreground='blue').grid(row=0, column=0, sticky="ew", padx=10, pady=5 )
            Checkbutton(self, takefocus = 0,onvalue=True, offvalue=False).grid(row=0, column=1, sticky="ew", padx=10, pady=5,)
            # Checkbutton(self, takefocus = 0,onvalue=True, offvalue=False, variable=self.selectAllCheck, command=self.checkAll).grid(row=0, column=1, sticky="ew", padx=10, pady=5,)
            Label(self, text="Process Name", anchor="w", font=('calibri', 16, ''), foreground='blue').grid(row=0, column=2, sticky="ew", padx=10, pady=5,)
            Label(self, text="Action", anchor="w", font=('calibri', 16, ''), foreground='blue').grid(row=0, column=3, sticky="ew", padx=10, pady=5,)
            Label(self, text="Interval (Min)", anchor="w", font=('calibri', 16, ''), foreground='blue').grid(row=0, column=4, sticky="ew", padx=10, pady=5,)
            Label(self, text="Settings", anchor="w", font=('calibri', 16, ''), foreground='blue').grid(row=0, column=5, sticky="ew", padx=10, pady=5,)
            style = Style()
            style.configure('W.TButton', font =
                ('calibri', 15, 'normal'),
                foreground = 'black', width=8)

            row = 1
            for (runFuncName, name, active, waitEntry, editbuttons, runButtons) in self.data:
                runFuncName_label = Label(self, text=str(runFuncName), anchor="w", font=('calibri', 15))
                
                name_label = name
                self.action_button = runButtons
                self.active_cb = active
                waitField = waitEntry
                edit_button = editbuttons
                
                runFuncName_label.grid(row=row, column=0, sticky="ew", padx=10, pady=5)
                self.active_cb.grid(row=row, column=1, sticky="ew", padx=10, pady=5)
                name_label.grid(row=row, column=2, sticky="ew", padx=10, pady=5)
                self.action_button.grid(row=row, column=3, sticky="ew", padx=10, pady=5)
                waitField.grid(row=row, column=4, sticky="ew", padx=10, pady=5)
                edit_button.grid(row=row, column=5, sticky="ew", padx=10, pady=5)     
                row += 1

        except Exception as error:
            print(error)
            logging.info('whatsapp initialize error')


    def statusCheck(self):
        try:
            if(Constant.Current_running_process == 'Whatsapp'):
                disabledArray = [self.runButton1, self.editButton1, self.startButton, self.stopButton, self.resumeButton]
                for disableBtn in disabledArray:
                    disableBtn.configure(state=DISABLED)
                    disableBtn.update()
            elif(self.CURRENT_RUNNING_PROCESS.get() == '' or Constant.Current_running_process == ''):
                enabledArray = [self.runButton1, self.editButton1, self.startButton, self.stopButton, self.resumeButton]
                for enabledeBtn in enabledArray:
                    enabledeBtn.configure(state=NORMAL)
                    enabledeBtn.update()
        except Exception as error:
            logging.info("Exe Function statusCheck button enable - disable error")
    
    def run_check_process(self, curr_process):
        process = {}
        process['Whatsapp'] = self.firstProcessCon
        
        check = False
        new_process = False
        for key, value in process.items():
            if(key == curr_process):
                check = True
            if(check and value):
                new_process = True
                self.NEXT_PROCESS.set(key)
            if(new_process):
                break
            
        if(new_process == False):         
            for key, value in process.items():
                if(value):
                    self.NEXT_PROCESS.set(key)
                    break

    def WHATSAPP(self):
        from whatsapp_process import WhatsappProcess
        print('enter in whatsapp function')
        try:
            self.name_label1.config(text='Whatsapp Send Messages (Running)', font=('calibri', 15, 'bold'))
            self.name_label1.update()
            self.CURRENT_RUNNING_PROCESS.set('Whatsapp')
            whatsapp_class = WhatsappProcess()
            self.whatsappFuncThread=threading.Thread(target=whatsapp_class.whatsapp_auto_04)
            self.whatsappFuncThread.start()
            
            self.whatsapp_thread()
            
        except Exception as error:
            logging.info(f"Exe Function WHATSAPP error : {error}")

    def whatsapp_thread(self):
        if(not(self.whatsappFuncThread.is_alive())):
            self.name_label1.config(text='Whatsapp Send Messages', font=('calibri', 15, 'normal'))
            self.name_label1.update()
            self.CURRENT_RUNNING_PROCESS.set('')
            Constant.Current_running_process = ''
            self.statusCheck()
            self.run_check_process('Whatsapp')
        else:
            root.after(1000, self.whatsapp_thread)

    def okClickBtn(self, edit_var):
        if(edit_var == 'WHATSAPP'):
            if(self.whatsappIndexInput.get() == ''):
                messagebox.showwarning('Warning', 'Whtsapp Index File Path Does Not Exists.')
            elif(self.whatsappFinalInput.get() == ''):
                messagebox.showwarning('Warning', 'Whtsapp Final File Path Does Not Exists.')
            else:
                self.childWin.destroy()
        
    
    def cancelClickBtn(self):
        try:
            self.childWin.destroy()
        except Exception as error:
            logging.info("Function: cancelClickBtn (window destry time error)")

    def edit01(self):
        try:
            self.childWin = Toplevel(self)
            self.childWin.grab_set()
            self.childWin.geometry('490x220')
            self.childWin.resizable(0,0)
            self.childWin.title('Whatsapp Send Messages')

            okButton = Button(self.childWin, text='Save', width=10,  command=lambda edit_var='WHATSAPP': self.okClickBtn(edit_var))
            okButton.place(x=170, y=170, height=33)
            cancelButton = Button(self.childWin, text='Cancel', width=10, command=self.cancelClickBtn)
            cancelButton.place(x=260, y=170, height=33)
            
            # Whatsapp Index FOLDER PATH
            whatsappIndexFolder = Label(self.childWin, text="Whatsapp Index File Path", font=('calibri', 13, 'bold')).place(x=10, y=10)
            self.whatsappIndexInput = Entry(self.childWin,textvariable=self.whatsapp_index_edit01, font = ('calibri', 15, ''), width=40)
            self.whatsappIndexInput.place(x=10, y=45)

            s = Style()
            s.configure('my.TButton', font=('calibri', 12, 'bold'))   
            
            cleanpathBtn = Button(self.childWin, text='...', command=lambda varPath='WHATSAPP_INDEX_PATH': self.rootButton(varPath), style='my.TButton')
            cleanpathBtn.place(x=420, y=45, width=55,height=33)

            # Whatsapp Re-allocate Path
            whatsappFinalFolder = Label(self.childWin, text="One message file to send and then move SS", font=('calibri', 13, 'bold')).place(x=10, y=90)
            self.whatsappFinalInput = Entry(self.childWin,textvariable=self.whatsapp_final_path_edit01, font = ('calibri', 15, ''), width=40)
            self.whatsappFinalInput.place(x=10, y=125)
            
            finalpathBtn = Button(self.childWin, text='...', command=lambda varPath='WHATSAPP_FINAL_PATH': self.rootButton(varPath), style='my.TButton')
            finalpathBtn.place(x=420, y=125, width=55,height=33)

        except Exception as error:
            logging.info("Whatsapp Send Messages Process Edit button EXE error")

    def setLogDataConstant(self, line, data):
        if(line.__contains__("whatsappRootPath")):
            Constant.Whatsapp_excel_file = data.strip()
            self.whatsapp_index_edit01.set(data.strip())
        elif(line.__contains__("whatsappFinalPath")):
            Constant.Whatsapp_final_path = data.strip()
            self.whatsapp_final_path_edit01.set(data.strip())

    def write_paths_in_log(self, fieldName, fileDialogType, editPath, varPath, rootbuttonPath):
        if(rootbuttonPath == varPath):
            if(fileDialogType == 'askopenfilename'):
                filename = filedialog.askopenfilename()
            else:
                filename = filedialog.askdirectory()

            if(filename):
                self.setLogDataConstant(fieldName, filename)
                editPath.set(filename)
                with open("keepAlive.log", "r") as f:
                    lines = f.readlines()  
                with open("keepAlive.log", "w") as f:
                    for line in lines:
                        if not line.startswith(fieldName):
                            f.write(line)
                    f.write(f"{fieldName}={filename}")
                    f.write("\n")

    def rootButton(self, rootbuttonPath):
        try:
            self.write_paths_in_log('whatsappRootPath', 'askopenfilename', self.whatsapp_index_edit01, 'WHATSAPP_INDEX_PATH', rootbuttonPath)
            self.write_paths_in_log('whatsappFinalPath', 'askopenfolder', self.whatsapp_final_path_edit01, 'WHATSAPP_FINAL_PATH', rootbuttonPath)
        except Exception as error:
            print(error)
            logging.info("Function rootButton (Path selction time error)")

    def run01(self):
        try:
            self.start_running = True
            whatsapp_index_path = self.whatsapp_index_edit01.get()
            whatsapp_final_path = self.whatsapp_final_path_edit01.get()
            if(str(whatsapp_index_path) == ''):
                logging.info("Whatsapp Index File Path Does not Exists.")
            elif(str(whatsapp_final_path) == ''):
                logging.info("Whatsapp Final File Path Does not Exists.")
            else:
                if(self.firstCheck.get() == 0):
                    logging.info("Whatsapp Please Select Anyone Checkbox.")
                else:
                    self.CURRENT_RUNNING_PROCESS.set('Whatsapp')
                    Constant.Current_running_process = 'Whatsapp'
                    self.statusCheck()
                    self.WHATSAPP()

        except Exception as error:
            print(error)
            logging.info("Whatsapp Send Messages - Run Button Click time error.")

    def start_sub_code(self):
        try:
            if(self.start_running):
                disabledArray = [self.runButton1, self.editButton1]
                for disableBtn in disabledArray:
                    disableBtn.configure(state=DISABLED)
                    disableBtn.update()

                if(self.firstCheck.get() == 1 and (self.CURRENT_RUNNING_PROCESS.get() != 'STOP' and self.CURRENT_RUNNING_PROCESS.get() == '' and self.NEXT_PROCESS.get() == 'Whatsapp') and self.firstProcessCon):

                    if(self.FIRST_TIME_INTERVAL.get()):
                        fourEntryTime = self.firstEntry.get()
                        fourEntryTime = int(fourEntryTime)*60
                        if((float(self.FUNCTION01_LAST_UPDATE_TIME.get()) + fourEntryTime) <= time.time()):
                            self.WHATSAPP()
                            self.FUNCTION01_LAST_UPDATE_TIME.set(time.time())
                        else:
                            self.run_check_process('Whatsapp')
                    else:
                        self.FIRST_TIME_INTERVAL.set(True)
                        self.FUNCTION01_LAST_UPDATE_TIME.set(time.time())
                        self.WHATSAPP()

            if(self.CURRENT_RUNNING_PROCESS.get() != 'STOP'):
                root.after(1000, self.start_sub_code)

        except Exception as error:
            print('****trace back***')
            import traceback
            print(traceback.format_exc())
            logging.info("Function: start_sub_code (Start Button click time error).")

    def start(self):
        try:
            self.start_running = True
            global running
            running = True
            # self.CURRENT_RUNNING_PROCESS.set('')

            if(self.firstCheck.get() == 1):
                self.firstProcessCon = True     
            
            if(self.firstCheck.get() == 0):
                logging.info("First Process - Please Select Anyone Checkbox.")
            else:
                if(not (self.firstEntry.get()).isdigit() and (self.firstCheck.get() == 1)):
                    self.firstProcessCon = False
                    messagebox.showwarning('Warning', "Whatsapp - Interval Field Required Only Digits.")

                self.startThread=threading.Thread(target=self.Start_thread)
                self.startThread.start()
        except Exception as error:
            logging.info("Function: start (Start Button click time error)")

    def Start_thread(self):
        self.run_check_process('Whatsapp')
        print('Start_thread')
        self.start_sub_code()

    def stop(self):
        from whatsapp_process import WhatsappProcess
        whatsapp_process = WhatsappProcess()
        try:
            whatsapp_process.running_whatsapp = False
            global running
            running = False
            if(self.CURRENT_RUNNING_PROCESS.get() != ''):
                self.CURRENT_RUNNING_PROCESS.set('STOP')

            if(self.CURRENT_RUNNING_PROCESS.get() == ''):
                disabledArray = [self.runButton1, self.editButton1, self.startButton, self.resumeButton]
                for disableBtn in disabledArray:
                    disableBtn["state"] = "normal"
                    disableBtn.update()

            self.start_running = False

        except Exception as error:
            logging.info("Function: stop (Stop Button click time error)")

    def resume(self):
        try:
            if(self.CURRENT_RUNNING_PROCESS.get() == '' or self.CURRENT_RUNNING_PROCESS.get() == 'STOP'):
                self.resumeThread=threading.Thread(target=self.start)
                self.resumeThread.start()
        except Exception as error:
            logging.info("Function: resume (Resume Button click time error)")

if __name__ == "__main__":
    root = tk.Tk()
    Sequencer(root, text="", font=("Arial", 10)).pack(side="top", fill="both", expand=True, padx=10, pady=10)
    root.title('Mepani - Whatsapp Sequencer')
    root.geometry('990x200')
    root.resizable(0,0)
    root.mainloop()
