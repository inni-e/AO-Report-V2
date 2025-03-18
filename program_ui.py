# UI for running program

# UI required packages
import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtCore import QThread
from PyQt5.QtCore import pyqtSignal
from pandasModel import DataFrameModel

from UIMainWindow import Ui_MainWindow

# Program modules
from data_preprocessing_SID import DataProcess
from missing_infoSID import MissingExcels
from report_creationSID import ReportCreation

# Send Email Function
from send_email import send_email

# Other required packages
import os
import pandas as pd
from distutils.dir_util import copy_tree
from threading import Thread
from datetime import datetime

class MainWindow():
    def __init__(self):
        self.main_win = QMainWindow()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self.main_win) 
        
        # Grab folders that need to be copied over to new directory
        self.pdfresources_folder = os.getcwd() + '\\pdf_resources'
        self.percentile_folder = os.getcwd() + '\\percentile_bands'
        
        
        ######################### Connecting UI elements (Backend functionality) ##########################
        
        ######## Exam Settings and Aggregate Data ####################
        self.ui.aggdata_button.setEnabled(False)
        self.ui.rollid_checkbox.setEnabled(False)
        self.ui.sort_cb.setEnabled(False)
        self.ui.missingdata_btn.setEnabled(False)
        self.ui.missing_email_btn.setEnabled(False)
        self.ui.batch_send.setChecked(True)
        self.ui.stackedWidget.setCurrentWidget(self.ui.batch_page)
        self.ui.sendall_btn.setEnabled(False)
        self.ui.send_btn.setEnabled(False)
        self.ui.send_btn_2.setEnabled(False)
        self.ui.console_window.setReadOnly(True)
        self.ui.progressBar.hide()
    
        
        self.ui.browseButton.clicked.connect(self.browse_dir)
        
        # Detect updates to test type and test number to update label
        self.update_name()
        self.ui.testtype_combo.currentTextChanged.connect(self.update_name)
        self.ui.testnum_spin.valueChanged.connect(self.update_name)
        
        # confirmation check 
        self.ui.confirmation_check.stateChanged.connect(self.activate_aggdata)
        
        # select rollID 
        self.ui.selectrollid_btn.clicked.connect(self.select_rollid)
        self.ui.rollid_checkbox.clicked.connect(self.activate_aggdata)
        
        # aggdata button
        self.ui.aggdata_button.clicked.connect(self.aggdata_btn)
        
        # Sort button
        self.ui.sort_cb.clicked.connect(self.sort_dataframe)
        
        # Missing Emails and Missing Data
        self.ui.missing_email_btn.clicked.connect(self.gen_missing_emails)
        self.ui.missingdata_btn.clicked.connect(self.gen_missing_data)
        
        # Stacked Widget Selection buttons
        self.ui.batch_send.clicked.connect(self.select_page)
        self.ui.individual_send.clicked.connect(self.select_page)
        
        # Different Send Email Buttons
        self.ui.sendall_btn.clicked.connect(lambda: self.send_batch('all'))
        self.ui.send_btn.clicked.connect(lambda: self.send_batch('from'))
        self.ui.send_btn_2.clicked.connect(self.partial_batch)
        
        
        
        
        
        
    def show(self):
        self.main_win.show()
        
    
    ######################## Backend Functions ######################
    def browse_dir(self):
        self.ui.directory_browser.clear()
        folderpath = QFileDialog.getExistingDirectory(self.main_win, 'Select Folder')
        self.ui.directory_browser.append(folderpath)
        
        # Change the directory of the program
        try:
            os.chdir(folderpath)

            # new
            self.check_q_files()
            
            if not os.path.exists('pdf_resources'):
                os.mkdir('pdf_resources')
                copy_tree(self.pdfresources_folder, 'pdf_resources')
            
            if not os.path.exists('percentile_bands'):
                os.mkdir('percentile_bands')
                copy_tree(self.percentile_folder, 'percentile_bands')
            
        except:
            self.ui.directory_browser.append('')
    
    # new
    def check_q_files(self):
        required_files = ["thinking_q_types.xlsx", "reading_q_types.xlsx", "maths_q_types.xlsx"]
        missing_files = []

        for file in required_files:
            if not os.path.exists(file):
                missing_files.append(file)

        if missing_files:
            missing_str = ", ".join(missing_files)
            self.console_message("Missing question type files: " + missing_str)
        else:
            self.console_message("All question type files are present.")
    
    def update_name(self):
        if self.ui.testtype_combo.currentText() == 'Selective':
            self.ui.testname_txt.setPlainText('Selective Trial Test Course {}'.format(self.ui.testnum_spin.value()))
        elif self.ui.testtype_combo.currentText() == 'OC':
            self.ui.testname_txt.setPlainText('OC Trial Test Course {}'.format(self.ui.testnum_spin.value()))
        else:
            self.ui.testname_txt.setPlainText('WEMT{} Term ___ Test'.format(self.ui.testnum_spin.value()))
    
    
    def select_rollid(self):
        folderpath, _ = QFileDialog.getOpenFileName(self.main_win, 'Select File')
        if folderpath != '':
            self.rollid = pd.read_excel(folderpath)
            self.ui.rollid_checkbox.setEnabled(True)
            
            self.console_message('Roll ID successfully loaded.')
        
    
    def activate_aggdata(self):
        if self.ui.confirmation_check.isChecked() and self.ui.rollid_checkbox.isChecked():
            self.ui.aggdata_button.setEnabled(True)
        else:
            self.ui.aggdata_button.setEnabled(False)

    
    def aggdata_btn(self):
        # Get test type
        if self.ui.testtype_combo.currentText() == 'Selective':
            self.test_type = 'sttc'
        elif self.ui.testtype_combo.currentText() == 'OC':
            self.test_type = 'octt'
        elif self.ui.testtype_combo.currentText() == 'WEMT OC':
            self.test_type = 'wemtoc'
        else:
            self.test_type = 'wemtsel'
        
        # Get data format
        if self.ui.examformat_combo.currentText() == 'Hybrid':
            self.data_format = 'hybrid'
        elif self.ui.examformat_combo.currentText() == 'Online Only':
            self.data_format = 'online'
        else:
            self.data_format = 'inperson'
            
        # Get test number
        self.test_no = str(self.ui.testnum_spin.value())
        
        # Get test name
        self.test_name = self.ui.testname_txt.toPlainText()
        
        
        # Data preprocessing
        try:
            files = DataProcess(self.test_type, self.test_no, self.data_format)
        except Exception as e:
            self.console_message(str(e))
            return()
        
        try:
            files.combine()
        except Exception as e:
            self.console_message(str(e))
            return()
        
        if self.test_type == 'octt':
            try:
                self.reading_combined, self.maths_combined, self.thinking_combined = files.diagnosis(output = True)
            except Exception as e:
                self.console_message(str(e))
                return()
            
        else:
            try:
                self.reading_combined, self.maths_combined, self.thinking_combined, self.writing_combined = files.diagnosis(output = True)
            except Exception as e:
                self.console_message(str(e))
                return()
        
        # Aggregate Data
        if self.test_type == 'octt':
            self.report_creation = ReportCreation(self.test_type, self.test_no, self.rollid, self.reading_combined, 
                                            self.maths_combined, self.thinking_combined, data_type = self.data_format)
        else:
            self.report_creation = ReportCreation(self.test_type, self.test_no, self.rollid, self.reading_combined, 
                                            self.maths_combined, self.thinking_combined, self.writing_combined, data_type = self.data_format)
            
        self.report_creation.prepare()
        self.agg_data = self.report_creation.aggregate_data()
    
        # Add agg data to table widget
        self.model = DataFrameModel(self.agg_data)
        self.ui.tableView.setModel(self.model)
        
        # Enable sort button
        self.ui.sort_cb.setEnabled(True)
        
        # Enable Send email buttons
        self.ui.sendall_btn.setEnabled(True)
        self.ui.send_btn.setEnabled(True)
        self.ui.send_btn_2.setEnabled(True)
        
        # Re-enable missing excel buttons and prepare class
        self.ui.missingdata_btn.setEnabled(True)
        self.ui.missing_email_btn.setEnabled(True)
        
        if self.test_type == 'octt':
            self.missing_info = MissingExcels(self.test_type, self.test_no, self.rollid, self.reading_combined, self.maths_combined, 
                                              self.thinking_combined)
        else:
            self.missing_info = MissingExcels(self.test_type, self.test_no, self.rollid, self.reading_combined, self.maths_combined, 
                                              self.thinking_combined, self.writing_combined)
        
        
    def sort_dataframe(self):
        if self.ui.sort_cb.isChecked():
            self.ui.tableView.setModel(DataFrameModel(self.agg_data.sort_values('Name').reset_index(drop = False)))
        else:
            self.ui.tableView.setModel(self.model)
    
    def gen_missing_emails(self):
        missing_email = self.missing_info.missing_emails()
        missing_email.to_excel('emails_' + self.test_type + self.test_no + '.xlsx')
        self.console_message('Missing Emails Excel has been created.')
        
    def gen_missing_data(self):
        missing_data = self.missing_info.missing_data()
        missing_data.to_excel('missing_' + self.test_type + self.test_no + '.xlsx')
        self.console_message('Missing Data Excel has been created.')
    
    def select_page(self):
        if self.ui.batch_send.isChecked():
            self.ui.stackedWidget.setCurrentWidget(self.ui.batch_page)
        elif self.ui.individual_send.isChecked():
            self.ui.stackedWidget.setCurrentWidget(self.ui.ind_page)

        
    def send_batch(self, type):
        self.ui.progressBar.show()
        if type == 'all':
            starting_index = 0
            self.ui.progressBar.setValue(0)
            self.ui.progressBar.setRange(0, len(self.agg_data))
            
        elif type == 'from':
            starting_index = self.ui.spinBox_2.value()
            self.ui.progressBar.setValue(0)
            self.ui.progressBar.setRange(0, len(self.agg_data)-starting_index)
        
        # Activate worker
        self.worker = WorkerThreadBatch(self.report_creation, self.agg_data, self.test_name, starting_index)
        self.worker.start()
        self.worker.update_progress.connect(self.update_progressBar)
        self.worker.pdf_finished.connect(self.console_message)
        self.worker.email_finished.connect(self.console_message)
        self.worker.finished.connect(self.hide_bar)


        
    def partial_batch(self):
        if self.ui.indices_txt.toPlainText() != '':
            required_ind = [int(x) for x in self.ui.indices_txt.toPlainText().split(',')]
            self.ui.progressBar.setRange(0, len(required_ind))
            self.ui.progressBar.setValue(0)
            self.ui.progressBar.show()
            
            # Activate workers
            self.worker = WorkerThreadPartial(self.report_creation, self.agg_data, self.test_name, required_ind)
            self.worker.start()
            self.worker.update_progress.connect(self.update_progressBar)
            self.worker.completed.connect(self.console_message)
            self.worker.finished.connect(self.hide_bar)

    
    def update_progressBar(self, val):
        self.ui.progressBar.setValue(val+1)
        
    def hide_bar(self):
        self.ui.progressBar.hide()
    
    
    def console_message(self, message):
        time = '({}): '.format(datetime.now().strftime('%H:%M'))
        self.ui.console_window.appendPlainText(time + message)
        
        

# Worker Threads for email sending
class WorkerThreadBatch(QThread):
    # Signals to emit 
    update_progress = pyqtSignal(int)
    pdf_finished = pyqtSignal(str)
    email_finished = pyqtSignal(str)
    
    def __init__(self, report_creation, agg_data, test_name, starting_index):
        super(QThread, self).__init__()
        
        #  Agg data and test name
        self.report_creation = report_creation
        self.agg_data = agg_data
        self.test_name = test_name
        self.starting_index = starting_index
    
    
    def run(self):
        # Send Emails
        threads = []
            
        for i in range(self.starting_index, len(self.agg_data)):
            if i < self.report_creation.num_complete:
                self.report_creation.complete_pdf(i, self.test_name)
                if self.agg_data.email[i] != 'Missing':
                    thread = Thread(target = send_email, args = (self.agg_data.email[i], self.test_name, self.agg_data.Name[i]))
                    threads.append(thread)
                    thread.start()
            else:
                self.report_creation.incomplete_pdf(i, self.test_name)
                if self.agg_data.email[i] != 'Missing':
                    thread = Thread(target = send_email, args = (self.agg_data.email[i], self.test_name, self.agg_data.Name[i]))
                    threads.append(thread)
                    thread.start()
            
            # Update progress bar
            self.update_progress.emit(i-self.starting_index)
        
        self.pdf_finished.emit('All PDFs have been generated. Remaining emails will be delivered shortly.')
        [t.join() for t in threads]
        self.email_finished.emit('All emails have been delivered!')




class WorkerThreadPartial(QThread):
    # Signals to emit
    update_progress = pyqtSignal(int)
    completed = pyqtSignal(str)
    
    def __init__(self, report_creation, agg_data, test_name, required_ind):
        super(QThread, self).__init__()
        self.report_creation = report_creation
        self.agg_data = agg_data
        self.test_name = test_name
        self.required_ind = required_ind
        
    
    def run(self):    
        for i in range(len(self.required_ind)):
            ind = self.required_ind[i]
            if ind < self.report_creation.num_complete:
                self.report_creation.complete_pdf(ind, self.test_name)
                if self.agg_data.email[ind] != 'Missing':
                    send_email(self.agg_data.email[ind], self.test_name, self.agg_data.Name[ind])
            else:
                self.report_creation.incomplete_pdf(ind, self.test_name)
                if self.agg_data.email[ind] != 'Missing':
                    send_email(self.agg_data.email[ind], self.test_name, self.agg_data.Name[ind])
        
            # update progress bar
            self.update_progress.emit(i)
        
        
        self.completed.emit('All emails have been sent!')

    

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_win = MainWindow()
    main_win.show()
    
    # Exit cleanly out of the application
    sys.exit(app.exec_())