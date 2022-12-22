##########################################################################################
# Filename:     main.py
# For:          Follow_Up_Creation_Tool
# Author:       Spyros Acheimastos (acheimastos@althom.eu)
# Date:         15/12/2022
##########################################################################################

#   Useful for multithreading: https://www.pythonguis.com/tutorials/multithreading-pyqt-applications-qthreadpool/

import os, sys
from PySide2.QtCore import *
from PySide2.QtGui import *
from PySide2.QtWidgets import *
from PySide2.QtUiTools import QUiLoader

# For Multithreading
from bin.multi import Worker
from PySide2.QtCore import QThreadPool

# Main Running Functions
from bin.fun_run_start import (
    fun_run_0_start, 
    fun_run_1_start, 
    fun_run_2_start, 
    fun_run_3_start, 
    fun_run_8_start,
    fun_run_9_start,
    fun_generate_authors_start,
    fun_generate_msns_start
)


SCRIPT_DIRECTORY = os.path.dirname(os.path.abspath(__file__))


class UiLoader(QUiLoader):
    """UiLoader
    Args:
        QUiLoader ([type]): [description]
    """
    def __init__(self, baseinstance, custom_widgets=None):
        QUiLoader.__init__(self, baseinstance)
        self.baseinstance = baseinstance
        self.customWidgets = custom_widgets

    def createWidget(self, class_name, parent=None, name=''):
        if parent is None and self.baseinstance:
            return self.baseinstance

        else:
            if class_name in self.availableWidgets():
                widget = QUiLoader.createWidget(self, class_name, parent, name)

            else:
                try:
                    widget = self.customWidgets[class_name](parent)

                except (TypeError, KeyError):
                    raise Exception('No custom widget '
                                    + class_name
                                    + ' found in customWidgets param of'
                                    + 'UiLoader __init__.')

            if self.baseinstance:
                setattr(self.baseinstance, name, widget)
            return widget

def load_ui(ui_file, baseinstance=None, custom_widgets=None,
            work_dir=None):
    """load_ui
    Args:
        ui_file ([type]): [description]
        baseinstance ([type], optional): [description]. Defaults to None.
        custom_widgets ([type], optional): [description]. Defaults to None.
        work_dir ([type], optional): [description]. Defaults to None.
    Returns:
        [type]: [description]
    """
    loader = UiLoader(baseinstance, custom_widgets)
    if work_dir is not None:
        loader.setWorkingDirectory(work_dir)
    widget = loader.load(ui_file)
    QMetaObject.connectSlotsByName(widget)
    return widget

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.setWindowIcon(QIcon('ui/althom.png'))

        # For Multithreading
        self.threadpool = QThreadPool()

        # Load UI
        load_ui(os.path.join(SCRIPT_DIRECTORY, 'ui/UI.ui'), self)

        # Add colour to all "Run" buttons
        for btn in ['btn_run_0', 'btn_run_1', 'btn_run_2', 'btn_run_3', 'btn_run_6', 'btn_run_7', 'btn_run_8', 'btn_run_9', 'btn_generate_msns', 'btn_generate_authors']:
            self.findChild(QPushButton, btn).setStyleSheet("background-color: #FFD966")


        ############ CREATE FOLLOW_UP ############
        # Step 1
        self.findChild(QPushButton, 'btn_latest_follow_up').clicked.connect(lambda: self.fun_latest_follow_up())
        self.findChild(QPushButton, 'btn_pseudo_db_1').clicked.connect(lambda: self.fun_pseudo_db_1())
        self.findChild(QPushButton, 'btn_run_1').clicked.connect(lambda: self.fun_run_1())

        # Step 2
        self.findChild(QPushButton, 'btn_json').clicked.connect(lambda: self.fun_json())
        self.findChild(QPushButton, 'btn_mdl').clicked.connect(lambda: self.fun_mdl())
        self.findChild(QPushButton, 'btn_pseudo_db_2').clicked.connect(lambda: self.fun_pseudo_db_2())
        self.findChild(QPushButton, 'btn_run_2').clicked.connect(lambda: self.fun_run_2())
        
        # Step 3
        self.findChild(QPushButton, 'btn_json_2').clicked.connect(lambda: self.fun_json())
        self.findChild(QPushButton, 'btn_mdl_2').clicked.connect(lambda: self.fun_mdl())
        self.findChild(QPushButton, 'btn_pseudo_db_3').clicked.connect(lambda: self.fun_pseudo_db_2())
        self.findChild(QPushButton, 'btn_json_authors').clicked.connect(lambda: self.fun_json_authors())
        self.findChild(QPushButton, 'btn_run_3').clicked.connect(lambda: self.fun_run_3())

        ############ UPDATE FOLLOW_UP ############
        # Step 1
        self.findChild(QPushButton, 'btn_json_3').clicked.connect(lambda: self.fun_json())
        self.findChild(QPushButton, 'btn_mdl_3').clicked.connect(lambda: self.fun_mdl())
        self.findChild(QPushButton, 'btn_pseudo_db_5').clicked.connect(lambda: self.fun_pseudo_db_2())
        self.findChild(QPushButton, 'btn_run_6').clicked.connect(lambda: self.fun_run_2())

        # Step 2
        self.findChild(QPushButton, 'btn_json_4').clicked.connect(lambda: self.fun_json())
        self.findChild(QPushButton, 'btn_mdl_4').clicked.connect(lambda: self.fun_mdl())
        self.findChild(QPushButton, 'btn_pseudo_db_6').clicked.connect(lambda: self.fun_pseudo_db_2())
        self.findChild(QPushButton, 'btn_run_7').clicked.connect(lambda: self.fun_run_7())

        # Step 3
        self.findChild(QPushButton, 'btn_json_5').clicked.connect(lambda: self.fun_json())
        self.findChild(QPushButton, 'btn_json_authors_2').clicked.connect(lambda: self.fun_json_authors())
        self.findChild(QPushButton, 'btn_old_follow_up').clicked.connect(lambda: self.fun_old_follow_up())
        self.findChild(QPushButton, 'btn_new_follow_up').clicked.connect(lambda: self.fun_new_follow_up())
        self.findChild(QPushButton, 'btn_run_8').clicked.connect(lambda: self.fun_run_8())

        ############ EXTRA ############
        # Generate JSON files
        self.findChild(QPushButton, 'btn_generate_authors').clicked.connect(lambda: self.fun_generate_authors())
        self.findChild(QPushButton, 'btn_generate_msns').clicked.connect(lambda: self.fun_generate_msns())

        # (Step 0) Generate a New PseudoDB
        self.findChild(QPushButton, 'btn_one_follow_up').clicked.connect(lambda: self.fun_one_follow_up())
        self.findChild(QPushButton, 'btn_run_0').clicked.connect(lambda: self.fun_run_0())

        # Create ALL-NCs
        self.findChild(QPushButton, 'btn_json_6').clicked.connect(lambda: self.fun_json())
        self.findChild(QPushButton, 'btn_mdl_all_MSNs').clicked.connect(lambda: self.fun_all_mdl())
        self.findChild(QPushButton, 'btn_mdl_rev_MSNs').clicked.connect(lambda: self.fun_rev_mdl())
        self.findChild(QPushButton, 'btn_run_9').clicked.connect(lambda: self.fun_run_9())


    def fun_one_follow_up(self):
        self.filepath_one_follow_up = QFileDialog.getOpenFileName(self, 'Select a Follow-up', SCRIPT_DIRECTORY, 'Excel File (*.xlsx)')[0]

    def fun_latest_follow_up(self):
        self.filepath_latest_follow_up = QFileDialog.getOpenFileName(self, 'Select latest Follow-up', SCRIPT_DIRECTORY, 'Excel File (*.xlsx)')[0]

    def fun_pseudo_db_1(self):
        self.filepath_pseudo_db_1 = QFileDialog.getOpenFileName(self, 'Select latest PseudoDataBase', SCRIPT_DIRECTORY, 'Excel File (*.xlsx)')[0]

    def fun_json(self):
        self.filepath_json = QFileDialog.getOpenFileName(self, 'Select JSON with MSNs', SCRIPT_DIRECTORY, 'JSON (*.json)')[0]
    
    def fun_json_authors(self):
        self.filepath_json_authors = QFileDialog.getOpenFileName(self, 'Select JSON with Authors', SCRIPT_DIRECTORY, 'JSON (*.json)')[0]

    def fun_mdl(self):
        self.filepath_mdl = QFileDialog.getExistingDirectory(self, 'Select folder with MDLs', SCRIPT_DIRECTORY)

    def fun_pseudo_db_2(self):
        self.filepath_pseudo_db_2 = QFileDialog.getOpenFileName(self, 'Select latest PseudoDataBase', SCRIPT_DIRECTORY, 'Excel File (*.xlsx)')[0]

    def fun_old_follow_up(self):
        self.filepath_old_follow_up = QFileDialog.getOpenFileName(self, 'Select OLD Follow-up', SCRIPT_DIRECTORY, 'Excel File (*.xlsx)')[0]
    
    def fun_new_follow_up(self):
        self.filepath_new_follow_up = QFileDialog.getOpenFileName(self, 'Select Temporary Follow-up', SCRIPT_DIRECTORY, 'Excel File (*.xlsx)')[0]

    def fun_all_mdl(self):
        self.filepath_all_mdl = QFileDialog.getExistingDirectory(self, 'Select folder with the Latest MDLs for All MSNs', SCRIPT_DIRECTORY)
    
    def fun_rev_mdl(self):
        self.filepath_rev_mdl = QFileDialog.getExistingDirectory(self, 'Select folder with the MDLs that where incorporated last time for Rev MSNs', SCRIPT_DIRECTORY)
    
    #############################################
    def my_console_update(self, text: str = '', clear: bool = False):
        """
        """
        if clear is True:
            self.my_textBrowser.setText('')
            
        previous_msg = self.my_textBrowser.toPlainText()
        if previous_msg == '':
            msg = text
        else:
            msg = previous_msg + '\n' + text
        self.my_textBrowser.setText(msg)


    def save_result(self, ressult_dict):
        """
        """
        if type(ressult_dict) != dict:
            return

        for key, value in ressult_dict.items():
            setattr(self, key, value)

    #############################################
    #############################################
    def fun_generate_authors(self):
        # Clear before starting
        self.my_console_update(clear=True)

        # Pass the function to execute
        self.worker = Worker(
            fun_generate_authors_start,
            console=True
        )  # Any other args, kwargs are passed to the run function

        # Make Connections
        self.worker.signals.console.connect(self.my_console_update)

        # Execute
        self.threadpool.start(self.worker)

    def fun_generate_msns(self):
        # Clear before starting
        self.my_console_update(clear=True)

        # Pass the function to execute
        self.worker = Worker(
            fun_generate_msns_start,
            console=True
        )  # Any other args, kwargs are passed to the run function

        # Make Connections
        self.worker.signals.console.connect(self.my_console_update)

        # Execute
        self.threadpool.start(self.worker)


    def fun_run_0(self):
        # Initial Checks
        if not hasattr(self, 'filepath_one_follow_up') or self.filepath_one_follow_up == '':
            return self.my_console_update(text='Give a Follow-up file first.', clear=True)

        # Clear before starting
        self.my_console_update(clear=True)

        # Pass the function to execute
        self.worker = Worker(
            fun_run_0_start,
            self.filepath_one_follow_up,
            console=True
        )  # Any other args, kwargs are passed to the run function

        # Make Connections
        # self.worker.signals.result.connect(self.save_result)
        self.worker.signals.console.connect(self.my_console_update)

        # Execute
        self.threadpool.start(self.worker)


    def fun_run_1(self):
        # Initial Checks
        if not hasattr(self, 'filepath_latest_follow_up') or self.filepath_latest_follow_up == '':
            return self.my_console_update(text='Give the Latest Follow-up first.', clear=True)
        if not hasattr(self, 'filepath_pseudo_db_1') or self.filepath_pseudo_db_1 == '':
            return self.my_console_update(text='Give the Latest PseudoDataBase first.', clear=True)

        # Clear before starting
        self.my_console_update(clear=True)

        # Pass the function to execute
        self.worker = Worker(
            fun_run_1_start,
            self.filepath_latest_follow_up,
            self.filepath_pseudo_db_1,
            console=True
        )  # Any other args, kwargs are passed to the run function

        # Make Connections
        # self.worker.signals.result.connect(self.save_result)
        self.worker.signals.console.connect(self.my_console_update)

        # Execute
        self.threadpool.start(self.worker)


    def fun_run_2(self):
        # Initial Checks
        if not hasattr(self, 'filepath_mdl') or self.filepath_mdl == '':
            return self.my_console_update(text='Give the Latest MDL Folder first.', clear=True)
        if not hasattr(self, 'filepath_json') or self.filepath_json == '':
            return self.my_console_update(text='Give the Latest JSON File with MSNs first.', clear=True)
        if not hasattr(self, 'filepath_pseudo_db_2') or self.filepath_pseudo_db_2 == '':
            return self.my_console_update(text='Give the Latest PseudoDataBase first.', clear=True)

        # Clear before starting
        self.my_console_update(clear=True)
        
        # Pass the function to execute
        self.worker = Worker(
            fun_run_2_start,
            self.filepath_json,
            self.filepath_mdl,
            self.filepath_pseudo_db_2,
            console=True
        )  # Any other args, kwargs are passed to the run function

        # Make Connections
        # self.worker.signals.result.connect(self.save_result)
        self.worker.signals.console.connect(self.my_console_update)

        # Execute
        self.threadpool.start(self.worker)


    def fun_run_3(self):
        # Get Revision and excelfilepath
        revision = self.findChild(QLineEdit, 'input_revision').text()
        if revision == '': revision = '99'
        excelfilepath = f'EFW Follow-up R{revision}.xlsx'

        # Initial Checks
        if not hasattr(self, 'filepath_mdl') or self.filepath_mdl == '':
            return self.my_console_update(text='Give the Latest MDL Folder first.', clear=True)
        if not hasattr(self, 'filepath_json') or self.filepath_json == '':
            return self.my_console_update(text='Give the Latest JSON File with MSNs first.', clear=True)
        if not hasattr(self, 'filepath_json_authors') or self.filepath_json_authors == '':
            return self.my_console_update(text='Give the Latest JSON File with Authors first.', clear=True)
        if not hasattr(self, 'filepath_pseudo_db_2') or self.filepath_pseudo_db_2 == '':
            return self.my_console_update(text='Give the Latest PseudoDataBase first.', clear=True)

        # Clear before starting
        self.my_console_update(clear=True)
        
        # Pass the function to execute
        self.worker = Worker(
            fun_run_3_start,
            self.filepath_json,
            self.filepath_mdl,
            self.filepath_pseudo_db_2,
            excelfilepath,
            self.filepath_json_authors,
            console=True
        )  # Any other args, kwargs are passed to the run function

        # Make Connections
        # self.worker.signals.result.connect(self.save_result)
        self.worker.signals.console.connect(self.my_console_update)

        # Execute
        self.threadpool.start(self.worker)

    
    def fun_run_7(self):
        # Set filename of Excel
        excelfilepath = f'EFW Follow-up New-Temporary.xlsx'

        # Initial Checks
        if not hasattr(self, 'filepath_mdl') or self.filepath_mdl == '':
            return self.my_console_update(text='Give the Latest MDL Folder first.', clear=True)
        if not hasattr(self, 'filepath_json') or self.filepath_json == '':
            return self.my_console_update(text='Give the Latest JSON File with MSNs first.', clear=True)
        if not hasattr(self, 'filepath_pseudo_db_2') or self.filepath_pseudo_db_2 == '':
            return self.my_console_update(text='Give the Latest PseudoDataBase first.', clear=True)
    
        # Clear before starting
        self.my_console_update(clear=True)
        
        # Dont add authors
        self.filepath_json_authors = None

        # Pass the function to execute
        self.worker = Worker(
            fun_run_3_start,
            self.filepath_json,
            self.filepath_mdl,
            self.filepath_pseudo_db_2,
            excelfilepath,
            self.filepath_json_authors,
            add_QBs=False,
            console=True
        )  # Any other args, kwargs are passed to the run function

        # Make Connections
        # self.worker.signals.result.connect(self.save_result)
        self.worker.signals.console.connect(self.my_console_update)

        # Execute
        self.threadpool.start(self.worker)


    def fun_run_8(self):
        # Initial Checks
        if not hasattr(self, 'filepath_json') or self.filepath_json == '':
            return self.my_console_update(text='Give the Latest JSON File with MSNs first.', clear=True)
        if not hasattr(self, 'filepath_json_authors') or self.filepath_json_authors == '':
            return self.my_console_update(text='Give the Latest JSON File with Authors first.', clear=True)
        if not hasattr(self, 'filepath_old_follow_up') or self.filepath_old_follow_up == '':
            return self.my_console_update(text='Give the Old Follow-up first.', clear=True)
        if not hasattr(self, 'filepath_new_follow_up') or self.filepath_new_follow_up == '':
            return self.my_console_update(text='Give the New Follow-up first.', clear=True)

        # Clear before starting
        self.my_console_update(clear=True)

        # Set filename of Excel
        excelfilepath = self.filepath_old_follow_up.replace('.xlsx', '_FINAL.xlsx')
        
        # Pass the function to execute
        self.worker = Worker(
            fun_run_8_start,
            self.filepath_json,
            self.filepath_json_authors,
            self.filepath_old_follow_up,
            self.filepath_new_follow_up,
            excelfilepath,
            console=True
        )  # Any other args, kwargs are passed to the run function

        # Make Connections
        # self.worker.signals.result.connect(self.save_result)
        self.worker.signals.console.connect(self.my_console_update)

        # Execute
        self.threadpool.start(self.worker)


    def fun_run_9(self):
        """
        For All-NCs
        """
        # Get Revision and excelfilepath
        revision = self.findChild(QLineEdit, 'input_revision_2').text()
        if revision == '': revision = 'XX'
        # excelfilepath = f'ALL_NCs_R{revision}.xlsx'

        # Initial Checks
        if not hasattr(self, 'filepath_json') or self.filepath_json == '':
            return self.my_console_update(text='Give the Latest JSON File with MSNs first.', clear=True)
        if not hasattr(self, 'filepath_all_mdl') or self.filepath_all_mdl == '':
            return self.my_console_update(text='Give folder with the Latest MDLs for All MSNs first.', clear=True)
        if not hasattr(self, 'filepath_rev_mdl') or self.filepath_rev_mdl == '':
            return self.my_console_update(text='Give folder with the MDLs that where incorporated last time for Rev MSNs first.', clear=True)

        # Clear before starting
        self.my_console_update(clear=True)

        # Pass the function to execute
        self.worker = Worker(
            fun_run_9_start,
            self.filepath_json,
            self.filepath_all_mdl,
            self.filepath_rev_mdl,
            revision,
            console=True
        )  # Any other args, kwargs are passed to the run function

        # Make Connections
        # self.worker.signals.result.connect(self.save_result)
        self.worker.signals.console.connect(self.my_console_update)

        # Execute
        self.threadpool.start(self.worker)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    ret = app.exec_()
    sys.exit()