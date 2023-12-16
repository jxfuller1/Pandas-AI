import sys
import shutil

import os
from PyQt6.QtWidgets import QApplication, QMainWindow, QTableView, QVBoxLayout, QTableWidget, QWidget, QSizePolicy, \
    QPushButton, QLabel, QLineEdit, QTextEdit, QHBoxLayout, QTableWidgetItem, QMessageBox, QGroupBox, QRadioButton
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QFont, QColor
from PyQt6.QtWidgets import QHeaderView
import random
from PyQt6.QtCore import Qt, QEvent, QThread, pyqtSignal
import pandas as pd

import pandasai
from pandasai import SmartDataframe, Agent
from pandasai.llm.google_palm import GooglePalm
google_llm = GooglePalm(api_key="redacted")


# delete the cache folder for pandasai on startup .  This is due to a bug i encounter for some reason at work
# where the LLM calls don't work unless you delete the cache if there's already a cache folder
# the issue only seems to arise when the program is turned into an executable though
script_path = os.path.abspath(__file__)
script_directory = os.path.dirname(script_path)
cache_path = os.path.join(script_directory, "cache")

try:
    if os.path.exists(cache_path):
        shutil.rmtree(cache_path)
except:
    pass

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df = None
        self.start_agent = None

        self.setup_main_window()

    def setup_main_window(self):
        self.setStyleSheet("QMainWindow {background-color: lightgrey;}")

        # Set up the central widget
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
         # Set up the layout
        layout = QVBoxLayout(central_widget)

        myfont = QFont()
        myfont.setPointSize(10)

        self.dataframe_path_label = QLabel("<b>Enter Path to Excel or Folder with Excel files</b>")
        self.dataframe_path_label.setFont(myfont)
        self.dataframe_path_label.setAlignment(Qt.AlignmentFlag.AlignHCenter)

        self.dataframe_path = QLineEdit()
        self.dataframe_path.setPlaceholderText("Enter Path to File")
        self.dataframe_path.textChanged.connect(self.dataframe_path_changed)

        self.table = TableWithCopy()
        self.table.horizontalHeader().setSectionsMovable(True)
        self.table.setSortingEnabled(True)

        # Apply the sunken style to the QTableWidget frame
        self.table.setStyleSheet("QTableWidget { border: 2px solid gray; border-radius: 4px;  }")

        self.AI_response_label = QLabel("<b>AI Response</b>")
        self.AI_response_label.setAlignment(Qt.AlignmentFlag.AlignHCenter)


        self.horizontal_layout_ai = QHBoxLayout()

        self.AI_text_edit = QTextEdit()
        self.AI_text_edit.setMaximumHeight(150)

        # Apply the sunken style to the QTableWidget frame
        self.AI_text_edit.setStyleSheet("QTextEdit { border: 2px solid gray; border-radius: 4px;  }")

        self.horizontal_layout_ai.addWidget(self.AI_text_edit)

        # =========================================================
        self.options_group = QGroupBox()
        self.options_group.setTitle("Help AI With Output Type")

        self.group_vertical_layout = QVBoxLayout()

        self.none_option = QRadioButton("None")
        self.none_option.setChecked(True)
        self.dataframe_option = QRadioButton("Table")
        self.chart_option = QRadioButton("Chart")
        self.text_option = QRadioButton("Chat")

        self.group_vertical_layout.addWidget(self.none_option)
        self.group_vertical_layout.addWidget(self.dataframe_option)
        self.group_vertical_layout.addWidget(self.chart_option)
        self.group_vertical_layout.addWidget(self.text_option)

        self.options_group.setLayout(self.group_vertical_layout)

        # ========================================================
        self.horizontal_layout_ai.addWidget(self.options_group)
        # =========================================================
        self.horizontal_layout = QHBoxLayout()

        self.question_edit = QLineEdit()
        self.question_edit.setPlaceholderText("Ask a question")

        self.button = QPushButton("Send")
        self.button.clicked.connect(self.activate_ai)

        self.reset_button = QPushButton("Reset AI")
        self.reset_button.clicked.connect(self.dataframe_path_changed)

        self.horizontal_layout.addWidget(self.question_edit)
        self.horizontal_layout.addWidget(self.button, stretch=0)
        self.horizontal_layout.addWidget(self.reset_button, stretch=0)
        # ========================================================

        # Add the table view to the layout
        layout.addWidget(self.dataframe_path_label)
        layout.addWidget(self.dataframe_path)
        layout.addWidget(self.table)
        layout.addWidget(self.AI_response_label)
        layout.addLayout(self.horizontal_layout_ai)
        layout.addLayout(self.horizontal_layout)

        central_widget.setLayout(layout)

        # Set window properties
        self.setWindowTitle("BinaryAI")
        self.setGeometry(100, 100, 600, 400)

    def get_selected_radio_button(self, group_box):
        for child in group_box.findChildren(QRadioButton):
            if child.isChecked():
                return child
        return None

    def dataframe_path_changed(self):
        self.df = None
        self.start_agent = None
        self.AI_text_edit.clear()

    def activate_ai(self):
        question_user = self.question_edit.text()
        selected_radio_button = self.get_selected_radio_button(self.options_group)

        self.AI_text_edit.clear()
        self.AI_text_edit.setText("Thinking....")

        self.start_ai = AI_thinking(self.dataframe_path.text(), question_user, selected_radio_button, self.df, self.start_agent)
        self.start_ai.onairesult.connect(self.set_AI_response)
        self.start_ai.ontablepopulate.connect(self.populate_table)
        self.start_ai.onagent.connect(self.onagentChange)
        self.start_ai.ondatachanged.connect(self.ondataframeChange)
        self.start_ai.start()

    def onagentChange(self, agent):
        if agent is None:
            self.set_AI_response("Path/File not correct! or possibly some other error that I can't determine.")
        else:
            self.start_agent = agent

    def ondataframeChange(self, dataframe):
        self.df = dataframe

    def msg_error(self, msg1, title):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Warning)
        msg.setText(msg1)
        msg.setWindowTitle(title)
        msg.exec()

    def set_AI_response(self, value):
        self.AI_text_edit.setText(value)

    # populate table with dataframe
    def populate_table(self, result):
        self.set_AI_response("Done.")

        rows, columns = result.shape

        self.table.setRowCount(rows)
        self.table.setColumnCount(columns)

        myfont = QFont()
        myfont.setBold(True)

        self.table.horizontalHeader().setFont(myfont)
        self.table.setHorizontalHeaderLabels(result.columns.to_list())

        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        # Populate table
        for row in range(rows):
            for col in range(columns):
                item = QTableWidgetItem(str(result.iat[row, col]))
                self.table.setItem(row, col, item)

        self.table.setAlternatingRowColors(True)
        self.table.resizeRowsToContents()

    # delete the cache folder for pandasai on close.  This is due to a bug i encounter for some reason at work
    # where the LLM calls don't work unless you delete the cache if there's already a cache folder
    # the issue only seems to arise when the program is turned into an executable though
    def closeEvent(self, event):
        try:
            if os.path.exists(cache_path):
                shutil.rmtree(cache_path)
        except:
            pass

class AI_thinking(QThread):
    onairesult = pyqtSignal(str)
    ontablepopulate = pyqtSignal(object)
    onagent = pyqtSignal(object)
    ondatachanged = pyqtSignal(object)

    def __init__(self, path, question_user, selected_radio_button,  df, agent):
        super().__init__()
        self.dataframe_path = path
        self.question_user = question_user
        self.df = df
        self.start_agent = agent
        self.selected_radio_button = selected_radio_button

    def run(self):
        # get user selected output type
        output = self.chat_output_type(self.selected_radio_button.text())

        # read file to dataframe and activate agent
        if self.df is None:
            self.start_agent = self.activate_ai_agent()

        if self.start_agent is not None:
            if output is not None:
                try:
                    result = self.start_agent.chat(self.question_user, output_type=output)
                except:
                    result = "I couldn't figure it out"
            else:
                try:
                    result = self.start_agent.chat(self.question_user)
                except:
                    result = "I couldn't figure it out"
            # ==============================================

            test_result = str(result).split()
            if "Unfortunately" in test_result[0]:
                try:
                    result = self.start_agent.clarification_questions(self.question_user)

                    ai_questions = []
                    if len(result) != 0:
                        for question in result:
                            try:
                                ai_questions.append(question['question'])
                            except:
                                ai_questions.append(question)
                        result = '\n'.join(ai_questions)

                except:
                    result = "I couldn't figure it out"

            # =================================================================
            if result is not None:
                try:
                    if len(result) == 0:
                        result = "Sorry I was unable to understand Query, try rewording"
                except:
                    pass
            # =========================================================================

            if isinstance(result, pandasai.smart_dataframe.SmartDataframe):
                self.ontablepopulate.emit(result)
                self.onairesult.emit("Done.")
            else:
                self.onairesult.emit(str(result))

            # ===============================================================================
        self.onagent.emit(self.start_agent)

        if self.df is not None:
            self.ondatachanged.emit(self.df)


    # output types for argument in .chat for agent
    def chat_output_type(self, value):
        output = None
        if value == "None":
            output = None
        elif value == "Table":
            output = "dataframe"
        elif value == "Chart":
            output = "plot"
        elif value == "Text":
            output = "string"
        return output

    def read_to_dataframe(self, path):
        if os.path.isfile(path):
            file_ext = path.split(".")[-1]
            extensions = ['xls', 'xlsx', 'xlsm']
            if file_ext in extensions:
                self.df = [pd.read_excel(path)]

        elif os.path.isdir(path):
            readfolder = os.listdir(path)
            for i in readfolder:
                file_ext = i.split(".")[-1]
                extensions = ['xls', 'xlsx', 'xlsm']
                if file_ext in extensions:
                    if type(self.df) == list:
                        self.df.append(pd.read_excel(os.path.join(path, i)))
                    else:
                        self.df = [pd.read_excel(os.path.join(path, i))]


    def path_exists(self):
        path = self.dataframe_path.replace('"', "")

        if os.path.exists(path):
            return path
        else:
            return None


    def activate_ai_agent(self):
        agent = None

        path = self.path_exists()
        if path is not None:
            self.read_to_dataframe(path)
            if self.df is not None:
                agent = Agent(self.df, config={"llm": google_llm})

        return agent


class TableWithCopy(QTableWidget):
    
   # this class extends QTableWidget
   # * supports copying multiple cell's text onto the clipboard
   # * formatted specifically to work with multiple-cell paste into programs
   #   like google sheets, excel, or numbers
   #   and also copying / pasting into cells by more than 1 at a time
   

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def keyPressEvent(self, event):
        super().keyPressEvent(event)
        if event.key() == Qt.Key.Key_C and (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
            copied_cells = sorted(self.selectedIndexes())

            copy_text = ''
            max_column = copied_cells[-1].column()
            for c in copied_cells:
                copy_text += self.item(c.row(), c.column()).text()
                if c.column() == max_column:
                    copy_text += '\n'
                else:
                    copy_text += '\t'
            QApplication.clipboard().setText(copy_text)

        if event.key() == Qt.Key.Key_C and (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
            self.copied_cells = sorted(self.selectedIndexes())
        elif event.key() == Qt.Key.Key_V and (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
            r = self.currentRow() - self.copied_cells[0].row()
            c = self.currentColumn() - self.copied_cells[0].column()
            for cell in self.copied_cells:
                self.setItem(cell.row() + r, cell.column() + c, QTableWidgetItem(cell.data()))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec())
