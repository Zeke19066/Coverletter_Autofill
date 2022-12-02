# https://www.linkedin.com/in/eze-abadie
# https://github.com/Zeke19066

#System
import json
import io
import os
import sys
from datetime import date

#GUI
from PyQt5 import QtWidgets as qtw
from PyQt5 import QtGui as qtg
from PyQt5 import QtCore as qtc

#API
from apiclient import discovery
from httplib2 import Http
from oauth2client import client, file, tools
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

#NOTE NOTE NOTE NOTE
#The google API class should really be moved to its own file.

print("Starting Program: Cover Letter Generator")
"""
NOTE: For this to work properly, the keywords must be highlighted in WHITE,
    while the rest of the document is highlighted in "NONE".

A PyQt5 GUI Class handles the form processing, and then
calls the API Class functions when the form is complete.

Google API Inputs are:
    -Starting Text
    -Ending Text
    Indexing will be determined by the length of the starting Text,
    so accuracy is paramount.

API Process
1) Create a copy from the template; working file.
2) Scan for matches to codewords, note index from high to low.
3) Work backwards from end to beggining, word by word.
    delete old text
    insert new text
4) Download a .docx version of the working file.
5) Delete the working file.

Set doc ID, as found at `https://docs.google.com/document/d/YOUR_DOC_ID/edit`
This is the ID of the template file.

To hide the windows terminal for the python file:
https://stackoverflow.com/questions/1689015/run-python-script-without-windows-console-appearing
if your Python files end with .pyw instead of .py, the
standard Windows installer will set up associations
correctly and run your Python in pythonw.exe.
"""

class Window(qtw.QDialog):

    def __init__(self):
        super(Window, self).__init__()

        self.targets = {"##JOB_BORD##":"LinkedIn", #Binary; LinkedIn/ZipRecruiter checkbox
                "##COMPANY##":"",
                "##HIRING_MANAGER##":"Hiring Manager",
                "##ROLE##":"Software Engineer",
                "##CUSTOM_PARAGRAPH##\n":"",
                "##DATE##\n":""
                }
        p1 = "From Machine Learning Models to Web Applications, "
        p2 = "I am passionate about developing solutions that bring value to the client."
        self.default_paragraph = p1 + p2
        self.targets["##CUSTOM_PARAGRAPH##\n"] = self.default_paragraph
        self.targets_template = self.targets.copy()

        self.complete_path = "" #path to downloaded doc
        self.open_bool = False
        #self.threadpool = qtc.QThreadPool()
        self.threadpool = qtc.QThreadPool().globalInstance()
        document_ID = '1ldhIwp5d7RAIPv0o79YKIwGvJFVHYMoUgA3UbJJkuKY'
        self.docs = Google_API(document_ID)

        self.setWindowTitle("Cover Letter Generator")
        self.setGeometry(1000, 500, 1800, 1100)

        #QGridLayout()  QVBoxLayout()
        self.mainLayout = qtw.QVBoxLayout()

        self.windowStack = qtw.QStackedWidget(self)
        self.window1 = qtw.QWidget()
        self.window2 = qtw.QWidget()
        self.window3 = qtw.QWidget()

        #generate all windows, toggle later.
        self.window_1_GUI()
        self.window_2_GUI()
        self.window_3_GUI()
        
        self.windowStack.addWidget(self.window1)
        self.windowStack.addWidget(self.window2)
        self.windowStack.addWidget(self.window3)
        self.mainLayout.addWidget(self.windowStack)

        self.setLayout(self.mainLayout)

    def window_1_GUI(self):
        win_1_layout = qtw.QVBoxLayout()
        self.win1_font = qtg.QFont('Arial', 12)
        
        #Radio Button
        self.radiobutton_1 = qtw.QRadioButton("LinkedIn")
        self.radiobutton_1.setChecked(True)
        self.radiobutton_1.jobboard = "LinkedIn"
        self.radiobutton_1.toggled.connect(self.radioClicked)
        self.radiobutton_2 = qtw.QRadioButton("ZipRecruiter")
        self.radiobutton_2.jobboard = "ZipRecruiter"
        self.radiobutton_2.toggled.connect(self.radioClicked)

		# creating a group box
        self.formGroupBox = qtw.QGroupBox("Cover Letter Generator")
        self.formGroupBox.setFont(self.win1_font)
        self.companyLineEdit = qtw.QLineEdit() ##COMPANY##
        self.companyLineEdit.setFont(self.win1_font)
        self.managerLineEdit = qtw.QLineEdit() ##HIRING_MANAGER##
        label_1 = self.targets["##HIRING_MANAGER##"]
        self.managerLineEdit.setPlaceholderText(label_1)
        self.managerLineEdit.setFont(self.win1_font)
        self.roleLineEdit = qtw.QLineEdit() ##ROLE##
        label_2 = self.targets["##ROLE##"]
        self.roleLineEdit.setPlaceholderText(label_2)
        self.roleLineEdit.setFont(self.win1_font)
        self.customParagraph = qtw.QPlainTextEdit(self) ##CUSTOM_PARAGRAPH##\n
        self.customParagraph.setFont(self.win1_font)
        self.label_3 = self.targets["##CUSTOM_PARAGRAPH##\n"]
        self.customParagraph.setPlainText(self.label_3)
        #self.customParagraph.setPlaceholderText(label_3 + " (Limit 675 Characters)")
        self.customParagraph.textChanged.connect(self.characterLimit)#675 character limit.
        self.checkbox = qtw.QCheckBox("Open Word Document?",self)
        self.checkbox.setFont(self.win1_font)
        self.checkbox.stateChanged.connect(self.checkBox)
        self.checkbox.setChecked(True)

        #self.createForm() # calling the method that create the form
        # creating a form layout
        layout = qtw.QFormLayout()
        layout.addRow(qtw.QLabel("Job Board:"))
        layout.addRow(self.radiobutton_1, self.radiobutton_2)
        layout.addRow(qtw.QLabel("Company:"), self.companyLineEdit)
        layout.addRow(qtw.QLabel("Hiring Manager:"), self.managerLineEdit)
        layout.addRow(qtw.QLabel("Role:"), self.roleLineEdit)
        #self.customParagraph.resize(280,40)
        layout.addRow(qtw.QLabel("Custom Paragraph:"))
        layout.addRow(self.customParagraph)
        layout.addRow(self.checkbox)

		# setting layout
        self.formGroupBox.setLayout(layout)

        # creating a dialog button for ok and cancel
        self.submit_button = qtw.QPushButton("Submit")
        self.submit_button.setFont(self.win1_font)
        self.quit_button_1 = qtw.QPushButton("Quit")
        self.quit_button_1.setFont(self.win1_font)
        self.submit_button.clicked.connect(self.processForm)
        self.quit_button_1.clicked.connect(self.reject)
        # adding button box to the layout

		# adding form group box to the layout
        win_1_layout.addWidget(self.formGroupBox)

		# adding buttons to the layout
        win_1_layout.addWidget(self.submit_button)
        win_1_layout.addWidget(self.quit_button_1)

        self.window1.setLayout(win_1_layout)

    def window_2_GUI(self):
        win_2_layout = qtw.QGridLayout()

        win2_label = qtw.QLabel("Processing")
        win2_font = qtg.QFont('Arial', 60)
        win2_label.setFont(win2_font)
        win2_label.setAlignment(qtc.Qt.AlignCenter)

        self.win2_sub_label = qtw.QLabel("Initializing")
        win2_sub_font = qtg.QFont('Arial', 30)
        self.win2_sub_label.setFont(win2_sub_font)
        self.win2_sub_label.setAlignment(qtc.Qt.AlignCenter)

        #win2_label.setStyleSheet("border: 1px solid black;")
        #win2_label.resize(200, 200)

        win_2_layout.addWidget(win2_label)
        win_2_layout.addWidget(self.win2_sub_label)
        self.window2.setLayout(win_2_layout)

    def window_3_GUI(self):
        win_3_layout = qtw.QVBoxLayout()

        # creating a dialog button for ok and cancel
        self.delete_button = qtw.QPushButton("Delete File")
        self.delete_button.setFont(self.win1_font)
        self.back_button = qtw.QPushButton("Back")
        self.back_button.setFont(self.win1_font)
        self.quit_button = qtw.QPushButton("Quit")
        self.quit_button.setFont(self.win1_font)
        self.delete_button.clicked.connect(self.delete_download)
        self.quit_button.clicked.connect(self.reject)
        self.back_button.clicked.connect(self.back_button_reset)
        # adding button box to the layout

        win3_label = qtw.QLabel("Task Complete!")
        font = qtg.QFont('Arial', 40)
        win3_label.setFont(font)
        win3_label.setAlignment(qtc.Qt.AlignCenter)
        #win3_label.setStyleSheet("border: 1px solid black;")
        #win3_label.resize(200, 200)

        win_3_layout.addWidget(win3_label)

        win_3_layout.addWidget(self.delete_button)
        win_3_layout.addWidget(self.back_button)
        win_3_layout.addWidget(self.quit_button)
        self.window3.setLayout(win_3_layout)

    # Method called when form is accepted
    def processForm(self):

        #Processing Screen
        self.displayToggle(1)
        # Pass the function to execute
        worker = Worker(self.api_task) # Any other args, kwargs are passed to the run function
        worker.signals.result.connect(self.status_bar)
        # Execute
        self.threadpool.start(worker)
        self.status_bar("text",mode="Reset")

    def api_task(self):

        self.targets["##COMPANY##"] = self.companyLineEdit.text()
        self.companyLineEdit.setText("") #Reset Value
        val_1 = len(self.managerLineEdit.text())
        val_2 = len(self.roleLineEdit.text())
        val_3 = len(self.customParagraph.toPlainText())
        if val_1 > 0: #not blank
            self.targets["##HIRING_MANAGER##"] = self.managerLineEdit.text()
        if val_2 > 0: #not blank
            self.targets["##ROLE##"] = self.roleLineEdit.text()
        if val_3 > 0: #not blank
            self.targets["##CUSTOM_PARAGRAPH##\n"] = self.customParagraph.toPlainText()

        #Date
        date_obj = date.today()
        date_string = date_obj.strftime("%b %d, %Y") #%B is full month
        self.targets["##DATE##\n"] = date_string + "\n"

        #Account for the line break
        paragraph = self.targets["##CUSTOM_PARAGRAPH##\n"]
        self.targets["##CUSTOM_PARAGRAPH##\n"] = paragraph + "\n"
        p_print = self.targets["##CUSTOM_PARAGRAPH##\n"]

		# printing the form information
        print(f'Job Board: {self.targets["##JOB_BORD##"]}')
        print(f'Company: {self.targets["##COMPANY##"]}')
        print(f'Hiring Manager: {self.targets["##HIRING_MANAGER##"]}')
        print(f'Role: {self.targets["##ROLE##"]}')
        print(f'Custom Paragraph: {p_print}')
        print('--------------------------------------')
        print("Form Complete, Initializing Google API")
        
        #begin Google API calls
        #[Index, Start_Word, End_Word]
        self.docs.api_main(self.targets)
        self.complete_path = self.docs.complete_path
        if self.open_bool:
            word_doc = os.startfile(self.complete_path)
    
        self.displayToggle(2)

    def radioClicked(self):
        radioButton = self.sender()
        if radioButton.isChecked():
            print("Job Board is %s" % (radioButton.jobboard))
            self.targets["##JOB_BORD##"] = radioButton.jobboard

    def delete_download(self):
        os.remove(self.complete_path)
        print("Download Deleted")
        self.delete_button.setDisabled(True)
    
    def displayToggle(self, i=0):
        self.windowStack.setCurrentIndex(i)

    def checkBox(self, state):

        if state == qtc.Qt.Checked:
            self.open_bool = True
            print('Document Open Enabled')
        else:
            self.open_bool = False
            print('Document Open Disabled')

    def characterLimit(self, limit=675):
        temp_text = self.customParagraph.toPlainText()
        if len(temp_text)>limit:
            temp_text = temp_text[:limit]
            #self.customParagraph = qtw.QPlainTextEdit(self) ##CUSTOM_PARAGRAPH##\n
            self.customParagraph.clear()
            self.customParagraph.insertPlainText(temp_text)

    def back_button_reset(self):
        self.delete_button.setDisabled(False)
        self.targets = self.targets_template.copy()
        self.customParagraph.clear()
        self.customParagraph.insertPlainText(self.label_3)
        self.roleLineEdit.clear()
        self.roleLineEdit.setText("")
        self.managerLineEdit.clear()
        self.managerLineEdit.setText("")
        self.displayToggle(0)

    def status_bar(self, text, mode="Read"):
        #print("recieved singnal")

        if mode == "Read":
            r = open('print_log.txt','r')
            text_1 = r.readlines()
            text_1 = text_1[-1]
            r.close()
            self.win2_sub_label.setText(text_1)

        elif mode == "Reset":
            w = open('print_log.txt','w')
            w.write("")
            w.close()

class Worker(qtc.QRunnable):
    '''
    Worker thread

    Inherits from QRunnable to handler worker thread setup, signals and wrap-up.
    entire functions can be submitted, including class variables.

    :param callback: The function callback to run on this worker thread. Supplied args and
                     kwargs will be passed through to the runner.
    :type callback: function
    :param args: Arguments to pass to the callback function
    :param kwargs: Keywords to pass to the callback function

    '''

    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()
        # Store constructor arguments (re-used for processing)
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()
        self.setAutoDelete(True) #sets worker to autodelete.

    #@pyqtSlot()
    def run(self):
        '''
        Initialise the runner function with passed args, kwargs.
        '''
        #Print Loggers
        def MyHookOut(text):
            a = open('print_log.txt','a')
            a.write(text)
            a.close()

            self.signals.result.emit("DING")
            return 1,1,'Out Hooked:'+ text

        print('Hook Start')
        phOut = PrintHook()
        phOut.Start(MyHookOut)

        self.fn(*self.args, **self.kwargs)

        phOut.Stop()
        print('STDOUT Hook end')

class WorkerSignals(qtc.QObject):
    '''
    Defines the signals available from a running worker thread.
    Supported signals are:
    finished: No data
    error: tuple (exctype, value, traceback.format_exc() )
    result: object data returned from processing, anything
    '''
    finished = qtc.pyqtSignal()
    error = qtc.pyqtSignal(tuple)
    #result = qtc.pyqtSignal(object)
    result = qtc.pyqtSignal(str)

class Google_API:
    # Class for interfacing with Google Docs & Drive API
    def __init__(self, doc_id, open_bool=True):
        #https://docs.google.com/document/d/doc_id/edit
        #open bool auto-opens downloaded word doc.
        self.requests = []
        self.document_copy_id = 0
        self.source_id = doc_id
        self.open_bool = open_bool
        self.complete_path = "" #path to downloaded doc
        self.docs_authorize()
        self.drive_authorize()

    #Token Authorization for Google Docs
    def docs_authorize(self):
        # Initialize credentials and instantiate Docs API service
        store = file.Storage('token_docs.json')
        creds = store.get()
        # Safe Mode: 'https://www.googleapis.com/auth/documents.readonly'
        scope_docs = 'https://www.googleapis.com/auth/documents'
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets('credentials.json', scope_docs)
            creds = tools.run_flow(flow, store)

        discovery_doc = ('https://docs.googleapis.com/$discovery/rest?'
                 'version=v1')
        self.docs_service = discovery.build('docs', 'v1', http=creds.authorize(
            Http()), discoveryServiceUrl=discovery_doc)

    #Token Authorization for Google Drive
    def drive_authorize(self):
        # Initialize credentials and instantiate Docs API service
        store = file.Storage('token_drive.json')
        creds = store.get()
        scope_drive = 'https://www.googleapis.com/auth/drive'
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets('credentials.json', scope_drive)
            creds = tools.run_flow(flow, store)
        self.drive_service = build('drive', 'v3', credentials=creds)

    #main function
    def api_main(self, targets_list):
        #targets_list = [[starting index, starting word, ending word],
        #                [starting index2, starting word2, ending word2],...]
        print("Initializing Coverletter Generator")
        self.copy_document(self.source_id)
        
        file_name = ""
        print("Generating Task List")
        task_list = self.text_scan(self.document_copy_id, targets_list)

        print("Processing Task List")
        for entry in task_list:
            i_start = entry[0]
            w_start = entry[1]
            w_end = entry[2]
            self.delete(i_start,w_start)
            self.insert(i_start,w_end)
            if w_start == "##COMPANY##":
                file_name = w_end
        #Send the requests
        self.submit_edit_requests(self.document_copy_id)
        self.download_document(self.document_copy_id, file_name)
        self.delete_document(self.document_copy_id)
        self.requests = []
        print("Task Complete")

    def insert(self, index, text):
        if len(text) > 0:
            subrequest = [
                    {
                    'insertText': {
                        'location': {
                            'index': 0,
                        },
                        'text': "###ERROR###"
                    }
                }
            ]
            subrequest[0]['insertText']['location']['index'] = index
            subrequest[0]['insertText']['text'] = text
            self.requests.append(subrequest[0])

    def delete(self, index, text):
        end = index + len(text)
        subrequest = [
            {
                'deleteContentRange': {
                    'range': {
                        'startIndex': 0,
                        'endIndex': 0,
                    }
                }
            },
        ]
        subrequest[0]['deleteContentRange']['range']['startIndex'] = index
        subrequest[0]['deleteContentRange']['range']['endIndex'] = end
        self.requests.append(subrequest[0])

    def submit_edit_requests(self, doc_id):
        result = self.docs_service.documents().batchUpdate(
            documentId=doc_id, body={'requests': self.requests}).execute()
        print("Edits Submitted")

    def copy_document(self, doc_id):
        print("Generating Template Copy", end="")
        #https://docs.google.com/document/d/doc_id/edit
        copy_title = 'Cover_Instance'
        body = {
            'name': copy_title
        }
        drive_response = self.drive_service.files().copy(
            fileId=doc_id, body=body).execute()
        self.document_copy_id = drive_response.get('id')
        print("Complete")

    def delete_document(self, doc_id):
        #https://docs.google.com/document/d/doc_id/edit

        drive_response = self.drive_service.files().delete(fileId=doc_id).execute()
        #When successful, drive_response should be blank.
        print(f"Delete Completed {drive_response}")

    def download_document(self, doc_id, file_name, destktop_bool=True):
        #https://developers.google.com/drive/api/guides/ref-export-formats  <- Has the formats
        #application/vnd.openxmlformats-officedocument.wordprocessingml.document
        """Download a Document file in PDF format.
        Args:
            real_file_id : file ID of any workspace document format file
        Returns : IO object with location

        Load pre-authorized user credentials from the environment.
        TODO(developer) - See https://developers.google.com/identity
        for guides on implementing OAuth2 for the application.
        """
        # word.docx format
        file_format = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        try:
            # pylint: disable=maybe-no-member
            request = self.drive_service.files().export_media(fileId=doc_id,
                                                mimeType=file_format)
            file = io.BytesIO()
            downloader = MediaIoBaseDownload(file, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
                print(F'Download {int(status.progress() * 100)}.')
            print("Download Complete")

        except HttpError as error:
            print(F'An error occurred: {error}')
            file = None

        #for Desktop
        if destktop_bool: #desktop storage
            subfile =  file.getvalue()
            d_path = r"C:\Users\Ezeab\Desktop"
            d_path = d_path + "//" + file_name + ' Cover Letter.docx'
            open(d_path, 'wb').write(subfile)
            self.complete_path = d_path

        elif not destktop_bool: #local storage
            subfile =  file.getvalue()
            path = "Generated_Letters//" + file_name + ' Cover Letter.docx'
            open(path, 'wb').write(subfile)
            cwd = os.getcwd()
            self.complete_path = cwd + "//" + path

    def text_scan(self, doc_id, targets):
        # PARSING STRUCTURE:
        # "body">"content"
        # for entry in "content":
        #    entry>"paragraph">"elements" #not every "entry" contains a "paragraph" subitem...
        #    for element in "elements":
        #       element>"textRun">"content"
        #           - HERE COMPARE THE CODEWORD
        #           if match:
        #               element>"startIndex"
        #                   - HERE IS THE START INDEX

        #Output is nested list [[index1,old_w1,new_w1],
        #                       [index2,old_w2,new_w2],...]

        def scrubber(result_dict, targets):
            print("Begin Scrubber")
            task_list = []

            for entry in result_dict["body"]["content"]:
                for key, val in entry.items():
                    if key == "paragraph": #it contains a paragraph item; not all do.
                        for element in entry["paragraph"]["elements"]:
                            subject = element["textRun"]["content"]
                            for key, val in targets.items():
                                #We have a match
                                if subject == key:
                                    index = element["startIndex"]
                                    subtask = [index,key,val]
                                    task_list.append(subtask)
            
            #set order high to low
            task_list.reverse()
            print(f"Scrubber Complete")
            return task_list

        # First we pull the unedited version
        # Do a document "get" request and print the results as formatted JSON
        result = self.docs_service.documents().get(documentId=doc_id).execute()
        
        ##Save JSON
        result_json = json.dumps(result)
        path = r'doc_sample.json'

        with open(path, 'w') as f:
            f.write(result_json)
            f.close()
        
        #print(json.dumps(result, indent=4, sort_keys=True))
        result_dict = result
        # Scrub for keyword matches
        task_list = scrubber(result_dict, targets)
        return task_list

class PrintHook:
    #This class intercepts python print statements 
    #out = 1 means stdout will be hooked
    #out = 0 means stderr will be hooked
    def __init__(self,out=1):
        self.func = None##self.func is userdefined function
        self.origOut = None
        self.out = out

    #user defined hook must return three variables
    #proceed,lineNoMode,newText
    def TestHook(self,text):
        f = open('print_log.txt','a')
        f.write(text)
        f.close()
        return 0,0,text

    def Start(self,func=None):
        if self.out:
            sys.stdout = self
            self.origOut = sys.__stdout__
        else:
            sys.stderr= self
            self.origOut = sys.__stderr__
            
        if func:
            self.func = func
        else:
            self.func = self.TestHook

    #Stop will stop routing of print statements thru this class
    def Stop(self):
        self.origOut.flush()
        if self.out:
            sys.stdout = sys.__stdout__
        else:
            sys.stderr = sys.__stderr__
        self.func = None

    #override write of stdout        
    def write(self,text):
        proceed = 1
        lineNo = 0
        addText = ''
        if self.func != None:
            proceed,lineNo,newText = self.func(text)
        if proceed:
            if text.split() == []:
                self.origOut.write(text)
            else:
                #if goint to stdout then only add line no file etc
                #for stderr it is already there
                if self.out:
                    if lineNo:
                        try:
                            raise "Dummy"
                        except:
                            newText =  'line('+str(sys.exc_info()[2].tb_frame.f_back.f_lineno)+'):'+newText
                            codeObject = sys.exc_info()[2].tb_frame.f_back.f_code
                            fileName = codeObject.co_filename
                            funcName = codeObject.co_name
                    #self.origOut.write('file '+fileName+','+'func '+funcName+':')
                #self.origOut.write(newText)
                self.origOut.write("hooked:"+text)
    
    #pass all other methods to __stdout__ so that we don't have to override them
    def __getattr__(self, name):
        return self.origOut.__getattr__(name)

def API_Troubleshooting_Main():
    #[Index, Start_Word, End_Word]
    targets = {"##JOB_BORD##":"Time Chat",
               "##COMPANY##":"Abraham Lincoln",
               "##ROLE##":"Major General", 
               "##CUSTOM_PARAGRAPH##\n":"Good Job on the civil war. I'd skip that play later...\n"
                }
    document_ID = '1ldhIwp5d7RAIPv0o79YKIwGvJFVHYMoUgA3UbJJkuKY'
    docs = Google_API(document_ID)
    docs.api_main(targets)

if __name__ == "__main__":
    # create pyqt5 app
    app = qtw.QApplication(sys.argv)

	# create the instance of our Window
    window = Window()

	# showing the window
    window.show()

    input("press close to exit")
    sys.exit(app.exec())

