import sys
import os
import PyPDF2

from PyQt5.QtWidgets import (
    QApplication, QDialog, QWidget, QVBoxLayout, QTableWidget,
    QTableWidgetItem, QPushButton, QTabWidget, QHBoxLayout, 
    QMessageBox, QFileDialog, QLineEdit, QLabel
)

from PyQt5.QtCore import Qt
import xlwings as wx
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import copy

class PDFProcessor(QWidget):
    def __init__(self):
        super().__init__()
        self.titles = []
        self.drivers = []
        self.filtered_titles = []
        self.filtered_drivers = []
        self.greyhounds = []
        self.rome_num = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"]
        self.times = []
        self.opinion = []

        self.initUi()

    def initUi(self):
        self.setWindowTitle('Data Editor')

        layout = QVBoxLayout()

        # Load PDF button
        load_button = QPushButton("Load PDF")
        load_button.clicked.connect(self.load_pdf)
        layout.addWidget(load_button)

        # Save Data and Make PPT buttons in a horizontal layout
        button_layout = QHBoxLayout()
        
        save_button = QPushButton("Save Data to CSV")
        save_button.clicked.connect(self.save_data_to_csv)
        button_layout.addWidget(save_button)

        ppt_button = QPushButton("Make PPT")
        ppt_button.clicked.connect(self.make_ppt)
        button_layout.addWidget(ppt_button)
        
        layout.addLayout(button_layout)


        self.tabs = QTabWidget()
        self.layout_titles_tab()
        self.layout_drivers_tab()

        layout.addWidget(self.tabs)
        self.setLayout(layout)

    def showErrorAlert(self, title, message):
        # Create a QMessageBox instance for error alert
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)       # Set icon to "Critical" for errors
        msg.setWindowTitle(title)               # Title of the alert box
        msg.setText(message)                    # Main error message
        msg.setStandardButtons(QMessageBox.Ok)  # Add an "OK" button
        msg.exec_()                             # Show the alert box
    
    def layout_titles_tab(self):
        titles_tab = QWidget()

        # Add search bar layout for Titles
        titles_search_layout = QHBoxLayout()
        titles_search_layout.addStretch(1)
        search_label = QLabel("Search Titles:")
        self.titles_search_input = QLineEdit()
        self.titles_search_input.setPlaceholderText("Search titles...")
        search_button = QPushButton("Search")
        search_button.clicked.connect(self.filter_titles_table)
        titles_search_layout.addWidget(search_label)
        titles_search_layout.addWidget(self.titles_search_input)
        titles_search_layout.addWidget(search_button)

        # Table for Titles
        self.titles_table = QTableWidget()
        self.titles_table.setColumnCount(6)
        self.titles_table.setHorizontalHeaderLabels(["Id", "Title", "Distance", "Time", "Start type", "Opinion"])
        self.load_titles_to_table()

        add_button = QPushButton("Add Row")
        add_button.clicked.connect(self.add_row_titles)

        remove_button = QPushButton("Remove Row")
        remove_button.clicked.connect(self.remove_row_titles)

        titles_layout = QVBoxLayout()
        titles_layout.addLayout(titles_search_layout)
        titles_layout.addWidget(self.titles_table)
        titles_layout.addWidget(add_button)
        titles_layout.addWidget(remove_button)

        titles_tab.setLayout(titles_layout)
        self.tabs.addTab(titles_tab, "Titles")
    
    def load_titles_to_table(self):
        if(self.titles_search_input.text() == ''):
            self.filtered_titles = self.titles

        self.titles_table.setRowCount(len(self.filtered_titles))
            
        for row, title in enumerate(self.filtered_titles):
            self.titles_table.setItem(row, 0, QTableWidgetItem(str(title["id"])))
            self.titles_table.setItem(row, 1, QTableWidgetItem(title["title"]))
            self.titles_table.setItem(row, 2, QTableWidgetItem(str(title["distance"])))
            self.titles_table.setItem(row, 3, QTableWidgetItem(title["time"]))
            self.titles_table.setItem(row, 4, QTableWidgetItem(title["stype"]))
            self.titles_table.setItem(row, 5, QTableWidgetItem(title["opinion"]))

    def update_title_data(self):
        # Update the titles list based on the current table contents
        for row in range(self.titles_table.rowCount()):
            # Access the current title entry in self.titles
            title_data = self.titles[row]
            
            # Update fields with data from the table
            title_data['id'] = self.titles_table.item(row, 0).text() if self.titles_table.item(row, 0) else ''
            title_data['title'] = self.titles_table.item(row, 1).text() if self.titles_table.item(row, 1) else ''
            title_data['distance'] = int(self.titles_table.item(row, 2).text()) if self.titles_table.item(row, 2) else 0
            title_data['time'] = self.titles_table.item(row, 3).text() if self.titles_table.item(row, 3) else ''
            title_data['stype'] = self.titles_table.item(row, 4).text() if self.titles_table.item(row, 4) else ''
            title_data['opinion'] = self.titles_table.item(row, 5).text() if self.titles_table.item(row, 5) else ''

    def filter_titles_table(self):
        search_text = self.titles_search_input.text().lower()
        self.filtered_titles = [title for title in self.titles if search_text in title["title"].lower()]
        self.load_titles_to_table()

    def add_row_titles(self):
        self.titles.append({"id": len(self.titles), "title": "", "distance": 0, "time": "", "stype": "", "opinion": ""})
        self.load_titles_to_table()

    def remove_row_titles(self):
        row = self.titles_table.currentRow()
        if row >= 0:
            del self.titles[row]
            self.load_titles_to_table()
        else:
            QMessageBox.warning(self, 'Warning', 'No row selected!')


    def layout_drivers_tab(self):
        drivers_tab = QWidget()

        # Add search bar layout for drivers
        drivers_search_layout = QHBoxLayout()
        drivers_search_layout.addStretch(1)
        search_label = QLabel("Search drivers:")
        self.drivers_search_input = QLineEdit()
        self.drivers_search_input.setPlaceholderText("Search drivers...")
        search_button = QPushButton("Search")
        search_button.clicked.connect(self.filter_drivers_table)
        drivers_search_layout.addWidget(search_label)
        drivers_search_layout.addWidget(self.drivers_search_input)
        drivers_search_layout.addWidget(search_button)

        # Table for drivers
        self.drivers_table = QTableWidget()
        self.drivers_table.setColumnCount(6)
        self.drivers_table.setHorizontalHeaderLabels(["Horse Number", "Horse Name","Horse Distance", "Driver Name", "Futam Number", "is run"])
        self.load_drivers_to_table()

        add_button = QPushButton("Add Row")
        add_button.clicked.connect(self.add_row_drivers)

        remove_button = QPushButton("Remove Row")
        remove_button.clicked.connect(self.remove_row_drivers)

        drivers_layout = QVBoxLayout()
        drivers_layout.addLayout(drivers_search_layout)
        drivers_layout.addWidget(self.drivers_table)
        drivers_layout.addWidget(add_button)
        drivers_layout.addWidget(remove_button)

        drivers_tab.setLayout(drivers_layout)
        self.tabs.addTab(drivers_tab, "drivers")

    def load_drivers_to_table(self):
        if(self.drivers_search_input.text() == ''):
            self.filtered_drivers = self.drivers
       

        self.drivers_table.setRowCount(len(self.filtered_drivers))
        #for row, driver in enumerate(self.filtered_drivers): print(row,self.filtered_drivers)
        
        try:
            for row, driver in enumerate(self.filtered_drivers):
                self.drivers_table.setItem(row, 0, QTableWidgetItem(driver["Hnum"]))
                self.drivers_table.setItem(row, 1, QTableWidgetItem(driver["Hname"]))
                self.drivers_table.setItem(row, 2, QTableWidgetItem(str(driver["Hdist"])))
                self.drivers_table.setItem(row, 3, QTableWidgetItem(driver["Dname"]))
                self.drivers_table.setItem(row, 4, QTableWidgetItem(str(driver["futam"])))
                self.drivers_table.setItem(row, 5, QTableWidgetItem(str(driver["is_run"])))
        except:
            print("can't load data")

    def update_drivers_data(self):
    # Update the drivers based on the current table contents
        #for row in range(self.drivers_table.rowCount()):
            #print( self.drivers[row]['Dname'] )
            #print(self.drivers_table.item(row, 3).text())

        for row in range(self.drivers_table.rowCount()):
            self.drivers[row]['Hnum']     = self.drivers_table.item(row, 0).text() if self.drivers_table.item(row, 0) else ''
            self.drivers[row]['Hname']    = self.drivers_table.item(row, 1).text() if self.drivers_table.item(row, 1) else ''
            self.drivers[row]['Hdist']    = self.drivers_table.item(row, 2).text() if self.drivers_table.item(row, 2) else ''
            self.drivers[row]['Dname']    = self.drivers_table.item(row, 3).text() if self.drivers_table.item(row, 4) else ''
            self.drivers[row]['futam']    = self.drivers_table.item(row, 4).text() if self.drivers_table.item(row, 5) else '0'
            self.drivers[row]['is_run']   = self.drivers_table.item(row, 5).text() if self.drivers_table.item(row, 6) else '1'

    def filter_drivers_table(self):
        search_text = self.drivers_search_input.text().lower()
        self.filtered_drivers = [driver for driver in self.drivers if search_text in driver["Dname"].lower()]
        self.load_drivers_to_table()
        
    def add_row_drivers(self):
        self.drivers.append({"id": len(self.drivers), "Hnum": "", "Hname":"","Hdist":0, "Dname":"", "futam":0, "is_run":1})
        self.load_drivers_to_table()

    def remove_row_drivers(self):
        row = self.drivers_table.currentRow()
        if row >= 0:
            del self.drivers[row]
            self.load_drivers_to_table()
        else:
            QMessageBox.warning(self, 'Warning', 'No row selected!')

    def load_pdf(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open PDF File", "", "PDF Files (*.pdf)")
        if file_name:
            self.extract_data_from_pdf(file_name)






    def extract_data_from_pdf(self, file_name):
       
        reader = PyPDF2.PdfReader(file_name)
        num_pages = len(reader.pages)
        pages = []

        for page_num in range(num_pages):
            page = reader.pages[page_num]
            text = page.extract_text()
            pages.append(text.split("\n"))


        for page in pages:
            for ln in page:
                if 'pálya: 10' in ln:
                    self.times = ln.split(" ")
                    break
        
        if "(ügető)" in self.times: del self.times[0:5]
        else: del self.times[0:4]
        
        self.times = [time.replace(".", ":") for time in self.times]

        #1 Erika Kapitány /Homokvihar Erika Kapitány/ (piros)

        for page in pages:
            for ln in page:

                try:
                    line = [data for data in ln.split(" ") if data != "" and data != "AGÁR"]
                    #colorGen = ["pk","stpk","sher","pher"]
                    ct = 0
                    for i in ln: 
                        if i=="/": ct+=1

                    if int(line[0]) in range(1,7) and "/" in ln and ct==2:
                        if not "/" in line[2]: self.greyhounds.append(" ".join(line[1:3]))
                        else: self.greyhounds.append(line[1])


                except:
                    pass
        #for i in self.greyhounds: print(i)

        for page in pages:
            for ln in page:
                isIn = False

                for i in self.greyhounds: 
                    if i in ln: 
                        isIn = True
                        break
                line = (ln + '.')[:-1]

                if 'Véleményünk:' in line and not line.replace("Véleményünk: ","") in self.opinion and not isIn:
                    op = line.replace("Véleményünk: ","")
                    #print(ln)

                    op_list= op.split(" ")
                    index = 0
                    for i in range(10,15): 
                        if str(i) in op_list:
                            index = op_list.index(str(i))
                            break
                   
                    if(index != 0): 
                        #self.drivers.append(self.clean_drivers_data(op_list[index:],fn))
                        op_list = op_list[:index]


                    #print(" ".join(op_list))

                    if "jackpot" in op_list:
                        for i in range(0,len(op_list)):
                            
                            if op_list[i] == "jackpot":
                                op_list = op_list[:i-1]
                                break
                    
                    
                    if len(op_list[-1])>5:
                        last = op_list[-1]
                        #print(last)
                        for i in range(0,len(last)):
                            if ")" == last[i]: 
                                op_list[-1]=last[:i+1]
                                break
                    
                    self.opinion.append(" ".join(op_list))
        

        
        
        n = 0
        fn = 0
        hn = 0
        for page in pages:
            for ln in page:
                if "Ft" in  ln:
                    pass #print(line)
                ln = ln.replace("↑","")
                ln = ln.replace("\t","")
                ln = ln.replace("A fogadási ajánlatot keresse a www.lovifogadas.hu oldalon, illetve a Kincsem Parkban!", "")
                line = ln.strip().split(" ")
                line = [item for item in line if item != '']
                #line = [item for item in line if item != ' ']
                line = [item for item in line if item != '*']
                #line = [item for item in line if item != '↑']

                for h in range(10,20):
                    for m in range(0,55,5):
                        min=""
                        if m<10:min=f"0{m}"
                        else: min=str(m)
                        line = [item for item in line if item != f'{h}:{min}']

                if 'Ügető' in line:
                    line.remove('Ügető')
                if 'ÜGETŐ' in line:
                    line.remove('ÜGETŐ')
                if '(IRE)' in line:
                    line.remove('(IRE)')
                if '(GB)' in line:
                    line.remove('(GB)')
                if '(FR)' in line:
                    line.remove('(FR)')
                if '(USA)' in line:
                    line.remove('(USA)')
                if '(GER)' in line:
                    line.remove('(GER)')
                if '(RO)' in line:
                    line.remove('(RO)')
                if '(CZE)' in line:
                    line.remove('(CZE)')
                if '(SRB)' in line:
                    line.remove('(SRB)')
                if '(FV)' in line:
                    line.remove('(FV)')
                if '(fv)' in line:
                    line.remove('(fv)')
                if '(SVK)' in line:
                    line.remove('(SVK)')
                if '(SVN)' in line:
                    line.remove('(SVN)') 

                #3 Petra La Roque 1960 m Csordás Emil (29,1-4) (Kék, sárga váll és szegély)
                
              
                if "KVALIFIKÁCIÓ" in ln and " m " in ln:
                    self.titles.append(self.clean_titles_data(ln.split(" ")))
                if "PRÓBAFUTAM" in ln and " m " in ln:
                    self.titles.append(self.clean_titles_data(ln.split(" ")))

                if "158/VIII. CONNOLLY’S RED MILLS" in ln: print(len(line))

                if len(line)>3 and len(line)<=14:
                    if "/" in line[0] :
                        
                        
                        if line[0].split("/")[0].isdigit() and not line[0].split("/")[1].isdigit():
                            
                            line[0] = line[0].split("/")[1]
                            self.titles.append(self.clean_titles_data(line))
                            
                




                if len(line) >= 5 and str(line[0]).isdigit() and ")" in line[-1]:
                    if 'm' in line and line[line.index("m")-1].isdigit():
                        if(int(line[0]) <= hn):
                            fn+=1
                        self.drivers.append(self.clean_drivers_data(line,fn))
                        #print(self.clean_drivers_data(line,fn))
                        hn = int(line[0])

                
                #Véleményünk: Exclusive Adri (6) – Galatea (7) – Gana Ris (9) – R enebell BD (3) 10 Fantom KE 1900 m  Weber Zsolt  (20,0-3) (Narancssárga-fekete)
                isIn = False

                for i in self.greyhounds: 
                    if i in ln: 
                        isIn = True
                        break
                if 'Véleményünk: ' in ln and not isIn:
                    op =  ln.split(" ")
                    op = [data for data in op if data!=""]
                    for i in op:
                        try:
                            if int(i) in range(10,15):
                                self.drivers.append(self.clean_drivers_data(op[op.index(i):],fn))
                        except:
                            continue
                 

        self.load_titles_to_table()
        self.load_drivers_to_table()
        


        #for i in self.titles: print(i)

                


        #for i in times:
        #   print(i)
    def clean_drivers_data(self,ln,fn):
        #['8', 'Flinga', 'VT', '1980', 'm', 'Tänczer', 'Tamás', 'am.', '(21,3-2)', '(Fehér,', 'ezüst', 'váll', 'és', 'ujjak)']
        forlen = len(ln)
        for i in range(0, len(ln)):
            if "(" in ln[i] and not (")" in ln[i]):
                for d in range(i, len(ln)):
                    if ")" in ln[d]:
                        text = " ".join(ln[i:d + 1])
                        del ln[i:d + 1]
                        ln.insert(i, text)
                        break
            if forlen != len(ln):
                break

        for i in range(0,len(ln)):
            txt = ""
            if "m"==ln[i]:
                txt = " ".join(ln[1:i-1])
                del ln[1:i-1]
                ln.insert(1,txt)
                break

        for i in range(0,len(ln)):
            txt = ""
            if "m"==ln[i]:
                for d in range(0,len(ln)):
                    if '(' in ln[d]:
                        txt = " ".join(ln[i+1:d])
                        del ln[i+1:d]
                        ln.insert(i+1,txt)
                        break
                break
        

        return {"Hnum": ln[0], "Hname":ln[1],"Hdist":int(ln[2]), "Dname":ln[4], "futam":fn, "is_run":1}

    
    def clean_titles_data(self,ln):
        
        ln = [data for data in ln if data != '' ]
        #ln = [data.replace(".","") for data in ln ]


        
        forlen = len(ln)
        for i in range(0, len(ln)):
            if "(" in ln[i] and not (")" in ln[i]):
                for d in range(i, len(ln)):
                    if ")" in ln[d]:
                        text = " ".join(ln[i:d + 1])
                        del ln[i:d + 1]
                        ln.insert(i, text)
                        break
            if forlen != len(ln):
                break

        for i in range(0,len(ln)):
            txt = ""
            if "m"==ln[i]:
                txt = " ".join(ln[1:i-1])
                del ln[1:i-1]
                ln.insert(1,txt)
                break 
        
        

        futam_number = ""


        if "Q" in ln[0]:
            futam_number = ln[0]
        elif "PRÓBAFUTAM" in ln[0]:
            futam_number="P"
        else:
            ln[0] = ln[0].replace(".","")
            if ln[0] in self.rome_num:
                futam_number = str(self.rome_num.index(ln[0]))
            elif "X" in ln[0]:
                futam_number = str(self.rome_num.index(ln[0][1:]) +10)
        tm = ""
        if futam_number.isdigit():
            tm = self.times[int(futam_number)]
        opinion = "" 
        if futam_number.isdigit():
            opinion = self.opinion[int(futam_number)]
        if futam_number == "P": title = "PRÓBAFUTAM"
        else: title = ln[1]
        return { "id": futam_number,"title": title, "distance": ln[2], "time": tm, "stype": ln[4], "opinion": opinion}


    def save_data_to_csv(self):
        self.update_drivers_data()
        self.update_title_data()
        if not os.path.isdir("csv"):
            os.makedirs("csv")

        with open("csv/titles_data.csv","w",encoding="utf-8") as f:
            f.write("Id;Title;Distance;Time;Start type;Opinion\n")
            for title in self.titles:
                if title["id"].isdigit():
                    f.write(f"{title["id"]};{title["title"]};{title["distance"]};{title["time"]};{title["stype"]};{title["opinion"]}  \n")



        with open("csv/drivers_data.csv","w",encoding="utf-8") as f:
            f.write("Horse Number;Horse Name;Horse Distance;Driver Name;Futam Number;is run\n")
            #{'Hnum': '8', 'Hname': 'Flinga VT', 'Hdist': '1980', 'Dname': '9', 'futam': '1', 'is_run': '1'}
            ct=0
            for i in self.titles:
                if not i["id"].isdigit(): ct+=1
            for driver in self.drivers:
                futam = int(driver["futam"]) - ct
                f.write(f"{driver["Hnum"]};{driver["Hname"]};{driver["Hdist"]};{driver["Dname"]};{futam};{driver["is_run"]} \n")

    def slide1(self,ppt,slide_layout,futam):
        slide1 = ppt.slides.add_slide(slide_layout)
        #self.set_slide_duration(slide1, 5) 
        #slide1.slide_show.transition.duration = 50
        slide1bc = slide1.background.fill
        slide1bc.solid()
        slide1bc.fore_color.rgb = RGBColor(0, 0, 0)
        
        title_box = slide1.shapes.add_textbox(Inches(0), Inches(0), Inches(10), Inches(1.4))
        title_frame = title_box.text_frame
        
        title = title_frame.add_paragraph()
        if futam.id < len(self.rome_num): title_str = f"{self.rome_num[futam.id]}. {futam.title}".strip()
        else:                             title_str = f"X{self.rome_num[futam.id-10]}. {futam.title}".strip()
        title.text = title_str
        if   (len(title_str) < 39 ): title.font.size = Pt(45)
        elif (len(title_str) <= 58): title.font.size = Pt(40)
        else:                        title.font.size = Pt(36)
        title.font.bold = True
        title.font.color.rgb = RGBColor(255, 229, 121)
        
        title.alignment = PP_ALIGN.CENTER
        title_frame.word_wrap = True 
        title_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        text_box = slide1.shapes.add_textbox(Inches(4.93), Inches(1.5), Inches(5.07), Inches(0.7))
        text_frame = text_box.text_frame
        
        text = text_frame.add_paragraph()
        text.text = f"Pálya: 11 Kincsem Park"
        text.font.size = Pt(36)
        text.font.bold = True
        text.font.color.rgb = RGBColor(255, 229, 121)
        
        text.alignment = PP_ALIGN.RIGHT
        text_frame.word_wrap = True 
        text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
        
        data_box = slide1.shapes.add_textbox(Inches(0), Inches(2.53), Inches(3.69), Inches(1.31))
        data_frame = data_box.text_frame
        
        data_text = data_frame.add_paragraph()
        data_text.text = f"{futam.dist} m\n{futam.stype}"

        data_text.font.size = Pt(36)
        data_text.font.bold = True
        data_text.font.color.rgb = RGBColor(255, 229, 121)
        
        data_text.alignment = PP_ALIGN.LEFT
        data_frame.word_wrap = True 
        data_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
        
        time_box = slide1.shapes.add_textbox(Inches(5.52), Inches(2.53), Inches(1.51), Inches(1.31))
        time_frame = time_box.text_frame
        
        time_text = time_frame.add_paragraph()
        time_text.text = f"Start:\n{futam.time}"

        time_text.font.size = Pt(36)
        time_text.font.bold = True
        time_text.font.color.rgb = RGBColor(255, 255, 255)
        
        time_text.alignment = PP_ALIGN.LEFT
        time_frame.word_wrap = True 
        time_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
        
        
        slide1.shapes.add_picture("clock.jpeg", Inches(7.47), Inches(2.63), Inches(1.17), Inches(1.13))
        
        text1_box = slide1.shapes.add_textbox(Inches(0), Inches(4.94), Inches(3.3), Inches(0.71))
        text1_frame = text1_box.text_frame
        
        text1 = text1_frame.add_paragraph()
        text1.text = f"Véleményünk:"
        text1.font.size = Pt(36)
        text1.font.bold = True
        text1.font.color.rgb = RGBColor(255, 255, 255)
        
        text1.alignment = PP_ALIGN.LEFT
        text1_frame.word_wrap = True 
        text1_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
        
        opinion_box = slide1.shapes.add_textbox(Inches(0), Inches(5.34), Inches(10), Inches(1.92))
        opinion_frame = opinion_box.text_frame
        
        opinion_text = opinion_frame.add_paragraph()
        opinion_text.text = f"{futam.opinion}"

        opinion_text.font.size = Pt(30)
        opinion_text.font.bold = False
        opinion_text.font.color.rgb = RGBColor(255, 255, 255)
        
        time_text.alignment = PP_ALIGN.LEFT
        opinion_frame.word_wrap = True 
        opinion_frame.vertical_anchor = MSO_ANCHOR.TOP

    def slide2(self,ppt,slide_layout,futam):
        drivers_list = [driver for driver in self.drivers if driver.futam == futam.id]

        slide2 = ppt.slides.add_slide(slide_layout)
        #slide2.slide_show.transition.duration = 50
        slide2bc = slide2.background.fill
        slide2bc.solid()
        slide2bc.fore_color.rgb = RGBColor(0, 0, 0)
        
        horse_box = slide2.shapes.add_textbox(Inches(0), Inches(0), Inches(5), Inches(7.5))
        horse_frame = horse_box.text_frame
        
        driver_box = slide2.shapes.add_textbox(Inches(5), Inches(0), Inches(5), Inches(7.5))
        driver_frame = driver_box.text_frame
        
        for row in drivers_list:
            
            #horse side

            horse = horse_frame.add_paragraph()
            horse.text = f"{row.Hnum}. {row.Hname.upper()}"
            #print(f"{row.horse_number}. {row.horse_name.upper()}")
            horse.font.size = Pt(32)
            horse.font.bold = True
            horse.font.color.rgb = RGBColor(255, 229, 121)
            
            horse.alignment = PP_ALIGN.LEFT
            horse_frame.word_wrap = True 
            horse_frame.vertical_anchor = MSO_ANCHOR.TOP
            
            #driver side
            
            driver = driver_frame.add_paragraph()
            driver.text = f"{row.Dname}"
            driver.font.size = Pt(32)
            driver.font.bold = False
            driver.font.color.rgb = RGBColor(255, 255, 255)
            
            driver.alignment = PP_ALIGN.LEFT
            driver_frame.word_wrap = True 
            driver_frame.vertical_anchor = MSO_ANCHOR.TOP

    def slide3(self,ppt,slide_layout,futam,hide=False):
        slide3 = ppt.slides.add_slide(slide_layout)
        #slide3.slide_show.transition.duration = 50
        slide3bc = slide3.background.fill
        slide3bc.solid()
        slide3bc.fore_color.rgb = RGBColor(0, 0, 0)
        

        
        title_box = slide3.shapes.add_textbox(Inches(0), Inches(0), Inches(10), Inches(1.4))
        title_frame = title_box.text_frame
        
        title = title_frame.add_paragraph()
        
        if futam.id < len(self.rome_num): title_str = f"{self.rome_num[futam.id]}. {futam.title}".strip()
        else:                             title_str = f"X{self.rome_num[futam.id-10]}. {futam.title}".strip()
        title.text = title_str
        if   (len(title_str) < 39 ): title.font.size = Pt(45)
        elif (len(title_str) <= 58): title.font.size = Pt(40)
        else:                        title.font.size = Pt(36)
        title.font.bold = True
        title.font.color.rgb = RGBColor(255, 229, 121)
        
        title.alignment = PP_ALIGN.CENTER
        title_frame.word_wrap = True 
        title_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            

            
        text_box = slide3.shapes.add_textbox(Inches(0), Inches(1.4), Inches(6.74), Inches(1.72))
        text_frame = text_box.text_frame
        
        text = text_frame.add_paragraph()
        text.text = f"Pálya: 10 Kincsem Park\nBefutási sorrend:"
        text.font.size = Pt(48)
        text.font.bold = True
        text.font.color.rgb = RGBColor(255, 229, 121)
        
        text.alignment = PP_ALIGN.LEFT
        text_frame.word_wrap = True 
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        order_box = slide3.shapes.add_textbox(Inches(0), Inches(3.37), Inches(10), Inches(2.52))
        order_frame = order_box.text_frame
        
        order = order_frame.add_paragraph()
        order.text = f"I.\nII.\nIII."
        order.font.size = Pt(48)
        order.font.bold = True
        order.font.color.rgb = RGBColor(255, 255, 255)
        
        order.alignment = PP_ALIGN.LEFT
        order_frame.word_wrap = True 
        order_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        time_box = slide3.shapes.add_textbox(Inches(0), Inches(5.9), Inches(3.39), Inches(0.91))
        time_frame = time_box.text_frame
        
        time_text = time_frame.add_paragraph()
        time_text.text = f"Idő:"

        time_text.font.size = Pt(48)
        time_text.font.bold = True
        time_text.font.color.rgb = RGBColor(255, 229, 121)
        
        time_text.alignment = PP_ALIGN.LEFT
        time_frame.word_wrap = True 
        time_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
        
    def slide4(self,ppt,slide_layout,futam,hide=False):
        slide4 = ppt.slides.add_slide(slide_layout)
        #slide4.slide_show.transition.duration = 50
        slide4bc = slide4.background.fill
        slide4bc.solid()
        slide4bc.fore_color.rgb = RGBColor(0, 0, 0)
        
        title_box = slide4.shapes.add_textbox(Inches(0), Inches(0), Inches(10), Inches(1.4))
        title_frame = title_box.text_frame
        title = title_frame.add_paragraph()

            
        if futam.id < len(self.rome_num): title_str = f"{self.rome_num[futam.id]}. {futam.title}".strip()
        else:                             title_str = f"X{self.rome_num[futam.id-10]}. {futam.title}".strip()
        title.text = title_str
        if   (len(title_str) < 39 ): title.font.size = Pt(45)
        elif (len(title_str) <= 58): title.font.size = Pt(40)
        else:                        title.font.size = Pt(36)
        title.font.bold = True
        title.font.color.rgb = RGBColor(255, 229, 121)
        
        title.alignment = PP_ALIGN.CENTER
        title_frame.word_wrap = True 
        title_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            
        
        text_box = slide4.shapes.add_textbox(Inches(0), Inches(1.4), Inches(10), Inches(1.72))
        text_frame = text_box.text_frame
        
        text = text_frame.add_paragraph()
        text.text = f"Pálya: 10 Kincsem Park\nBefutási sorrend:"
        text.font.size = Pt(48)
        text.font.bold = True
        text.font.color.rgb = RGBColor(255, 229, 121)
        
        text.alignment = PP_ALIGN.LEFT
        text_frame.word_wrap = True 
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        dividend_box = slide4.shapes.add_textbox(Inches(0), Inches(3.37), Inches(10), Inches(2.52))
        dividend_frame = dividend_box.text_frame
        
        dividend = dividend_frame.add_paragraph()
        dividend.text = f"Tét:					1\nHely:					1\nBefutó:    		1\nHbefutó: 		1"
        dividend.font.size = Pt(48)
        dividend.font.bold = True
        dividend.font.color.rgb = RGBColor(255, 255, 255)
        
        dividend.alignment = PP_ALIGN.LEFT
        dividend_frame.word_wrap = True 
        dividend_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        
    def slide5(self,ppt,slide_layout):
        slide5 = ppt.slides.add_slide(slide_layout)
        #slide5.slide_show.transition.duration = 50
        slide5bc = slide5.background.fill
        slide5bc.solid()
        slide5bc.fore_color.rgb = RGBColor(0, 0, 0)
        
        title_box = slide5.shapes.add_textbox(Inches(0), Inches(0), Inches(10), Inches(2.3))
        title_frame = title_box.text_frame
        
        title = title_frame.add_paragraph()
        title.text = f"Nem hivatalos\n  befutási sorrend:"
        title.font.size = Pt(88)
        title.font.bold = True
        title.font.color.rgb = RGBColor(255, 229, 121)
        
        title.alignment = PP_ALIGN.LEFT
        title_frame.word_wrap = True 
        title_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        text_box = slide5.shapes.add_textbox(Inches(0), Inches(3.75), Inches(10), Inches(2.21))
        text_frame = text_box.text_frame
        
        text = text_frame.add_paragraph()
        text.text = f" –  – "
        text.font.size = Pt(125)
        text.font.bold = True
        text.font.color.rgb = RGBColor(255, 255, 255)
        
        text.alignment = PP_ALIGN.CENTER
        text_frame.word_wrap = True 
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    

    def run_vba_macro(self):
        wb = wx.Book("add macro.xlsm")
        macro1 = wb.macro("Module1.SetPPTSlidesFromFolderAndExitExcel")
        macro1()
    
    def inputbox(self, title, message, inputmess = ""):
        # Create a custom dialog
        dialog = QDialog(self)
        dialog.setWindowTitle(title)

        # Set up layout
        layout = QVBoxLayout()

        # Add error message label
        label = QLabel(message)
        layout.addWidget(label)

        # Add input field
        input_field = QLineEdit()
        input_field.setPlaceholderText(inputmess)
        layout.addWidget(input_field)

        # Add Accept button
        button = QPushButton("Accept")
        button.clicked.connect(dialog.accept)  # Close dialog on button click
        layout.addWidget(button)

        dialog.setLayout(layout)

        # Execute the dialog and check if the user accepted
        if dialog.exec_() == QDialog.Accepted:
            return input_field.text()  # Return the input from the user
        return None

    


    def make_ppt(self):
        #Horse Number;Horse Name;Horse Distance;Driver Name;Futam Number;is run
        class Drivers:
            def __init__(self,line):
                Hnum,Hname,Hdist,Dname,futam,isRun = line.strip().split(";")
                self.ln = line.strip().split(";")
                self.Hnum  = int(Hnum)
                self.Hname = Hname
                self.Hdist = int(Hdist)
                self.Dname = Dname
                self.futam = int(futam)
                self.isRun = int(isRun) == 1

        
        #Id;Title;Distance;Time;Start type;Opinion
        class Titles:
            def __init__(self,line):
                id,title,dist,time,stype,opinion = line.strip().split(";")
                self.ln = line.strip().split(";")
                self.id      = int(id)
                self.title   = title
                self.dist    = int(dist)
                self.time    = time
                self.stype   = stype
                self.opinion = opinion
                
        with open("csv/drivers_data.csv","r",encoding="utf-8") as f:
            first_line = f.readline()
            self.drivers =[Drivers(line) for line in f]

        with open("csv/titles_data.csv","r",encoding="utf-8") as f:
            first_line = f.readline()
            self.titles = [Titles(line) for line in f]

        #for i in self.titles: print("; ".join(i.ln))

        if not os.path.isdir("ppt"):
            os.makedirs("ppt")

        for title in self.titles:
            ppt = Presentation()
            
            ppt.slide_width = Inches(10)  # Set width
            ppt.slide_height = Inches(7.5)  # Set height
            slide_layout = ppt.slide_layouts[6]

            self.slide1(ppt,slide_layout,title)
            self.slide2(ppt,slide_layout,title)

            if(title.id==0):
                self.slide3(ppt,slide_layout,title)
                self.slide4(ppt,slide_layout,title,True)

            elif(self.titles[-1].id == title.id):
                self.slide3(ppt,slide_layout,self.titles[title.id-1])
                self.slide4(ppt,slide_layout,self.titles[title.id-1])
                self.slide3(ppt,slide_layout,title,True)
                self.slide4(ppt,slide_layout,title,True)

            else:
                self.slide3(ppt,slide_layout,self.titles[title.id-1])
                self.slide4(ppt,slide_layout,self.titles[title.id-1])

            self.slide5(ppt,slide_layout)

            file_name = ""
            if(title.id < len(self.rome_num)):
                file_name =f".\\ppt\\{self.rome_num[title.id]}. futam.pptx"
            else:
                file_name =f".\\ppt\\X{self.rome_num[title.id-10]}. futam.pptx"
            
            ppt.save(file_name)

        self.run_vba_macro()
        
        QMessageBox.information(self, 'Success', 'PowerPoint file created successfully!')



if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = PDFProcessor()
    ex.show()
    sys.exit(app.exec_())
