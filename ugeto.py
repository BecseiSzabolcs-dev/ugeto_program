#!/usr/bin/env python3
"""
ugeto_ui.py
PyQt6 GUI for the Ügető program (titles + drivers editing, load PDF placeholder, save CSV, make PPT)

Dependencies:
  pip install PyQt6 pandas python-pptx

Run:
  python ugeto_ui.py

This file implements a self-contained UI. PDF parsing and advanced PPT layout are left as clear hooks
so you can plug your existing extraction logic later.
"""

import sys
import os
import csv
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QMessageBox, QLabel, QLineEdit, QTabWidget,
    QTableView, QHeaderView, QAbstractItemView, QToolBar, QStatusBar
)
from PyQt6.QtGui import QAction, QKeySequence
from PyQt6.QtCore import Qt, QSortFilterProxyModel, QModelIndex
from PyQt6.QtGui import QStandardItemModel, QStandardItem
from modules import *
# Optional libraries
try:
    import pandas as pd
except Exception:
    pd = None

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
except Exception:
    Presentation = None


CSV_DIR = Path("csv")
PPT_DIR = Path("ppt")
CSV_DIR.mkdir(exist_ok=True)
PPT_DIR.mkdir(exist_ok=True)


class EditableTable(QWidget):
    """A small reusable widget: search box + table + add/delete buttons"""

    def __init__(self, columns, parent=None):
        super().__init__(parent)
        self.columns = columns
        self.model = QStandardItemModel(0, len(columns), self)
        self.model.setHorizontalHeaderLabels(columns)

        self.proxy = QSortFilterProxyModel(self)
        self.proxy.setSourceModel(self.model)
        self.proxy.setFilterKeyColumn(-1)  # filter all columns

        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.SelectedClicked)

        self.search = QLineEdit()
        self.search.setPlaceholderText("Keresés...")
        self.search.textChanged.connect(self.on_search)

        add_btn = QPushButton("Sor hozzáadása")
        del_btn = QPushButton("Sor törlése")
        add_btn.clicked.connect(self.add_row)
        del_btn.clicked.connect(self.delete_selected_row)

        h = QHBoxLayout()
        h.addWidget(self.search)
        h.addWidget(add_btn)
        h.addWidget(del_btn)

        layout = QVBoxLayout(self)
        layout.addLayout(h)
        layout.addWidget(self.table)

    def on_search(self, text):
        self.proxy.setFilterFixedString(text)

    def add_row(self, values=None):
        if values is None:
            values = ["" for _ in self.columns]
        row = [QStandardItem(str(v)) for v in values]
        for it in row:
            it.setEditable(True)
        self.model.appendRow(row)
        # select newly added row in the view
        idx = self.model.index(self.model.rowCount() - 1, 0)
        proxy_idx = self.proxy.mapFromSource(idx)
        self.table.scrollTo(proxy_idx)
        self.table.setCurrentIndex(proxy_idx)
        self.table.edit(proxy_idx)

    def delete_selected_row(self):
        sel = self.table.selectionModel().currentIndex()
        if not sel.isValid():
            QMessageBox.information(self, "Törlés", "Nincs kiválasztott sor.")
            return
        src = self.proxy.mapToSource(sel)
        self.model.removeRow(src.row())

    def to_list_of_dicts(self):
        rows = []
        for r in range(self.model.rowCount()):
            d = {}
            for c, name in enumerate(self.columns):
                it = self.model.item(r, c)
                d[name] = it.text() if it is not None else ""
            rows.append(d)
        return rows

    def load_from_list_of_dicts(self, data):
        self.model.removeRows(0, self.model.rowCount())
        for row in data:
            vals = [row.get(k, "") for k in self.columns]
            self.add_row(vals)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ügető program")
        self.resize(1100, 700)

        # Toolbar
        toolbar = QToolBar()
        self.addToolBar(toolbar)

        load_pdf_act = QAction("Load PDF", self)
        load_pdf_act.setShortcut(QKeySequence("Ctrl+O"))
        load_pdf_act.triggered.connect(self.load_pdf)

        save_csv_act = QAction("Save Data to CSV", self)
        save_csv_act.setShortcut(QKeySequence("Ctrl+S"))
        save_csv_act.triggered.connect(self.save_csv)

        make_ppt_act = QAction("Make PPT", self)
        make_ppt_act.triggered.connect(self.make_ppt)

        toolbar.addAction(load_pdf_act)
        toolbar.addAction(save_csv_act)
        toolbar.addAction(make_ppt_act)

        # Tabs for Titles and Drivers
        self.tabs = QTabWidget()

        titles_columns = ["Azonosító", "Cím", "Táv", "Időpont", "Start típusa", "Vélemény"]
        drivers_columns = ["Lószám", "Lónév", "Táv", "Hajtó neve", "Futam száma", "Futott-e"]

        self.titles_widget = EditableTable(titles_columns)
        self.drivers_widget = EditableTable(drivers_columns)

        self.tabs.addTab(self.titles_widget, "Titles (Címek)")
        self.tabs.addTab(self.drivers_widget, "Drivers (Hajtók)")

        # Central layout
        central = QWidget()
        main_layout = QVBoxLayout(central)
        main_layout.addWidget(self.tabs)

        # Info label and status bar
        info = QLabel("PDF betöltése után töltsd fel az adatokat, szerkeszd, majd mentsd CSV-be és készíts PPT-t.")
        main_layout.addWidget(info)

        self.setCentralWidget(central)
        self.status = QStatusBar()
        self.setStatusBar(self.status)

        # Shortcuts: pressing any printable key will insert into selected cell
        self.table_widgets = [self.titles_widget.table, self.drivers_widget.table]

    def load_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF files (*.pdf);;All files (*)")
        if not path:
            return
        self.status.showMessage(f"Betöltött PDF: {path}")

        pdf_data = ReadPDF(path)
        

        """
        for ln in data.futam_data:
            title = Futam(ln)
            print(f"{title}")
            for horse in ln["participants"]:
                print(f"        {Horses(horse,title.id)}")
        """


        #self.titles_widget.load_from_list_of_dicts(titles_sample)
        #self.drivers_widget.load_from_list_of_dicts(drivers_sample)
        QMessageBox.information(self, "PDF betöltve", "A PDF feldolgozása kész (példaadatok hozzáadva). Cseréld ki a parse_pdf helyen a logikát.")

    def save_csv(self):
        # Ensure directories exist
        CSV_DIR.mkdir(exist_ok=True)
        titles = self.titles_widget.to_list_of_dicts()
        drivers = self.drivers_widget.to_list_of_dicts()

        titles_file = CSV_DIR / "titles_data.csv"
        drivers_file = CSV_DIR / "drivers_data.csv"

        # Use pandas if available for nicer formatting, else csv module
        try:
            if pd is not None:
                pd.DataFrame(titles).to_csv(titles_file, index=False, encoding='utf-8')
                pd.DataFrame(drivers).to_csv(drivers_file, index=False, encoding='utf-8')
            else:
                # Write titles
                with titles_file.open('w', encoding='utf-8', newline='') as f:
                    writer = csv.DictWriter(f, fieldnames=self.titles_widget.columns)
                    writer.writeheader()
                    writer.writerows(titles)
                # Write drivers
                with drivers_file.open('w', encoding='utf-8', newline='') as f:
                    writer = csv.DictWriter(f, fieldnames=self.drivers_widget.columns)
                    writer.writeheader()
                    writer.writerows(drivers)
            self.status.showMessage(f"CSV-ek mentve: {titles_file} és {drivers_file}")
            QMessageBox.information(self, "CSV mentés", f"A fájlok sikeresen mentve:\n{titles_file}\n{drivers_file}")
        except Exception as e:
            QMessageBox.critical(self, "Hiba", f"CSV mentése sikertelen: {e}")

    def make_ppt(self):
        if Presentation is None:
            QMessageBox.warning(self, "pptx hiányzik", "A python-pptx csomag nincs telepítve. Telepítsd: pip install python-pptx")
            return

        PPT_DIR.mkdir(exist_ok=True)
        titles = self.titles_widget.to_list_of_dicts()
        drivers = self.drivers_widget.to_list_of_dicts()

        # Create a ppt for each title (futam)
        try:
            for t in titles:
                prs = Presentation()
                slide = prs.slides.add_slide(prs.slide_layouts[5])  # blank
                left = Inches(0.5)
                top = Inches(0.5)
                width = Inches(9)
                height = Inches(1.0)
                title_box = slide.shapes.add_textbox(left, top, width, height)
                tf = title_box.text_frame
                tf.text = f"{t.get('Azonosító','')} - {t.get('Cím','')} ({t.get('Táv','')} m)"
                p = tf.add_paragraph()
                p.text = f"Időpont: {t.get('Időpont','')}  Start: {t.get('Start típusa','')}"

                # Add participants table
                part_slide = prs.slides.add_slide(prs.slide_layouts[5])
                rows = 1 + sum(1 for d in drivers if d.get('Futam száma') == t.get('Azonosító'))
                cols = len(self.drivers_widget.columns)
                table_shape = part_slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.0), Inches(9), Inches(4)).table
                # header
                for c, name in enumerate(self.drivers_widget.columns):
                    table_shape.cell(0, c).text = name
                r = 1
                for d in drivers:
                    if d.get('Futam száma') == t.get('Azonosító'):
                        for c, name in enumerate(self.drivers_widget.columns):
                            table_shape.cell(r, c).text = str(d.get(name, ''))
                        r += 1

                # Save with the title name safe-filename
                safe = ''.join(ch for ch in f"{t.get('Azonosító','')}_{t.get('Cím','')}" if ch.isalnum() or ch in ' _-').strip()
                filename = PPT_DIR / f"{safe}.pptx"
                prs.save(filename)

            QMessageBox.information(self, "PPT kész", f"PPT-k elkészítve a mappában: {PPT_DIR.resolve()}")
            self.status.showMessage("PPT elkészítve")
        except Exception as e:
            QMessageBox.critical(self, "Hiba PPT", f"Hiba a PPT létrehozásakor: {e}")

    # Optional: override keyPressEvent to insert text to the selected cell when typing
    def keyPressEvent(self, event):
        key = event.key()
        if event.text() and event.text().isprintable():
            for table in self.table_widgets:
                if table.hasFocus():
                    sel = table.selectionModel().currentIndex()
                    if sel.isValid():
                        src = table.model().mapToSource(sel) if isinstance(table.model(), QSortFilterProxyModel) else sel
                        item = self.titles_widget.model.itemFromIndex(src) if table is self.titles_widget.table else self.drivers_widget.model.itemFromIndex(src)
                        if item is not None:
                            item.setText(item.text() + event.text())
                            return
        super().keyPressEvent(event)


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
