import copy
import inspect
import os
import sys
from collections import defaultdict

import lxml
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import QHeaderView, QMessageBox, QProgressBar, QMenu

filename = inspect.getframeinfo(inspect.currentframe()).filename
path_programm = os.path.dirname(os.path.abspath(filename))

path_other_files = 'other_files'
path_xml_inter = f'{path_programm}/{path_other_files}/test_inter.xml'
path_excel = f'{path_programm}/{path_other_files}/XML_View.xlsx'
f = open(f'{path_programm}/{path_other_files}/darksource.stylesheet', 'r')
style = f.read()

import re
import xlsxwriter
import openpyxl
from lxml import etree
from widgets import widj_tab1


class XML_Editor(QThread):
    signal_end_open = pyqtSignal('PyQt_PyObject')
    signal_end_export = pyqtSignal('PyQt_PyObject')
    signal_end_import = pyqtSignal('PyQt_PyObject')
    error = pyqtSignal('PyQt_PyObject')

    def __init__(self):
        QThread.__init__(self)
        self.process_num = 0

    def set_current_process(self, process_num=0):
        # Номер процесса 1-открыть 2-to Excel 3-from Excel
        self.process_num = process_num

    def set_path_xml(self, path_xml):
        self.path_xml = path_xml

    def set_parents_dict(self, parents_dict):
        self.item_parents_dict = parents_dict

    # run method gets called when we start the thread
    def run(self):
        try:
            if self.process_num == 1:
                #     Получение Excel
                nice_tree = self.get_attrib(new_open=True)
                self.signal_end_open.emit(nice_tree)
            elif self.process_num == 2:
                self.toExcel()
                nice_tree = self.get_attrib(new_open=False)
                self.signal_end_export.emit(nice_tree)
            elif self.process_num == 3:
                self.fromExcel()
                nice_tree = self.get_attrib(new_open=False)
                self.signal_end_import.emit(nice_tree)
            else:
                self.error.emit(f"Режим не выбран")
        except Exception as err:
            self.error.emit(f"Упс, ошибочка вышла :(, \n {err.__str__()}")

    def get_chil(self, prefix, parent):
        for element in parent.getchildren():
            if element == parent:
                continue
            new_name_tag = self.find_new_atrname(f'{prefix}{element.tag}')
            for atr in element.attrib:
                self.all_data[f'{new_name_tag}@{atr}'].update({self.id_atribute: element.attrib[atr]})

            if element.text:
                self.all_data[f'{new_name_tag}#text'].update({self.id_atribute: element.text})
            else:
                self.all_data[f'{new_name_tag}#text'].update({self.id_atribute: ''})

            self.get_chil(prefix=f'{new_name_tag}+', parent=element)

    def add_tag(self, parent, tag_name, attr_name, attr_value, text):
        new_tag_name = tag_name.split('№')
        subtag = lxml.etree.SubElement(parent, new_tag_name[0])
        # subtag.text = ''
        return subtag

    def add_property(self, parent, data):
        for index in range(max(data[list(data.keys())[0]].keys())):
            dict_data = {}
            for key in data.keys():
                data_value = data[key].get(index + 1, 666)
                if data_value == 666:
                    continue
                attr, text = None, None
                if '#text' in key:
                    tags_full, text = key.split('#text')
                else:
                    tags_full, attr = key.split('@')

                tags = tags_full.split('+')

                element = dict_data.get(tags_full, None)
                if element is None:
                    parent_tag = dict_data.get('+'.join(tags[0:len(tags) - 1])) if len(tags) > 1 else parent

                    element = self.add_tag(parent=parent_tag,
                                           tag_name=tags[-1], attr_name=attr, attr_value=data_value,
                                           text=text)

                if text is None:
                    element.set(attr, str(data_value) if data_value is not None and data_value != 'None' else '')
                else:
                    element.text = data_value if data_value is not None and data_value != 'None' else ''

                dict_data[tags_full] = element

    def fromExcel(self):
        wb = openpyxl.load_workbook(filename=path_excel)
        for index_sheet, sheet in enumerate(wb.worksheets):
            excel_data = defaultdict(dict)
            for coll in sheet.columns:
                for index, cell in enumerate(coll):
                    if index == 0:
                        field_name = cell.internal_value
                        continue
                    if cell.internal_value == 'None':
                        continue
                    excel_data[field_name].update({index: str(cell.internal_value)})

            self.add_property(parent=self.parents_list[index_sheet], data=excel_data)
        mydata = lxml.etree.tostring(self.raw_tree, pretty_print=True, method="xml", encoding='unicode')
        mydata = mydata.replace('(/n]', '\n')
        mydata = '<?xml version="1.0" encoding="UTF-8"?>\n' + mydata

        with open(self.path_xml, "w", encoding='utf8') as fh:
            fh.write(mydata)

    def get_selected_attr(self, item):
        selected_item = []
        for child in self.item_parents_dict[item]:
            try:
                element = self.raw_tree.xpath(self.mainDict[child])[0]
                selected_item.append(element)
            except:
                continue
        return selected_item

    def find_new_atrname(self, tag_name):
        index = 0
        new_tag_name = tag_name
        while new_tag_name in self.all_tags:
            index += 1
            new_tag_name = f'{tag_name}№{index}'
        self.all_tags.add(new_tag_name)
        return new_tag_name

    def get_children(self, item):

        self.all_data = defaultdict(dict)
        parent = self.raw_tree.xpath(self.mainDict[item])[0]
        self.id_atribute = 0
        element_list = self.get_selected_attr(item)

        for element in element_list:
            # if element not in element_list:
            #     continue
            self.all_tags = set()
            self.id_atribute += 1
            new_name_tag = self.find_new_atrname(element.tag)
            for atr in element.attrib:
                self.all_data[f'{new_name_tag}@{atr}'].update({self.id_atribute: element.attrib[atr]})

            self.all_data[f'{new_name_tag}#text'].update({self.id_atribute: element.text if element.text else ''})

            self.get_chil(prefix=f'{new_name_tag}+', parent=element)

            try:
                parent.remove(element)
            except Exception as err:
                pass

        self.add_new_sheet_Excel()

    def add_new_sheet_Excel(self):
        worksheet = self.workbook.add_worksheet()
        max_row = 0
        col = 0
        for key in self.all_data.keys():
            worksheet.write(0, col, key)
            data_dict = self.all_data[key]
            max_row = max(max(data_dict.keys()), max_row)
            for row in range(1, max_row + 1):
                data = data_dict.get(row, 'None')
                worksheet.write(row, col, data)
            col += 1

    def toExcel(self):
        self.parents_list = []
        self.workbook = xlsxwriter.Workbook(path_excel)
        for item in self.item_parents_dict.keys():
            parent = self.raw_tree.xpath(self.mainDict[item])[0]
            self.parents_list.append(parent)
            self.get_children(item)
        self.workbook.close()

    def add_tree(self, parent, text):
        child = QtWidgets.QTreeWidgetItem(parent)
        child.setText(0, text)

    def replace_n(self):
        with open(self.path_xml, "r", encoding='utf8') as fh:
            data = fh.readlines()
            data_xml = ''.join(data)
            elements_with = re.findall(r'''="[^"]*[\n]+[^"]*"''', data_xml)

            # data_xml = data_xml.replace('\n', '')
            for element in elements_with:
                element_new = element.replace('\n', '(/n]')
                data_xml = data_xml.replace(element, element_new)

        with open(path_xml_inter, "w", encoding='utf8') as fh:
            fh.write(data_xml)

    def get_attrib(self, new_open):
        self.mainDict = {}
        self.id_dict = 0
        parser = lxml.etree.XMLParser(remove_blank_text=True)
        if new_open:
            self.replace_n()
            self.raw_tree = lxml.etree.parse(path_xml_inter, parser=parser)
        nice_tree = []

        for tag in self.raw_tree.iter():
            path = re.sub(r'\[[0-9]+\]', '', self.raw_tree.getpath(tag))
            path = f"{path}-{tag.attrib.get('Name', tag.attrib.get('ID', tag.text if tag.text else ''))}"
            nice_tree.append([path, self.id_dict, [attrib for attrib in tag.keys()]])
            self.mainDict[self.id_dict] = self.raw_tree.getpath(tag)
            self.id_dict += 1

        return nice_tree


class TreeCustom(QtWidgets.QTreeWidget):

    def __init__(self):
        super().__init__()
        self.setAlternatingRowColors(True)
        self.setColumnCount(3)
        self.setHeaderLabels(['Parent', 'Path', 'Edit'])
        self.setDropIndicatorShown(False)
        myQHeaderView = self.header()
        myQHeaderView.setStretchLastSection(True)
        myQHeaderView.setSectionResizeMode(0, QHeaderView.Stretch)  # ResizeToContents)
        myQHeaderView.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # ResizeToContents)
        myQHeaderView.setSectionResizeMode(2, QHeaderView.Stretch)  # ResizeToContents)

    def contextMenuEvent(self, event):
        contextMenu = QMenu(self)
        select = contextMenu.addAction("Выделить все дочерние элементы")
        unselect = contextMenu.addAction("Снять выделение с всех дочерних элементов")
        action = contextMenu.exec_(self.mapToGlobal(event.pos()))
        if action == select:
            self.select_children()
        if action == unselect:
            self.unselect_children()

    def unselect_children(self):
        for item in self.findItems("", Qt.MatchContains | Qt.MatchRecursive):
            if item.checkState(0) == 2:
                for index in range(item.childCount()):
                    child = item.child(index)
                    child.setCheckState(2, 0)

    def select_children(self):
        for item in self.findItems("", Qt.MatchContains | Qt.MatchRecursive):
            if item.checkState(0) == 2:
                for index in range(item.childCount()):
                    child = item.child(index)
                    child.setCheckState(2, 2)


class Tab(QtWidgets.QWidget):
    xml_editor_clc = XML_Editor()

    def __init__(self):
        super().__init__()

        self.xml_editor_clc.signal_end_open.connect(self.tree_completion)
        self.xml_editor_clc.signal_end_import.connect(self.import_done)
        self.xml_editor_clc.signal_end_export.connect(self.export_done)
        self.xml_editor_clc.error.connect(self.some_error_in_thread)

        self.setStyleSheet(style)
        layout = QtWidgets.QGridLayout(self)
        self.tree = TreeCustom()

        layout.addWidget(self.tree, 1, 0, 30, 1)

        self.progressbar = QProgressBar(textVisible=True)

        btn_export = QtWidgets.QPushButton('To Excel')
        btn_export.clicked.connect(self.toExcel)
        btn_import = QtWidgets.QPushButton('From Excel')
        btn_import.clicked.connect(self.fromExcel)
        btn_open = QtWidgets.QPushButton('Open Excel')
        btn_open.clicked.connect(self.openExcel)

        self.xml_win = widj_tab1.line_btn_xml_path(self)
        self.xml_win.my_signal.connect(self.tree_fill)

        # self.xlsx_win = widj_tab1.line_btn_xls_path(self)
        # self.xlsx_win.my_signal.connect(self.xls_path_change)

        layout.addWidget(self.progressbar, 0, 0, 1, 2)
        layout.addWidget(self.xml_win, 1, 1, 1, 1)
        # layout.addWidget(self.xlsx_win, 2, 1, 1, 1)
        layout.addWidget(btn_export, 6, 1, 1, 1)
        layout.addWidget(btn_import, 7, 1, 1, 1)
        layout.addWidget(btn_open, 8, 1, 1, 1)
        self.setLayout(layout)

    # def xls_path_change(self, xlsx_path):
    #     self.xlsx_path = xlsx_path

    def openExcel(self):
        try:
            # открыть
            os.system(f'start excel.exe {path_excel}')
        except:
            os.system(f'start {path_other_files}')

    def progress_bar_status(self, status=0):
        if status == 1:
            # Stop
            self.progressbar.setMaximum(100)
            self.progressbar.setMinimum(0)
            self.progressbar.setValue(100)
        elif status == 2:
            # Start
            self.progressbar.setMaximum(0)
            self.progressbar.setMinimum(0)
            self.progressbar.setValue(100)
        else:
            # Reset
            self.progressbar.setMaximum(100)
            self.progressbar.setMinimum(0)
            self.progressbar.setValue(0)

    def some_error_in_thread(self, error_mess):
        self.progress_bar_status(0)
        QMessageBox().warning(self, "Ошибка", error_mess, QMessageBox.Ok,
                              QMessageBox.Ok)

    def tree_fill(self, path_xml):
        self.xml_editor_clc.set_path_xml(path_xml)
        self.xml_editor_clc.set_current_process(process_num=1)
        self.progress_bar_status(2)
        self.xml_editor_clc.start()

    def tree_redraw(self, nice_tree):
        self.tree.clear()
        self.tree.expandAll()
        parent = self.tree
        pre_indent = 0
        item = self.tree
        for path, id_path, attribs in nice_tree:
            indent = int(path.count('/') - 1)
            if indent > pre_indent:
                parent = item
            elif indent < pre_indent:
                while indent - 1 != pre_indent:
                    item = item.parent()
                    pre_indent -= 1
                parent = item

            item = QtWidgets.QTreeWidgetItem(parent)
            item.setCheckState(0, 0)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setText(0, '{0}: {1} [{2}]'.format(indent, path.split('/')[-1],
                                                    ', '.join(attribs) if len(attribs) > 0 else '-'))

            item.setText(1, str(id_path))
            item.setCheckState(2, 0)
            pre_indent = copy.deepcopy(indent)
            if id_path == 0:
                self.tree.expandItem(item)

    def tree_completion(self, nice_tree):
        self.progress_bar_status(1)
        self.tree_redraw(nice_tree)
        self.progress_bar_status(0)

    def import_done(self, nice_tree):
        self.tree_redraw(nice_tree)
        self.progress_bar_status(0)
        QMessageBox().information(self, 'Готово', 'Изменения приняты', QMessageBox.Ok)
        self.progress_bar_status(1)

    def export_done(self, nice_tree):
        self.tree_redraw(nice_tree)
        self.openExcel()
        self.progress_bar_status(1)
        QMessageBox().information(self, 'Готово', 'Экспорт в Excel', QMessageBox.Ok)
        self.progress_bar_status(0)

    def get_child_edit(self, item):
        selected_item = []
        for index in range(item.childCount()):
            child = item.child(index)
            if child.checkState(2) != 2:
                continue
            try:
                selected_item.append(int(child.text(1)))
            except:
                continue
        return selected_item

    def toExcel(self):
        self.xml_editor_clc.set_current_process(process_num=2)
        self.progress_bar_status(2)
        parents_dict = {}
        for item in self.tree.findItems("", Qt.MatchContains | Qt.MatchRecursive):
            if item.checkState(0) == 2:
                parents_dict.update({int(item.text(1)): self.get_child_edit(item)})
        self.xml_editor_clc.set_parents_dict(parents_dict=parents_dict)
        self.xml_editor_clc.start()

    def fromExcel(self):
        self.xml_editor_clc.set_current_process(process_num=3)
        self.progress_bar_status(2)
        self.xml_editor_clc.start()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    root = Tab()
    root.showMaximized()
    sys.exit(app.exec_())
