from comtypes.client import CreateObject
import os
import random

class GH:
    def __init__(self, app):
        self.app = app

    def get_group_list(self):
        group = []
        self.open_group_editor()
        tree = self.group_editor.window(auto_id="uxAddressTreeView")
        root = tree.tree_root()
        group = [node.text() for node in root.children()]
        self.close_group_editor()
        return group

    def add_new_group(self,name):
        self.open_group_editor()
        self.group_editor.window(auto_id="uxNewAddressButton").click()
        input = self.group_editor.window(class_name="Edit")
        input.set_text(name)
        input.type_keys("\n")
        self.close_group_editor()

    def delete_new_group(self,name):
        self.open_group_editor()
        tree = self.group_editor.window(auto_id="uxAddressTreeView")
        root = tree.tree_root()
        for node in root.children():
            if node.text() == name:
                node.click()
        self.group_editor.window(auto_id="uxDeleteAddressButton").click()
        dialog = self.app.application.window(title="Delete group")
        dialog.window(auto_id="uxOKAddressButton").click()
        self.close_group_editor()

    def open_group_editor(self):
        self.app.main_window.window(auto_id="groupButton").click()
        self.group_editor = self.app.application.window(title="Group editor")
        self.group_editor.wait("visible")

    def close_group_editor(self):
        self.group_editor.close()

    def read_excel(self):
        project_dir = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
        xl = CreateObject("Excel.application")
        xl.Workbooks.Open(os.path.join(project_dir,"groups.xlsx"))
        xl.Visible = True
        sheet = xl.Workbooks[1].Sheets["Лист1"]
        list = []
        for i in sheet.Range("A1", "A10"):
            list.append(i.Value())
        item = random.choice(list)
        xl.Quit()
        return item