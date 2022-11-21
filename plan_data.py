import os.path

# Xml Handler
from openpyxl import load_workbook

# Outlook Handler
import win32com.client


class Data:

    def __init__(self, path_file):
        self.path_file = path_file

    def load_plan(self):
        return load_workbook(self.path_file)


class EmailSender:

    def __init__(self, to, subject, html_body, attach):
        self.to = to
        self.subject = subject
        self.html_body = html_body
        self.attach = attach

    def sender(self):
        outlook = win32com.client.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        email.Subject = self.subject
        email.To = self.to
        email.HTMLBody = self.html_body
        email.Attachments.Add(os.path.join(os.getcwd(), self.attach))
        email.display()


class ClientObj:

    def __init__(self, to, subject, html_body, attach):
        self.to = to
        self.subject = subject
        self.html_body = html_body
        self.attach = attach

    def client_obj_create(self):
        return ClientObj(self.to, self.subject, self.html_body, self.attach)
