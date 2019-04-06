from pywinauto import Application as WinApplication
from fixture.group import GH

class Application:

    def __init__(self, target):
        self.application = WinApplication(backend="win32").start(target)
        self.main_window = self.application.window(title="Free Address Book")
        self.main_window.wait("visible")
        self.groups = GH(self)

    def destroy(self):
        self.application.window(title="Free Address Book").close()