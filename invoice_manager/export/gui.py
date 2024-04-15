import sys
import os

# current_dir = os.path.dirname(os.path.abspath(__file__))
# sys.path.insert(0, current_dir)

sys.path.append(os.path.join(os.path.dirname(__file__), ".."))
# current path
sys.path.append(os.path.dirname(__file__))

from PyQt5.QtWidgets import QApplication, QWidget
from custom_form import CustomForm


class MyApp(CustomForm):
    def __init__(self):
        super(MyApp, self).__init__()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
