from PyQt5.QtWidgets import QApplication
import sys
from MyWindow import MyWindow

# Create Window
app = QApplication(sys.argv)
window = MyWindow()
window.show()
sys.exit(app.exec_())
