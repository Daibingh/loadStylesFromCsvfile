# -*- coding: utf-8 -*-

import sys
from PyQt5.QtWidgets import QApplication
from StylesLoader import StylesLoader

if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = StylesLoader()
    w.show()
    sys.exit(app.exec_())