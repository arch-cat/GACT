from PyQt5.QtWidgets import QApplication, QStyledItemDelegate, QTableWidget
from PyQt5.QtGui import QColor, QPalette


class ColorDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        if index.data() == 'Online':
            option.palette.setColor(QPalette.Text, QColor("green"))
        elif index.data() == 'Offline':
            option.palette.setColor(QPalette.Text, QColor("red"))
        QStyledItemDelegate.paint(self, painter, option, index)


if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    w = QTableWidget(3, 3)
    w.setItemDelegate(ColorDelegate())
    w.show()
    sys.exit(app.exec_())