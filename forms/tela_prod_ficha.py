# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'tela_prod_ficha.ui'
#
# Created by: PyQt5 UI code generator 5.15.10
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1100, 720)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setContentsMargins(6, 6, 6, 6)
        self.verticalLayout.setObjectName("verticalLayout")
        self.widget_cabecalho = QtWidgets.QWidget(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.widget_cabecalho.sizePolicy().hasHeightForWidth())
        self.widget_cabecalho.setSizePolicy(sizePolicy)
        self.widget_cabecalho.setMinimumSize(QtCore.QSize(0, 40))
        self.widget_cabecalho.setMaximumSize(QtCore.QSize(16777215, 40))
        self.widget_cabecalho.setStyleSheet("background-color: rgb(0, 170, 255);")
        self.widget_cabecalho.setObjectName("widget_cabecalho")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.widget_cabecalho)
        self.horizontalLayout.setContentsMargins(15, 6, 15, 6)
        self.horizontalLayout.setSpacing(0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label_3 = QtWidgets.QLabel(self.widget_cabecalho)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_3.sizePolicy().hasHeightForWidth())
        self.label_3.setSizePolicy(sizePolicy)
        self.label_3.setMinimumSize(QtCore.QSize(0, 0))
        self.label_3.setMaximumSize(QtCore.QSize(16777215, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setText("")
        self.label_3.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout.addWidget(self.label_3)
        self.label_13 = QtWidgets.QLabel(self.widget_cabecalho)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_13.sizePolicy().hasHeightForWidth())
        self.label_13.setSizePolicy(sizePolicy)
        self.label_13.setMinimumSize(QtCore.QSize(300, 0))
        font = QtGui.QFont()
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label_13.setFont(font)
        self.label_13.setAlignment(QtCore.Qt.AlignCenter)
        self.label_13.setObjectName("label_13")
        self.horizontalLayout.addWidget(self.label_13)
        self.label_11 = QtWidgets.QLabel(self.widget_cabecalho)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_11.sizePolicy().hasHeightForWidth())
        self.label_11.setSizePolicy(sizePolicy)
        self.label_11.setMinimumSize(QtCore.QSize(0, 0))
        self.label_11.setMaximumSize(QtCore.QSize(16777215, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.label_11.setFont(font)
        self.label_11.setText("")
        self.label_11.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_11.setObjectName("label_11")
        self.horizontalLayout.addWidget(self.label_11)
        self.verticalLayout.addWidget(self.widget_cabecalho)
        self.widget_2 = QtWidgets.QWidget(self.centralwidget)
        self.widget_2.setMinimumSize(QtCore.QSize(0, 80))
        self.widget_2.setMaximumSize(QtCore.QSize(16777215, 80))
        self.widget_2.setObjectName("widget_2")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.widget_2)
        self.horizontalLayout_3.setContentsMargins(0, 0, 20, 0)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.widget_Cor4 = QtWidgets.QWidget(self.widget_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.widget_Cor4.sizePolicy().hasHeightForWidth())
        self.widget_Cor4.setSizePolicy(sizePolicy)
        self.widget_Cor4.setMinimumSize(QtCore.QSize(200, 0))
        self.widget_Cor4.setMaximumSize(QtCore.QSize(200, 16777215))
        self.widget_Cor4.setStyleSheet("background-color: rgb(182, 182, 182);")
        self.widget_Cor4.setObjectName("widget_Cor4")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.widget_Cor4)
        self.verticalLayout_4.setContentsMargins(1, 1, 1, 3)
        self.verticalLayout_4.setSpacing(1)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.widget_16 = QtWidgets.QWidget(self.widget_Cor4)
        self.widget_16.setMinimumSize(QtCore.QSize(0, 25))
        self.widget_16.setMaximumSize(QtCore.QSize(16777215, 25))
        self.widget_16.setStyleSheet("")
        self.widget_16.setObjectName("widget_16")
        self.horizontalLayout_17 = QtWidgets.QHBoxLayout(self.widget_16)
        self.horizontalLayout_17.setContentsMargins(3, 3, 3, 3)
        self.horizontalLayout_17.setSpacing(3)
        self.horizontalLayout_17.setObjectName("horizontalLayout_17")
        self.label_52 = QtWidgets.QLabel(self.widget_16)
        self.label_52.setMaximumSize(QtCore.QSize(16777215, 16777215))
        font = QtGui.QFont()
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label_52.setFont(font)
        self.label_52.setAlignment(QtCore.Qt.AlignCenter)
        self.label_52.setObjectName("label_52")
        self.horizontalLayout_17.addWidget(self.label_52)
        self.verticalLayout_4.addWidget(self.widget_16)
        self.widget_37 = QtWidgets.QWidget(self.widget_Cor4)
        self.widget_37.setMinimumSize(QtCore.QSize(0, 30))
        self.widget_37.setMaximumSize(QtCore.QSize(16777215, 30))
        self.widget_37.setStyleSheet("")
        self.widget_37.setObjectName("widget_37")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout(self.widget_37)
        self.horizontalLayout_4.setContentsMargins(5, 3, 5, 3)
        self.horizontalLayout_4.setSpacing(5)
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_54 = QtWidgets.QLabel(self.widget_37)
        self.label_54.setMinimumSize(QtCore.QSize(70, 0))
        self.label_54.setMaximumSize(QtCore.QSize(70, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.label_54.setFont(font)
        self.label_54.setAlignment(QtCore.Qt.AlignCenter)
        self.label_54.setObjectName("label_54")
        self.horizontalLayout_4.addWidget(self.label_54)
        self.line_Codigo_Manu = QtWidgets.QLineEdit(self.widget_37)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line_Codigo_Manu.sizePolicy().hasHeightForWidth())
        self.line_Codigo_Manu.setSizePolicy(sizePolicy)
        self.line_Codigo_Manu.setMinimumSize(QtCore.QSize(60, 0))
        self.line_Codigo_Manu.setMaximumSize(QtCore.QSize(80, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.line_Codigo_Manu.setFont(font)
        self.line_Codigo_Manu.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.line_Codigo_Manu.setAlignment(QtCore.Qt.AlignCenter)
        self.line_Codigo_Manu.setPlaceholderText("")
        self.line_Codigo_Manu.setObjectName("line_Codigo_Manu")
        self.horizontalLayout_4.addWidget(self.line_Codigo_Manu)
        self.verticalLayout_4.addWidget(self.widget_37)
        self.horizontalLayout_3.addWidget(self.widget_Cor4)
        self.widget_4 = QtWidgets.QWidget(self.widget_2)
        self.widget_4.setObjectName("widget_4")
        self.horizontalLayout_3.addWidget(self.widget_4)
        self.verticalLayout.addWidget(self.widget_2)
        self.widget_3 = QtWidgets.QWidget(self.centralwidget)
        self.widget_3.setObjectName("widget_3")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.widget_3)
        self.horizontalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.widget_Cor5 = QtWidgets.QWidget(self.widget_3)
        self.widget_Cor5.setStyleSheet("background-color: rgb(182, 182, 182);")
        self.widget_Cor5.setObjectName("widget_Cor5")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.widget_Cor5)
        self.verticalLayout_3.setContentsMargins(6, 6, 6, 6)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.label_Titulo = QtWidgets.QLabel(self.widget_Cor5)
        self.label_Titulo.setMaximumSize(QtCore.QSize(16777215, 18))
        font = QtGui.QFont()
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label_Titulo.setFont(font)
        self.label_Titulo.setAlignment(QtCore.Qt.AlignCenter)
        self.label_Titulo.setObjectName("label_Titulo")
        self.verticalLayout_3.addWidget(self.label_Titulo)
        self.table_Estrutura = QtWidgets.QTableWidget(self.widget_Cor5)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.table_Estrutura.sizePolicy().hasHeightForWidth())
        self.table_Estrutura.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(8)
        self.table_Estrutura.setFont(font)
        self.table_Estrutura.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.table_Estrutura.setObjectName("table_Estrutura")
        self.table_Estrutura.setColumnCount(9)
        self.table_Estrutura.setRowCount(0)
        item = QtWidgets.QTableWidgetItem()
        self.table_Estrutura.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Estrutura.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Estrutura.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Estrutura.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Estrutura.setHorizontalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Estrutura.setHorizontalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Estrutura.setHorizontalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Estrutura.setHorizontalHeaderItem(7, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Estrutura.setHorizontalHeaderItem(8, item)
        self.verticalLayout_3.addWidget(self.table_Estrutura)
        self.widget = QtWidgets.QWidget(self.widget_Cor5)
        self.widget.setMaximumSize(QtCore.QSize(16777215, 25))
        self.widget.setSizeIncrement(QtCore.QSize(0, 25))
        self.widget.setObjectName("widget")
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout(self.widget)
        self.horizontalLayout_10.setContentsMargins(6, 0, 6, 0)
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        self.btn_ExcluirTudo = QtWidgets.QPushButton(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.btn_ExcluirTudo.sizePolicy().hasHeightForWidth())
        self.btn_ExcluirTudo.setSizePolicy(sizePolicy)
        self.btn_ExcluirTudo.setMinimumSize(QtCore.QSize(80, 0))
        self.btn_ExcluirTudo.setMaximumSize(QtCore.QSize(80, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.btn_ExcluirTudo.setFont(font)
        self.btn_ExcluirTudo.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btn_ExcluirTudo.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.btn_ExcluirTudo.setAutoFillBackground(False)
        self.btn_ExcluirTudo.setStyleSheet("background-color: rgb(141, 141, 141);")
        self.btn_ExcluirTudo.setCheckable(False)
        self.btn_ExcluirTudo.setAutoDefault(False)
        self.btn_ExcluirTudo.setDefault(False)
        self.btn_ExcluirTudo.setFlat(False)
        self.btn_ExcluirTudo.setObjectName("btn_ExcluirTudo")
        self.horizontalLayout_10.addWidget(self.btn_ExcluirTudo)
        self.btn_ExcluirItem = QtWidgets.QPushButton(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.btn_ExcluirItem.sizePolicy().hasHeightForWidth())
        self.btn_ExcluirItem.setSizePolicy(sizePolicy)
        self.btn_ExcluirItem.setMinimumSize(QtCore.QSize(80, 0))
        self.btn_ExcluirItem.setMaximumSize(QtCore.QSize(80, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.btn_ExcluirItem.setFont(font)
        self.btn_ExcluirItem.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btn_ExcluirItem.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.btn_ExcluirItem.setAutoFillBackground(False)
        self.btn_ExcluirItem.setStyleSheet("background-color: rgb(141, 141, 141);")
        self.btn_ExcluirItem.setCheckable(False)
        self.btn_ExcluirItem.setAutoDefault(False)
        self.btn_ExcluirItem.setDefault(False)
        self.btn_ExcluirItem.setFlat(False)
        self.btn_ExcluirItem.setObjectName("btn_ExcluirItem")
        self.horizontalLayout_10.addWidget(self.btn_ExcluirItem)
        self.label_Texto = QtWidgets.QLabel(self.widget)
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.label_Texto.setFont(font)
        self.label_Texto.setText("")
        self.label_Texto.setObjectName("label_Texto")
        self.horizontalLayout_10.addWidget(self.label_Texto)
        self.widget_Progress = QtWidgets.QWidget(self.widget)
        self.widget_Progress.setMinimumSize(QtCore.QSize(0, 30))
        self.widget_Progress.setMaximumSize(QtCore.QSize(16777215, 30))
        self.widget_Progress.setStyleSheet("background-color: rgb(182, 182, 182);")
        self.widget_Progress.setObjectName("widget_Progress")
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout(self.widget_Progress)
        self.horizontalLayout_7.setContentsMargins(3, 3, 3, 3)
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.progressBar = QtWidgets.QProgressBar(self.widget_Progress)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.progressBar.sizePolicy().hasHeightForWidth())
        self.progressBar.setSizePolicy(sizePolicy)
        self.progressBar.setMaximum(0)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setTextVisible(False)
        self.progressBar.setObjectName("progressBar")
        self.horizontalLayout_7.addWidget(self.progressBar)
        self.horizontalLayout_10.addWidget(self.widget_Progress)
        self.btn_PDF = QtWidgets.QPushButton(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.btn_PDF.sizePolicy().hasHeightForWidth())
        self.btn_PDF.setSizePolicy(sizePolicy)
        self.btn_PDF.setMaximumSize(QtCore.QSize(80, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.btn_PDF.setFont(font)
        self.btn_PDF.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btn_PDF.setStyleSheet("background-color: rgb(141, 141, 141);")
        self.btn_PDF.setObjectName("btn_PDF")
        self.horizontalLayout_10.addWidget(self.btn_PDF)
        self.verticalLayout_3.addWidget(self.widget)
        self.horizontalLayout_2.addWidget(self.widget_Cor5)
        self.verticalLayout.addWidget(self.widget_3)
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Imprimir Ficha Produto"))
        self.label_13.setText(_translate("MainWindow", "Impressão de Fichas de Produtos"))
        self.label_52.setText(_translate("MainWindow", "Lançamento Manual"))
        self.label_54.setText(_translate("MainWindow", "Código:"))
        self.label_Titulo.setText(_translate("MainWindow", "Lista de Produtos"))
        item = self.table_Estrutura.horizontalHeaderItem(0)
        item.setText(_translate("MainWindow", "CÓDIGO"))
        item = self.table_Estrutura.horizontalHeaderItem(1)
        item.setText(_translate("MainWindow", "DESCRIÇÃO"))
        item = self.table_Estrutura.horizontalHeaderItem(2)
        item.setText(_translate("MainWindow", "DESCRIÇÃO COMPLEMENTAR"))
        item = self.table_Estrutura.horizontalHeaderItem(3)
        item.setText(_translate("MainWindow", "REFERÊNCIA"))
        item = self.table_Estrutura.horizontalHeaderItem(4)
        item.setText(_translate("MainWindow", "UM"))
        item = self.table_Estrutura.horizontalHeaderItem(5)
        item.setText(_translate("MainWindow", "NCM"))
        item = self.table_Estrutura.horizontalHeaderItem(6)
        item.setText(_translate("MainWindow", "CONJUNTO"))
        item = self.table_Estrutura.horizontalHeaderItem(7)
        item.setText(_translate("MainWindow", "SALDO"))
        item = self.table_Estrutura.horizontalHeaderItem(8)
        item.setText(_translate("MainWindow", "LOCAL"))
        self.btn_ExcluirTudo.setText(_translate("MainWindow", "Excluir Tudo"))
        self.btn_ExcluirItem.setText(_translate("MainWindow", "Excluir Item"))
        self.btn_PDF.setText(_translate("MainWindow", "Gerar PDF"))