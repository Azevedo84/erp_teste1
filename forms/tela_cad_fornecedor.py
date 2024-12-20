# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'tela_cad_fornecedor.ui'
#
# Created by: PyQt5 UI code generator 5.15.10
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1000, 600)
        MainWindow.setStyleSheet("")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setContentsMargins(6, 6, 6, 6)
        self.verticalLayout.setSpacing(6)
        self.verticalLayout.setObjectName("verticalLayout")
        self.widget = QtWidgets.QWidget(self.centralwidget)
        self.widget.setMinimumSize(QtCore.QSize(0, 40))
        self.widget.setMaximumSize(QtCore.QSize(16777215, 40))
        self.widget.setStyleSheet("background-color: rgb(0, 170, 255);")
        self.widget.setObjectName("widget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.widget)
        self.horizontalLayout.setContentsMargins(10, 5, 10, 5)
        self.horizontalLayout.setSpacing(2)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label_4 = QtWidgets.QLabel(self.widget)
        self.label_4.setMinimumSize(QtCore.QSize(48, 0))
        self.label_4.setMaximumSize(QtCore.QSize(48, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.label_4.setFont(font)
        self.label_4.setStyleSheet("")
        self.label_4.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout.addWidget(self.label_4)
        self.line_Num = QtWidgets.QLineEdit(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line_Num.sizePolicy().hasHeightForWidth())
        self.line_Num.setSizePolicy(sizePolicy)
        self.line_Num.setMinimumSize(QtCore.QSize(92, 0))
        self.line_Num.setMaximumSize(QtCore.QSize(92, 16777215))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.line_Num.setFont(font)
        self.line_Num.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.line_Num.setInputMask("")
        self.line_Num.setText("")
        self.line_Num.setFrame(True)
        self.line_Num.setEchoMode(QtWidgets.QLineEdit.Normal)
        self.line_Num.setAlignment(QtCore.Qt.AlignCenter)
        self.line_Num.setDragEnabled(False)
        self.line_Num.setPlaceholderText("")
        self.line_Num.setObjectName("line_Num")
        self.horizontalLayout.addWidget(self.line_Num)
        self.label_13 = QtWidgets.QLabel(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_13.sizePolicy().hasHeightForWidth())
        self.label_13.setSizePolicy(sizePolicy)
        self.label_13.setMinimumSize(QtCore.QSize(0, 0))
        font = QtGui.QFont()
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label_13.setFont(font)
        self.label_13.setStyleSheet("")
        self.label_13.setAlignment(QtCore.Qt.AlignCenter)
        self.label_13.setObjectName("label_13")
        self.horizontalLayout.addWidget(self.label_13)
        self.label_11 = QtWidgets.QLabel(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_11.sizePolicy().hasHeightForWidth())
        self.label_11.setSizePolicy(sizePolicy)
        self.label_11.setMinimumSize(QtCore.QSize(100, 0))
        self.label_11.setMaximumSize(QtCore.QSize(100, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.label_11.setFont(font)
        self.label_11.setStyleSheet("")
        self.label_11.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_11.setObjectName("label_11")
        self.horizontalLayout.addWidget(self.label_11)
        self.date_Emissao = QtWidgets.QDateEdit(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.date_Emissao.sizePolicy().hasHeightForWidth())
        self.date_Emissao.setSizePolicy(sizePolicy)
        self.date_Emissao.setMinimumSize(QtCore.QSize(101, 0))
        self.date_Emissao.setMaximumSize(QtCore.QSize(101, 16777215))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.date_Emissao.setFont(font)
        self.date_Emissao.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.date_Emissao.setAlignment(QtCore.Qt.AlignCenter)
        self.date_Emissao.setObjectName("date_Emissao")
        self.horizontalLayout.addWidget(self.date_Emissao)
        self.verticalLayout.addWidget(self.widget)
        self.widget_2 = QtWidgets.QWidget(self.centralwidget)
        self.widget_2.setStyleSheet("")
        self.widget_2.setObjectName("widget_2")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.widget_2)
        self.horizontalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_3.setSpacing(5)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.widget_4 = QtWidgets.QWidget(self.widget_2)
        self.widget_4.setMinimumSize(QtCore.QSize(391, 0))
        self.widget_4.setMaximumSize(QtCore.QSize(391, 16777215))
        self.widget_4.setStyleSheet("")
        self.widget_4.setObjectName("widget_4")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.widget_4)
        self.verticalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_2.setSpacing(5)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.widget_6 = QtWidgets.QWidget(self.widget_4)
        self.widget_6.setMinimumSize(QtCore.QSize(0, 110))
        self.widget_6.setMaximumSize(QtCore.QSize(16777215, 110))
        self.widget_6.setStyleSheet("background-color: rgb(182, 182, 182);")
        self.widget_6.setObjectName("widget_6")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.widget_6)
        self.verticalLayout_3.setContentsMargins(6, 6, 6, 6)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.label_2 = QtWidgets.QLabel(self.widget_6)
        self.label_2.setMinimumSize(QtCore.QSize(0, 18))
        self.label_2.setMaximumSize(QtCore.QSize(16777215, 18))
        font = QtGui.QFont()
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("label_2")
        self.verticalLayout_3.addWidget(self.label_2)
        self.widget_10 = QtWidgets.QWidget(self.widget_6)
        self.widget_10.setObjectName("widget_10")
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout(self.widget_10)
        self.horizontalLayout_6.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        self.label_25 = QtWidgets.QLabel(self.widget_10)
        self.label_25.setMinimumSize(QtCore.QSize(80, 0))
        self.label_25.setMaximumSize(QtCore.QSize(80, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.label_25.setFont(font)
        self.label_25.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_25.setObjectName("label_25")
        self.horizontalLayout_6.addWidget(self.label_25)
        self.line_Registro = QtWidgets.QLineEdit(self.widget_10)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line_Registro.sizePolicy().hasHeightForWidth())
        self.line_Registro.setSizePolicy(sizePolicy)
        self.line_Registro.setMinimumSize(QtCore.QSize(0, 25))
        self.line_Registro.setMaximumSize(QtCore.QSize(16777215, 25))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.line_Registro.setFont(font)
        self.line_Registro.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.line_Registro.setText("")
        self.line_Registro.setMaxLength(45)
        self.line_Registro.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.line_Registro.setObjectName("line_Registro")
        self.horizontalLayout_6.addWidget(self.line_Registro)
        self.verticalLayout_3.addWidget(self.widget_10)
        self.widget_8 = QtWidgets.QWidget(self.widget_6)
        self.widget_8.setObjectName("widget_8")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout(self.widget_8)
        self.horizontalLayout_4.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_24 = QtWidgets.QLabel(self.widget_8)
        self.label_24.setMinimumSize(QtCore.QSize(80, 0))
        self.label_24.setMaximumSize(QtCore.QSize(80, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.label_24.setFont(font)
        self.label_24.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_24.setObjectName("label_24")
        self.horizontalLayout_4.addWidget(self.label_24)
        self.line_Descricao = QtWidgets.QLineEdit(self.widget_8)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line_Descricao.sizePolicy().hasHeightForWidth())
        self.line_Descricao.setSizePolicy(sizePolicy)
        self.line_Descricao.setMinimumSize(QtCore.QSize(0, 25))
        self.line_Descricao.setMaximumSize(QtCore.QSize(16777215, 25))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.line_Descricao.setFont(font)
        self.line_Descricao.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.line_Descricao.setText("")
        self.line_Descricao.setMaxLength(45)
        self.line_Descricao.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.line_Descricao.setObjectName("line_Descricao")
        self.horizontalLayout_4.addWidget(self.line_Descricao)
        self.verticalLayout_3.addWidget(self.widget_8)
        self.verticalLayout_2.addWidget(self.widget_6)
        self.widget_7 = QtWidgets.QWidget(self.widget_4)
        self.widget_7.setObjectName("widget_7")
        self.verticalLayout_2.addWidget(self.widget_7)
        self.widget_3 = QtWidgets.QWidget(self.widget_4)
        self.widget_3.setMinimumSize(QtCore.QSize(0, 50))
        self.widget_3.setMaximumSize(QtCore.QSize(16777215, 50))
        self.widget_3.setStyleSheet("background-color: rgb(182, 182, 182);")
        self.widget_3.setObjectName("widget_3")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.widget_3)
        self.horizontalLayout_2.setContentsMargins(6, 6, 6, 6)
        self.horizontalLayout_2.setSpacing(6)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.btn_Excluir = QtWidgets.QPushButton(self.widget_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.btn_Excluir.sizePolicy().hasHeightForWidth())
        self.btn_Excluir.setSizePolicy(sizePolicy)
        self.btn_Excluir.setMaximumSize(QtCore.QSize(80, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.btn_Excluir.setFont(font)
        self.btn_Excluir.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btn_Excluir.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.btn_Excluir.setAutoFillBackground(False)
        self.btn_Excluir.setStyleSheet("background-color: rgb(141, 141, 141);")
        self.btn_Excluir.setCheckable(False)
        self.btn_Excluir.setAutoDefault(False)
        self.btn_Excluir.setDefault(False)
        self.btn_Excluir.setFlat(False)
        self.btn_Excluir.setObjectName("btn_Excluir")
        self.horizontalLayout_2.addWidget(self.btn_Excluir)
        self.label = QtWidgets.QLabel(self.widget_3)
        self.label.setText("")
        self.label.setObjectName("label")
        self.horizontalLayout_2.addWidget(self.label)
        self.btn_Limpar = QtWidgets.QPushButton(self.widget_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.btn_Limpar.sizePolicy().hasHeightForWidth())
        self.btn_Limpar.setSizePolicy(sizePolicy)
        self.btn_Limpar.setMaximumSize(QtCore.QSize(80, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.btn_Limpar.setFont(font)
        self.btn_Limpar.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btn_Limpar.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.btn_Limpar.setAutoFillBackground(False)
        self.btn_Limpar.setStyleSheet("background-color: rgb(141, 141, 141);")
        self.btn_Limpar.setCheckable(False)
        self.btn_Limpar.setAutoDefault(False)
        self.btn_Limpar.setDefault(False)
        self.btn_Limpar.setFlat(False)
        self.btn_Limpar.setObjectName("btn_Limpar")
        self.horizontalLayout_2.addWidget(self.btn_Limpar)
        self.btn_Salvar = QtWidgets.QPushButton(self.widget_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.btn_Salvar.sizePolicy().hasHeightForWidth())
        self.btn_Salvar.setSizePolicy(sizePolicy)
        self.btn_Salvar.setMaximumSize(QtCore.QSize(80, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.btn_Salvar.setFont(font)
        self.btn_Salvar.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btn_Salvar.setStyleSheet("background-color: rgb(141, 141, 141);")
        self.btn_Salvar.setObjectName("btn_Salvar")
        self.horizontalLayout_2.addWidget(self.btn_Salvar)
        self.verticalLayout_2.addWidget(self.widget_3)
        self.horizontalLayout_3.addWidget(self.widget_4)
        self.widget_5 = QtWidgets.QWidget(self.widget_2)
        self.widget_5.setStyleSheet("")
        self.widget_5.setObjectName("widget_5")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout(self.widget_5)
        self.verticalLayout_5.setContentsMargins(0, 0, 6, 0)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.widget_12 = QtWidgets.QWidget(self.widget_5)
        self.widget_12.setMaximumSize(QtCore.QSize(16777215, 45))
        self.widget_12.setSizeIncrement(QtCore.QSize(0, 45))
        self.widget_12.setStyleSheet("background-color: rgb(182, 182, 182);")
        self.widget_12.setObjectName("widget_12")
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout(self.widget_12)
        self.horizontalLayout_7.setContentsMargins(6, 6, 6, 6)
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.label_6 = QtWidgets.QLabel(self.widget_12)
        self.label_6.setMinimumSize(QtCore.QSize(120, 0))
        self.label_6.setMaximumSize(QtCore.QSize(120, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.label_6.setFont(font)
        self.label_6.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_6.setObjectName("label_6")
        self.horizontalLayout_7.addWidget(self.label_6)
        self.line_Consulta = QtWidgets.QLineEdit(self.widget_12)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line_Consulta.sizePolicy().hasHeightForWidth())
        self.line_Consulta.setSizePolicy(sizePolicy)
        self.line_Consulta.setMinimumSize(QtCore.QSize(0, 25))
        self.line_Consulta.setMaximumSize(QtCore.QSize(16777215, 25))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.line_Consulta.setFont(font)
        self.line_Consulta.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.line_Consulta.setInputMask("")
        self.line_Consulta.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.line_Consulta.setObjectName("line_Consulta")
        self.horizontalLayout_7.addWidget(self.line_Consulta)
        self.btn_Consulta = QtWidgets.QPushButton(self.widget_12)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.btn_Consulta.sizePolicy().hasHeightForWidth())
        self.btn_Consulta.setSizePolicy(sizePolicy)
        self.btn_Consulta.setMinimumSize(QtCore.QSize(80, 0))
        self.btn_Consulta.setMaximumSize(QtCore.QSize(80, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.btn_Consulta.setFont(font)
        self.btn_Consulta.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btn_Consulta.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.btn_Consulta.setAutoFillBackground(False)
        self.btn_Consulta.setStyleSheet("background-color: rgb(141, 141, 141);")
        self.btn_Consulta.setCheckable(False)
        self.btn_Consulta.setAutoDefault(False)
        self.btn_Consulta.setDefault(False)
        self.btn_Consulta.setFlat(False)
        self.btn_Consulta.setObjectName("btn_Consulta")
        self.horizontalLayout_7.addWidget(self.btn_Consulta)
        self.verticalLayout_5.addWidget(self.widget_12)
        self.widget_11 = QtWidgets.QWidget(self.widget_5)
        self.widget_11.setStyleSheet("background-color: rgb(182, 182, 182);")
        self.widget_11.setObjectName("widget_11")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.widget_11)
        self.verticalLayout_4.setContentsMargins(6, 6, 6, 6)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.label_5 = QtWidgets.QLabel(self.widget_11)
        self.label_5.setMinimumSize(QtCore.QSize(0, 18))
        self.label_5.setMaximumSize(QtCore.QSize(16777215, 18))
        font = QtGui.QFont()
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label_5.setFont(font)
        self.label_5.setAlignment(QtCore.Qt.AlignCenter)
        self.label_5.setObjectName("label_5")
        self.verticalLayout_4.addWidget(self.label_5)
        self.table_Lista = QtWidgets.QTableWidget(self.widget_11)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.table_Lista.sizePolicy().hasHeightForWidth())
        self.table_Lista.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.table_Lista.setFont(font)
        self.table_Lista.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.table_Lista.setObjectName("table_Lista")
        self.table_Lista.setColumnCount(4)
        self.table_Lista.setRowCount(0)
        item = QtWidgets.QTableWidgetItem()
        self.table_Lista.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Lista.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Lista.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Lista.setHorizontalHeaderItem(3, item)
        self.table_Lista.verticalHeader().setDefaultSectionSize(0)
        self.table_Lista.verticalHeader().setMinimumSectionSize(0)
        self.verticalLayout_4.addWidget(self.table_Lista)
        self.verticalLayout_5.addWidget(self.widget_11)
        self.horizontalLayout_3.addWidget(self.widget_5)
        self.verticalLayout.addWidget(self.widget_2)
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Cadastro de Fornecedores"))
        self.label_4.setText(_translate("MainWindow", "Código:"))
        self.label_13.setText(_translate("MainWindow", "Cadastro de Fornecedores"))
        self.label_11.setText(_translate("MainWindow", "Data Emissao: "))
        self.label_2.setText(_translate("MainWindow", "Dados Cadastro"))
        self.label_25.setText(_translate("MainWindow", "Registro:"))
        self.label_24.setText(_translate("MainWindow", "Descrição:"))
        self.btn_Excluir.setText(_translate("MainWindow", "Excluir"))
        self.btn_Limpar.setText(_translate("MainWindow", "Limpar"))
        self.btn_Salvar.setText(_translate("MainWindow", "Salvar"))
        self.label_6.setText(_translate("MainWindow", "Consulta Descrição:"))
        self.btn_Consulta.setText(_translate("MainWindow", "Consultar"))
        self.label_5.setText(_translate("MainWindow", "Lista de Fornecedores"))
        item = self.table_Lista.horizontalHeaderItem(0)
        item.setText(_translate("MainWindow", "CÓDIGO"))
        item = self.table_Lista.horizontalHeaderItem(1)
        item.setText(_translate("MainWindow", "CRIAÇÃO"))
        item = self.table_Lista.horizontalHeaderItem(2)
        item.setText(_translate("MainWindow", "REGISTRO"))
        item = self.table_Lista.horizontalHeaderItem(3)
        item.setText(_translate("MainWindow", "DESCRIÇÃO"))
