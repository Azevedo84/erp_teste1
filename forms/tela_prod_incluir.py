# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'tela_prod_incluir.ui'
#
# Created by: PyQt5 UI code generator 5.15.11
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(683, 450)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout_2.setContentsMargins(6, 6, 6, 6)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
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
        self.label_13 = QtWidgets.QLabel(self.widget_cabecalho)
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
        self.label_13.setAlignment(QtCore.Qt.AlignCenter)
        self.label_13.setObjectName("label_13")
        self.horizontalLayout.addWidget(self.label_13)
        self.verticalLayout_2.addWidget(self.widget_cabecalho)
        self.widget_3 = QtWidgets.QWidget(self.centralwidget)
        self.widget_3.setObjectName("widget_3")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.widget_3)
        self.horizontalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.widget_5 = QtWidgets.QWidget(self.widget_3)
        self.widget_5.setMinimumSize(QtCore.QSize(0, 0))
        self.widget_5.setMaximumSize(QtCore.QSize(400, 16777215))
        self.widget_5.setObjectName("widget_5")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout(self.widget_5)
        self.verticalLayout_5.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.widget_Cor3 = QtWidgets.QWidget(self.widget_5)
        self.widget_Cor3.setMinimumSize(QtCore.QSize(0, 0))
        self.widget_Cor3.setMaximumSize(QtCore.QSize(16777215, 250))
        self.widget_Cor3.setStyleSheet("background-color: rgb(208, 208, 208);")
        self.widget_Cor3.setObjectName("widget_Cor3")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.widget_Cor3)
        self.verticalLayout.setContentsMargins(6, 6, 6, 6)
        self.verticalLayout.setObjectName("verticalLayout")
        self.label_14 = QtWidgets.QLabel(self.widget_Cor3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_14.sizePolicy().hasHeightForWidth())
        self.label_14.setSizePolicy(sizePolicy)
        self.label_14.setMinimumSize(QtCore.QSize(0, 0))
        self.label_14.setMaximumSize(QtCore.QSize(16777215, 18))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_14.setFont(font)
        self.label_14.setAlignment(QtCore.Qt.AlignCenter)
        self.label_14.setObjectName("label_14")
        self.verticalLayout.addWidget(self.label_14)
        self.widget_42 = QtWidgets.QWidget(self.widget_Cor3)
        self.widget_42.setMinimumSize(QtCore.QSize(0, 25))
        self.widget_42.setMaximumSize(QtCore.QSize(16777215, 25))
        self.widget_42.setStyleSheet("")
        self.widget_42.setObjectName("widget_42")
        self.horizontalLayout_34 = QtWidgets.QHBoxLayout(self.widget_42)
        self.horizontalLayout_34.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_34.setSpacing(6)
        self.horizontalLayout_34.setObjectName("horizontalLayout_34")
        self.label_3 = QtWidgets.QLabel(self.widget_42)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_3.sizePolicy().hasHeightForWidth())
        self.label_3.setSizePolicy(sizePolicy)
        self.label_3.setMinimumSize(QtCore.QSize(85, 0))
        self.label_3.setMaximumSize(QtCore.QSize(85, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_34.addWidget(self.label_3)
        self.line_Codigo = QtWidgets.QLineEdit(self.widget_42)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line_Codigo.sizePolicy().hasHeightForWidth())
        self.line_Codigo.setSizePolicy(sizePolicy)
        self.line_Codigo.setMinimumSize(QtCore.QSize(70, 0))
        self.line_Codigo.setMaximumSize(QtCore.QSize(70, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.line_Codigo.setFont(font)
        self.line_Codigo.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.line_Codigo.setInputMask("")
        self.line_Codigo.setText("")
        self.line_Codigo.setFrame(True)
        self.line_Codigo.setEchoMode(QtWidgets.QLineEdit.Normal)
        self.line_Codigo.setAlignment(QtCore.Qt.AlignCenter)
        self.line_Codigo.setDragEnabled(False)
        self.line_Codigo.setPlaceholderText("")
        self.line_Codigo.setObjectName("line_Codigo")
        self.horizontalLayout_34.addWidget(self.line_Codigo)
        self.label_17 = QtWidgets.QLabel(self.widget_42)
        self.label_17.setText("")
        self.label_17.setObjectName("label_17")
        self.horizontalLayout_34.addWidget(self.label_17)
        self.label_11 = QtWidgets.QLabel(self.widget_42)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_11.sizePolicy().hasHeightForWidth())
        self.label_11.setSizePolicy(sizePolicy)
        self.label_11.setMinimumSize(QtCore.QSize(100, 0))
        self.label_11.setMaximumSize(QtCore.QSize(100, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_11.setFont(font)
        self.label_11.setAlignment(QtCore.Qt.AlignCenter)
        self.label_11.setObjectName("label_11")
        self.horizontalLayout_34.addWidget(self.label_11)
        self.date_Emissao = QtWidgets.QDateEdit(self.widget_42)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.date_Emissao.sizePolicy().hasHeightForWidth())
        self.date_Emissao.setSizePolicy(sizePolicy)
        self.date_Emissao.setMinimumSize(QtCore.QSize(101, 0))
        self.date_Emissao.setMaximumSize(QtCore.QSize(101, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.date_Emissao.setFont(font)
        self.date_Emissao.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.date_Emissao.setAlignment(QtCore.Qt.AlignCenter)
        self.date_Emissao.setObjectName("date_Emissao")
        self.horizontalLayout_34.addWidget(self.date_Emissao)
        self.verticalLayout.addWidget(self.widget_42)
        self.widget_37 = QtWidgets.QWidget(self.widget_Cor3)
        self.widget_37.setMinimumSize(QtCore.QSize(0, 25))
        self.widget_37.setMaximumSize(QtCore.QSize(16777215, 25))
        self.widget_37.setStyleSheet("")
        self.widget_37.setObjectName("widget_37")
        self.horizontalLayout_30 = QtWidgets.QHBoxLayout(self.widget_37)
        self.horizontalLayout_30.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_30.setSpacing(6)
        self.horizontalLayout_30.setObjectName("horizontalLayout_30")
        self.label_59 = QtWidgets.QLabel(self.widget_37)
        self.label_59.setMinimumSize(QtCore.QSize(85, 0))
        self.label_59.setMaximumSize(QtCore.QSize(85, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_59.setFont(font)
        self.label_59.setAlignment(QtCore.Qt.AlignCenter)
        self.label_59.setObjectName("label_59")
        self.horizontalLayout_30.addWidget(self.label_59)
        self.line_Referencia = QtWidgets.QLineEdit(self.widget_37)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line_Referencia.sizePolicy().hasHeightForWidth())
        self.line_Referencia.setSizePolicy(sizePolicy)
        self.line_Referencia.setMinimumSize(QtCore.QSize(0, 0))
        self.line_Referencia.setMaximumSize(QtCore.QSize(160, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.line_Referencia.setFont(font)
        self.line_Referencia.setCursor(QtGui.QCursor(QtCore.Qt.IBeamCursor))
        self.line_Referencia.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.line_Referencia.setInputMask("")
        self.line_Referencia.setText("")
        self.line_Referencia.setMaxLength(20)
        self.line_Referencia.setFrame(True)
        self.line_Referencia.setEchoMode(QtWidgets.QLineEdit.Normal)
        self.line_Referencia.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.line_Referencia.setObjectName("line_Referencia")
        self.horizontalLayout_30.addWidget(self.line_Referencia)
        self.label_16 = QtWidgets.QLabel(self.widget_37)
        self.label_16.setText("")
        self.label_16.setObjectName("label_16")
        self.horizontalLayout_30.addWidget(self.label_16)
        self.label_26 = QtWidgets.QLabel(self.widget_37)
        self.label_26.setMaximumSize(QtCore.QSize(25, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_26.setFont(font)
        self.label_26.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_26.setObjectName("label_26")
        self.horizontalLayout_30.addWidget(self.label_26)
        self.combo_UM = QtWidgets.QComboBox(self.widget_37)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.combo_UM.sizePolicy().hasHeightForWidth())
        self.combo_UM.setSizePolicy(sizePolicy)
        self.combo_UM.setMinimumSize(QtCore.QSize(50, 0))
        self.combo_UM.setMaximumSize(QtCore.QSize(50, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.combo_UM.setFont(font)
        self.combo_UM.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"text-align: center;")
        self.combo_UM.setObjectName("combo_UM")
        self.combo_UM.addItem("")
        self.combo_UM.setItemText(0, "")
        self.combo_UM.addItem("")
        self.combo_UM.addItem("")
        self.combo_UM.addItem("")
        self.combo_UM.addItem("")
        self.combo_UM.addItem("")
        self.combo_UM.addItem("")
        self.combo_UM.addItem("")
        self.combo_UM.addItem("")
        self.combo_UM.addItem("")
        self.horizontalLayout_30.addWidget(self.combo_UM)
        self.verticalLayout.addWidget(self.widget_37)
        self.widget_38 = QtWidgets.QWidget(self.widget_Cor3)
        self.widget_38.setMinimumSize(QtCore.QSize(0, 25))
        self.widget_38.setMaximumSize(QtCore.QSize(16777215, 25))
        self.widget_38.setStyleSheet("")
        self.widget_38.setObjectName("widget_38")
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout(self.widget_38)
        self.horizontalLayout_5.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_5.setSpacing(6)
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.label_50 = QtWidgets.QLabel(self.widget_38)
        self.label_50.setMinimumSize(QtCore.QSize(85, 0))
        self.label_50.setMaximumSize(QtCore.QSize(85, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_50.setFont(font)
        self.label_50.setAlignment(QtCore.Qt.AlignCenter)
        self.label_50.setObjectName("label_50")
        self.horizontalLayout_5.addWidget(self.label_50)
        self.line_Descricao = QtWidgets.QLineEdit(self.widget_38)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line_Descricao.sizePolicy().hasHeightForWidth())
        self.line_Descricao.setSizePolicy(sizePolicy)
        self.line_Descricao.setMinimumSize(QtCore.QSize(0, 0))
        self.line_Descricao.setMaximumSize(QtCore.QSize(16777215, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.line_Descricao.setFont(font)
        self.line_Descricao.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.line_Descricao.setText("")
        self.line_Descricao.setMaxLength(30)
        self.line_Descricao.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.line_Descricao.setObjectName("line_Descricao")
        self.horizontalLayout_5.addWidget(self.line_Descricao)
        self.verticalLayout.addWidget(self.widget_38)
        self.widget_33 = QtWidgets.QWidget(self.widget_Cor3)
        self.widget_33.setMinimumSize(QtCore.QSize(0, 25))
        self.widget_33.setMaximumSize(QtCore.QSize(16777215, 25))
        self.widget_33.setStyleSheet("")
        self.widget_33.setObjectName("widget_33")
        self.horizontalLayout_26 = QtWidgets.QHBoxLayout(self.widget_33)
        self.horizontalLayout_26.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_26.setSpacing(6)
        self.horizontalLayout_26.setObjectName("horizontalLayout_26")
        self.label_64 = QtWidgets.QLabel(self.widget_33)
        self.label_64.setMinimumSize(QtCore.QSize(85, 0))
        self.label_64.setMaximumSize(QtCore.QSize(85, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_64.setFont(font)
        self.label_64.setAlignment(QtCore.Qt.AlignCenter)
        self.label_64.setObjectName("label_64")
        self.horizontalLayout_26.addWidget(self.label_64)
        self.line_DescrCompl = QtWidgets.QLineEdit(self.widget_33)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line_DescrCompl.sizePolicy().hasHeightForWidth())
        self.line_DescrCompl.setSizePolicy(sizePolicy)
        self.line_DescrCompl.setMinimumSize(QtCore.QSize(0, 0))
        self.line_DescrCompl.setMaximumSize(QtCore.QSize(16777215, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.line_DescrCompl.setFont(font)
        self.line_DescrCompl.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.line_DescrCompl.setText("")
        self.line_DescrCompl.setMaxLength(30)
        self.line_DescrCompl.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.line_DescrCompl.setObjectName("line_DescrCompl")
        self.horizontalLayout_26.addWidget(self.line_DescrCompl)
        self.verticalLayout.addWidget(self.widget_33)
        self.widget_34 = QtWidgets.QWidget(self.widget_Cor3)
        self.widget_34.setMinimumSize(QtCore.QSize(0, 25))
        self.widget_34.setMaximumSize(QtCore.QSize(16777215, 25))
        self.widget_34.setStyleSheet("")
        self.widget_34.setObjectName("widget_34")
        self.horizontalLayout_27 = QtWidgets.QHBoxLayout(self.widget_34)
        self.horizontalLayout_27.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_27.setSpacing(6)
        self.horizontalLayout_27.setObjectName("horizontalLayout_27")
        self.label_72 = QtWidgets.QLabel(self.widget_34)
        self.label_72.setMinimumSize(QtCore.QSize(85, 0))
        self.label_72.setMaximumSize(QtCore.QSize(85, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_72.setFont(font)
        self.label_72.setAlignment(QtCore.Qt.AlignCenter)
        self.label_72.setObjectName("label_72")
        self.horizontalLayout_27.addWidget(self.label_72)
        self.combo_Classificacao = QtWidgets.QComboBox(self.widget_34)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.combo_Classificacao.sizePolicy().hasHeightForWidth())
        self.combo_Classificacao.setSizePolicy(sizePolicy)
        self.combo_Classificacao.setMinimumSize(QtCore.QSize(220, 0))
        self.combo_Classificacao.setMaximumSize(QtCore.QSize(220, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.combo_Classificacao.setFont(font)
        self.combo_Classificacao.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"text-align: center;")
        self.combo_Classificacao.setObjectName("combo_Classificacao")
        self.horizontalLayout_27.addWidget(self.combo_Classificacao)
        self.label_10 = QtWidgets.QLabel(self.widget_34)
        self.label_10.setText("")
        self.label_10.setObjectName("label_10")
        self.horizontalLayout_27.addWidget(self.label_10)
        self.verticalLayout.addWidget(self.widget_34)
        self.widget_35 = QtWidgets.QWidget(self.widget_Cor3)
        self.widget_35.setMinimumSize(QtCore.QSize(0, 25))
        self.widget_35.setMaximumSize(QtCore.QSize(16777215, 25))
        self.widget_35.setStyleSheet("")
        self.widget_35.setObjectName("widget_35")
        self.horizontalLayout_28 = QtWidgets.QHBoxLayout(self.widget_35)
        self.horizontalLayout_28.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_28.setSpacing(6)
        self.horizontalLayout_28.setObjectName("horizontalLayout_28")
        self.label_63 = QtWidgets.QLabel(self.widget_35)
        self.label_63.setMinimumSize(QtCore.QSize(85, 0))
        self.label_63.setMaximumSize(QtCore.QSize(85, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_63.setFont(font)
        self.label_63.setAlignment(QtCore.Qt.AlignCenter)
        self.label_63.setObjectName("label_63")
        self.horizontalLayout_28.addWidget(self.label_63)
        self.line_NCM = QtWidgets.QLineEdit(self.widget_35)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line_NCM.sizePolicy().hasHeightForWidth())
        self.line_NCM.setSizePolicy(sizePolicy)
        self.line_NCM.setMinimumSize(QtCore.QSize(0, 0))
        self.line_NCM.setMaximumSize(QtCore.QSize(150, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.line_NCM.setFont(font)
        self.line_NCM.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.line_NCM.setMaxLength(10)
        self.line_NCM.setAlignment(QtCore.Qt.AlignCenter)
        self.line_NCM.setObjectName("line_NCM")
        self.horizontalLayout_28.addWidget(self.line_NCM)
        self.label_7 = QtWidgets.QLabel(self.widget_35)
        self.label_7.setText("")
        self.label_7.setObjectName("label_7")
        self.horizontalLayout_28.addWidget(self.label_7)
        self.label_74 = QtWidgets.QLabel(self.widget_35)
        self.label_74.setMinimumSize(QtCore.QSize(85, 0))
        self.label_74.setMaximumSize(QtCore.QSize(85, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_74.setFont(font)
        self.label_74.setAlignment(QtCore.Qt.AlignCenter)
        self.label_74.setObjectName("label_74")
        self.horizontalLayout_28.addWidget(self.label_74)
        self.line_Qtde_Mini = QtWidgets.QLineEdit(self.widget_35)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line_Qtde_Mini.sizePolicy().hasHeightForWidth())
        self.line_Qtde_Mini.setSizePolicy(sizePolicy)
        self.line_Qtde_Mini.setMinimumSize(QtCore.QSize(80, 0))
        self.line_Qtde_Mini.setMaximumSize(QtCore.QSize(80, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.line_Qtde_Mini.setFont(font)
        self.line_Qtde_Mini.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.line_Qtde_Mini.setText("")
        self.line_Qtde_Mini.setMaxLength(30)
        self.line_Qtde_Mini.setAlignment(QtCore.Qt.AlignCenter)
        self.line_Qtde_Mini.setObjectName("line_Qtde_Mini")
        self.horizontalLayout_28.addWidget(self.line_Qtde_Mini)
        self.verticalLayout.addWidget(self.widget_35)
        self.widget_39 = QtWidgets.QWidget(self.widget_Cor3)
        self.widget_39.setMinimumSize(QtCore.QSize(0, 30))
        self.widget_39.setMaximumSize(QtCore.QSize(16777215, 30))
        self.widget_39.setStyleSheet("")
        self.widget_39.setObjectName("widget_39")
        self.horizontalLayout_31 = QtWidgets.QHBoxLayout(self.widget_39)
        self.horizontalLayout_31.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_31.setSpacing(6)
        self.horizontalLayout_31.setObjectName("horizontalLayout_31")
        self.label_80 = QtWidgets.QLabel(self.widget_39)
        self.label_80.setMinimumSize(QtCore.QSize(85, 0))
        self.label_80.setMaximumSize(QtCore.QSize(85, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_80.setFont(font)
        self.label_80.setAlignment(QtCore.Qt.AlignCenter)
        self.label_80.setObjectName("label_80")
        self.horizontalLayout_31.addWidget(self.label_80)
        self.plain_Obs = QtWidgets.QPlainTextEdit(self.widget_39)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.plain_Obs.sizePolicy().hasHeightForWidth())
        self.plain_Obs.setSizePolicy(sizePolicy)
        self.plain_Obs.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.plain_Obs.setObjectName("plain_Obs")
        self.horizontalLayout_31.addWidget(self.plain_Obs)
        self.verticalLayout.addWidget(self.widget_39)
        self.verticalLayout_5.addWidget(self.widget_Cor3)
        self.widget_Cor2 = QtWidgets.QWidget(self.widget_5)
        self.widget_Cor2.setMinimumSize(QtCore.QSize(0, 31))
        self.widget_Cor2.setMaximumSize(QtCore.QSize(16777215, 31))
        self.widget_Cor2.setStyleSheet("background-color: rgb(208, 208, 208);")
        self.widget_Cor2.setObjectName("widget_Cor2")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.widget_Cor2)
        self.horizontalLayout_3.setContentsMargins(0, 3, 0, 3)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.label_4 = QtWidgets.QLabel(self.widget_Cor2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_4.sizePolicy().hasHeightForWidth())
        self.label_4.setSizePolicy(sizePolicy)
        self.label_4.setMinimumSize(QtCore.QSize(0, 0))
        self.label_4.setMaximumSize(QtCore.QSize(50, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_4.setFont(font)
        self.label_4.setAlignment(QtCore.Qt.AlignCenter)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_3.addWidget(self.label_4)
        self.line_ID_Pre = QtWidgets.QLineEdit(self.widget_Cor2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line_ID_Pre.sizePolicy().hasHeightForWidth())
        self.line_ID_Pre.setSizePolicy(sizePolicy)
        self.line_ID_Pre.setMinimumSize(QtCore.QSize(0, 0))
        self.line_ID_Pre.setMaximumSize(QtCore.QSize(92, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.line_ID_Pre.setFont(font)
        self.line_ID_Pre.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.line_ID_Pre.setInputMask("")
        self.line_ID_Pre.setText("")
        self.line_ID_Pre.setFrame(True)
        self.line_ID_Pre.setEchoMode(QtWidgets.QLineEdit.Normal)
        self.line_ID_Pre.setAlignment(QtCore.Qt.AlignCenter)
        self.line_ID_Pre.setDragEnabled(False)
        self.line_ID_Pre.setPlaceholderText("")
        self.line_ID_Pre.setObjectName("line_ID_Pre")
        self.horizontalLayout_3.addWidget(self.line_ID_Pre)
        self.label_Fornecedor = QtWidgets.QLabel(self.widget_Cor2)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_Fornecedor.setFont(font)
        self.label_Fornecedor.setText("")
        self.label_Fornecedor.setAlignment(QtCore.Qt.AlignCenter)
        self.label_Fornecedor.setObjectName("label_Fornecedor")
        self.horizontalLayout_3.addWidget(self.label_Fornecedor)
        self.label_2 = QtWidgets.QLabel(self.widget_Cor2)
        self.label_2.setText("")
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_3.addWidget(self.label_2)
        self.label_Maquina_Des = QtWidgets.QLabel(self.widget_Cor2)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_Maquina_Des.setFont(font)
        self.label_Maquina_Des.setText("")
        self.label_Maquina_Des.setAlignment(QtCore.Qt.AlignCenter)
        self.label_Maquina_Des.setObjectName("label_Maquina_Des")
        self.horizontalLayout_3.addWidget(self.label_Maquina_Des)
        self.verticalLayout_5.addWidget(self.widget_Cor2)
        self.widget = QtWidgets.QWidget(self.widget_5)
        self.widget.setObjectName("widget")
        self.verticalLayout_5.addWidget(self.widget)
        self.widget_Cor4 = QtWidgets.QWidget(self.widget_5)
        self.widget_Cor4.setMinimumSize(QtCore.QSize(0, 40))
        self.widget_Cor4.setMaximumSize(QtCore.QSize(16777215, 40))
        self.widget_Cor4.setStyleSheet("background-color: rgb(208, 208, 208);")
        self.widget_Cor4.setObjectName("widget_Cor4")
        self.horizontalLayout_25 = QtWidgets.QHBoxLayout(self.widget_Cor4)
        self.horizontalLayout_25.setContentsMargins(6, 6, 6, 6)
        self.horizontalLayout_25.setSpacing(6)
        self.horizontalLayout_25.setObjectName("horizontalLayout_25")
        self.btn_Limpar = QtWidgets.QPushButton(self.widget_Cor4)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.btn_Limpar.sizePolicy().hasHeightForWidth())
        self.btn_Limpar.setSizePolicy(sizePolicy)
        self.btn_Limpar.setMinimumSize(QtCore.QSize(80, 0))
        self.btn_Limpar.setMaximumSize(QtCore.QSize(80, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.btn_Limpar.setFont(font)
        self.btn_Limpar.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btn_Limpar.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.btn_Limpar.setAutoFillBackground(False)
        self.btn_Limpar.setStyleSheet("background-color: rgb(170, 170, 170);")
        self.btn_Limpar.setCheckable(False)
        self.btn_Limpar.setAutoDefault(False)
        self.btn_Limpar.setDefault(False)
        self.btn_Limpar.setFlat(False)
        self.btn_Limpar.setObjectName("btn_Limpar")
        self.horizontalLayout_25.addWidget(self.btn_Limpar)
        self.label = QtWidgets.QLabel(self.widget_Cor4)
        self.label.setText("")
        self.label.setObjectName("label")
        self.horizontalLayout_25.addWidget(self.label)
        self.btn_Salvar = QtWidgets.QPushButton(self.widget_Cor4)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.btn_Salvar.sizePolicy().hasHeightForWidth())
        self.btn_Salvar.setSizePolicy(sizePolicy)
        self.btn_Salvar.setMinimumSize(QtCore.QSize(80, 0))
        self.btn_Salvar.setMaximumSize(QtCore.QSize(80, 16777215))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.btn_Salvar.setFont(font)
        self.btn_Salvar.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btn_Salvar.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.btn_Salvar.setAutoFillBackground(False)
        self.btn_Salvar.setStyleSheet("background-color: rgb(170, 170, 170);")
        self.btn_Salvar.setCheckable(False)
        self.btn_Salvar.setAutoDefault(False)
        self.btn_Salvar.setDefault(False)
        self.btn_Salvar.setFlat(False)
        self.btn_Salvar.setObjectName("btn_Salvar")
        self.horizontalLayout_25.addWidget(self.btn_Salvar)
        self.verticalLayout_5.addWidget(self.widget_Cor4)
        self.horizontalLayout_2.addWidget(self.widget_5)
        self.widget_Cor1 = QtWidgets.QWidget(self.widget_3)
        self.widget_Cor1.setStyleSheet("background-color: rgb(208, 208, 208);")
        self.widget_Cor1.setObjectName("widget_Cor1")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.widget_Cor1)
        self.verticalLayout_3.setContentsMargins(6, 6, 6, 6)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.label_Titulo = QtWidgets.QLabel(self.widget_Cor1)
        self.label_Titulo.setMaximumSize(QtCore.QSize(16777215, 18))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_Titulo.setFont(font)
        self.label_Titulo.setAlignment(QtCore.Qt.AlignCenter)
        self.label_Titulo.setObjectName("label_Titulo")
        self.verticalLayout_3.addWidget(self.label_Titulo)
        self.table_Produto = QtWidgets.QTableWidget(self.widget_Cor1)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.table_Produto.sizePolicy().hasHeightForWidth())
        self.table_Produto.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(8)
        self.table_Produto.setFont(font)
        self.table_Produto.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.table_Produto.setObjectName("table_Produto")
        self.table_Produto.setColumnCount(9)
        self.table_Produto.setRowCount(0)
        item = QtWidgets.QTableWidgetItem()
        self.table_Produto.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Produto.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Produto.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Produto.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Produto.setHorizontalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Produto.setHorizontalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Produto.setHorizontalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Produto.setHorizontalHeaderItem(7, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_Produto.setHorizontalHeaderItem(8, item)
        self.verticalLayout_3.addWidget(self.table_Produto)
        self.horizontalLayout_2.addWidget(self.widget_Cor1)
        self.verticalLayout_2.addWidget(self.widget_3)
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        MainWindow.setTabOrder(self.line_Descricao, self.line_DescrCompl)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Incluir Produto"))
        self.label_13.setText(_translate("MainWindow", "Cadastro de Produtos"))
        self.label_14.setText(_translate("MainWindow", "Dados do Produto"))
        self.label_3.setText(_translate("MainWindow", "Código Siger:"))
        self.label_11.setText(_translate("MainWindow", "Data Emissao: "))
        self.label_59.setText(_translate("MainWindow", "Referência:"))
        self.label_26.setText(_translate("MainWindow", "UM:"))
        self.combo_UM.setItemText(1, _translate("MainWindow", "UN"))
        self.combo_UM.setItemText(2, _translate("MainWindow", "MT"))
        self.combo_UM.setItemText(3, _translate("MainWindow", "MM"))
        self.combo_UM.setItemText(4, _translate("MainWindow", "LT"))
        self.combo_UM.setItemText(5, _translate("MainWindow", "KG"))
        self.combo_UM.setItemText(6, _translate("MainWindow", "CT"))
        self.combo_UM.setItemText(7, _translate("MainWindow", "ML"))
        self.combo_UM.setItemText(8, _translate("MainWindow", "M²"))
        self.combo_UM.setItemText(9, _translate("MainWindow", "M³"))
        self.label_50.setText(_translate("MainWindow", "Descrição:"))
        self.label_64.setText(_translate("MainWindow", "D. Compl.:"))
        self.label_72.setText(_translate("MainWindow", "Class. Contábil:"))
        self.label_63.setText(_translate("MainWindow", "NCM:"))
        self.label_74.setText(_translate("MainWindow", "Qtde Mínima:"))
        self.label_80.setText(_translate("MainWindow", "Observação:"))
        self.label_4.setText(_translate("MainWindow", "ID Pré"))
        self.btn_Limpar.setText(_translate("MainWindow", "Limpar"))
        self.btn_Salvar.setText(_translate("MainWindow", "Salvar"))
        self.label_Titulo.setText(_translate("MainWindow", "Lista Pré Cadastro"))
        item = self.table_Produto.horizontalHeaderItem(0)
        item.setText(_translate("MainWindow", "Data"))
        item = self.table_Produto.horizontalHeaderItem(1)
        item.setText(_translate("MainWindow", "Cód."))
        item = self.table_Produto.horizontalHeaderItem(2)
        item.setText(_translate("MainWindow", "Observação"))
        item = self.table_Produto.horizontalHeaderItem(3)
        item.setText(_translate("MainWindow", "Descrição"))
        item = self.table_Produto.horizontalHeaderItem(4)
        item.setText(_translate("MainWindow", "D. Compl."))
        item = self.table_Produto.horizontalHeaderItem(5)
        item.setText(_translate("MainWindow", "Referência"))
        item = self.table_Produto.horizontalHeaderItem(6)
        item.setText(_translate("MainWindow", "Um"))
        item = self.table_Produto.horizontalHeaderItem(7)
        item.setText(_translate("MainWindow", "NCM"))
        item = self.table_Produto.horizontalHeaderItem(8)
        item.setText(_translate("MainWindow", "Fornecedor"))