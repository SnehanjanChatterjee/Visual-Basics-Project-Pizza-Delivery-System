VERSION 5.00
Object = "{3A8BD65E-9922-4162-A649-83F2D5326BBE}#1.0#0"; "FoxitReaderBrowserAx.dll"
Begin VB.Form About_Us 
   Caption         =   "ABOUT US"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   13845
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin FOXITREADERLibCtl.FoxitCtl FoxitCtl1 
      Height          =   10935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20295
      _cx             =   5080
      _cy             =   5080
      src             =   ""
   End
End
Attribute VB_Name = "About_Us"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FoxitCtl1.OpenFile ("D:\College\College(4th Sem)\VB Project(Pizza Ordering System)\PROJECT-PIZZA DELIVERY STSTEM\About us.pdf")
End Sub
