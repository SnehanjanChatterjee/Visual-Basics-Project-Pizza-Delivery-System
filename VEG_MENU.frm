VERSION 5.00
Begin VB.Form VEG_MENU 
   Caption         =   "VEG-MENU"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "VEG_MENU.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton BACK 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      TabIndex        =   14
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton HOME 
      Caption         =   "PIZZA   FACTORY"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD ITEM"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   7920
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD ITEM"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   7920
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ADD ITEM"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   5
      Top             =   7920
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   4035
      Left            =   960
      Picture         =   "VEG_MENU.frx":1CE64
      ScaleHeight     =   3975
      ScaleWidth      =   4005
      TabIndex        =   4
      Top             =   2880
      Width           =   4065
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   4035
      Left            =   5640
      Picture         =   "VEG_MENU.frx":3E3B3
      ScaleHeight     =   3975
      ScaleWidth      =   4005
      TabIndex        =   3
      Top             =   2880
      Width           =   4065
   End
   Begin VB.PictureBox Picture3 
      Height          =   4095
      Left            =   10560
      Picture         =   "VEG_MENU.frx":630EB
      ScaleHeight     =   4035
      ScaleWidth      =   3915
      TabIndex        =   2
      Top             =   2880
      Width           =   3975
   End
   Begin VB.PictureBox Picture4 
      Height          =   4095
      Left            =   15360
      Picture         =   "VEG_MENU.frx":842A6
      ScaleHeight     =   4035
      ScaleWidth      =   3915
      TabIndex        =   1
      Top             =   2880
      Width           =   3975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ADD ITEM"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15360
      TabIndex        =   0
      Top             =   7920
      Width           =   3855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PIZZA FACTORY"
      BeginProperty Font 
         Name            =   "Adobe Gothic Std B"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   8040
      TabIndex        =   12
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "DOUBLE CHEESE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   6960
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FARM HOUSE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   6960
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "MARGHERITA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10560
      TabIndex        =   9
      Top             =   6960
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "EXTRA VEGGIE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   15240
      TabIndex        =   8
      Top             =   6960
      Width           =   3975
   End
End
Attribute VB_Name = "VEG_MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public base As Double
Dim a, price As Integer
Public flag As Integer
Function value(a As Integer)
flag = 1
Load CUSTOMIZATION
CUSTOMIZATION.Show
Unload Me
End Function

Private Sub Form_Load()
flag = 0
End Sub

Private Sub HOME_Click()
Load Front_page
Front_page.Show
Unload Me
End Sub

Private Sub BACK_Click()
Load Ordering
Ordering.Show
Unload Me
End Sub

Private Sub Command1_Click()
value (1)
CUSTOMIZATION.Picture1.Picture = VEG_MENU.Picture1.Picture
price = 150
CUSTOMIZATION.base = price
CHECKOUT.List1.AddItem "Double Cheese"
Payment_Summary.List1.AddItem "Double Cheese"
End Sub

Private Sub Command2_Click()
value (1)
CUSTOMIZATION.Picture1.Picture = VEG_MENU.Picture2.Picture
price = 250
CUSTOMIZATION.base = price
CHECKOUT.List1.AddItem "Farm House"
Payment_Summary.List1.AddItem "Farm House"
End Sub

Private Sub Command3_Click()
value (1)
CUSTOMIZATION.Picture1.Picture = VEG_MENU.Picture3.Picture
price = 100
CUSTOMIZATION.base = price
CHECKOUT.List1.AddItem "Margherita"
Payment_Summary.List1.AddItem "Margherita"
End Sub

Private Sub Command4_Click()
flag = 1
value (1)
CUSTOMIZATION.Picture1.Picture = VEG_MENU.Picture4.Picture
price = 200
CUSTOMIZATION.base = price
CHECKOUT.List1.AddItem "Extra Veggie"
Payment_Summary.List1.AddItem "Extra Veggie"
End Sub
