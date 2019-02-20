VERSION 5.00
Begin VB.Form NONVEG_MENU 
   Caption         =   "NON VEG-MENU"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "NONVEG_MENU.frx":0000
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
      TabIndex        =   11
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton HOME 
      Caption         =   "PIZZA FACTORY"
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
      TabIndex        =   10
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   4080
      Picture         =   "NONVEG_MENU.frx":1CE64
      ScaleHeight     =   3915
      ScaleWidth      =   3915
      TabIndex        =   5
      Top             =   3120
      Width           =   3975
   End
   Begin VB.PictureBox Picture2 
      Height          =   3975
      Left            =   8520
      Picture         =   "NONVEG_MENU.frx":4209B
      ScaleHeight     =   3915
      ScaleWidth      =   3915
      TabIndex        =   4
      Top             =   3120
      Width           =   3975
   End
   Begin VB.PictureBox Picture3 
      Height          =   3975
      Left            =   13080
      Picture         =   "NONVEG_MENU.frx":65DB7
      ScaleHeight     =   3915
      ScaleWidth      =   3915
      TabIndex        =   3
      Top             =   3120
      Width           =   3975
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
      Left            =   4080
      TabIndex        =   2
      Top             =   7920
      Width           =   3975
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
      Left            =   8520
      TabIndex        =   1
      Top             =   7920
      Width           =   4095
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
      Left            =   13080
      TabIndex        =   0
      Top             =   7920
      Width           =   4095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "   PIZZA   FACTORY"
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
      Height          =   1935
      Left            =   7800
      TabIndex        =   9
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "GOLDEN DELIGHT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4080
      TabIndex        =   8
      Top             =   7320
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "DOMINATOR"
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
      Left            =   8520
      TabIndex        =   7
      Top             =   7320
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "CHICKEN TIKKA"
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
      Left            =   13080
      TabIndex        =   6
      Top             =   7320
      Width           =   4095
   End
End
Attribute VB_Name = "NONVEG_MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public price As Double
Dim a As Integer
Public flag1 As Integer

Private Sub Form_Load()
flag1 = 0
End Sub

Function value(a As Integer)
flag1 = 1
Load CUSTOMIZATION
CUSTOMIZATION.Show
Unload Me
End Function

Private Sub BACK_Click()
Load Ordering
Ordering.Show
Unload Me
End Sub

Private Sub Command1_Click()
value (1)
CUSTOMIZATION.Picture1.Picture = NONVEG_MENU.Picture1.Picture
price = 200
CUSTOMIZATION.base = price
CHECKOUT.List1.AddItem "Golden Delight"
Payment_Summary.List1.AddItem "Golden Delight"
End Sub

Private Sub Command2_Click()
value (1)
CUSTOMIZATION.Picture1.Picture = NONVEG_MENU.Picture2.Picture
price = 300
CUSTOMIZATION.base = price
CHECKOUT.List1.AddItem "Dominator"
Payment_Summary.List1.AddItem "Dominator"
End Sub

Private Sub Command3_Click()
value (1)
CUSTOMIZATION.Picture1.Picture = NONVEG_MENU.Picture3.Picture
price = 100
CUSTOMIZATION.base = price
CHECKOUT.List1.AddItem "Chicken Tikka"
Payment_Summary.List1.AddItem "Chicken Tikka"
End Sub

Private Sub HOME_Click()
Load Front_page
Front_page.Show
Unload Me
End Sub
