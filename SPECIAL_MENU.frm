VERSION 5.00
Begin VB.Form SPECIAL_MENU 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SPECIAL-MENU"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "SPECIAL_MENU.frx":0000
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
   Begin VB.PictureBox Picture2 
      Height          =   3975
      Left            =   8520
      Picture         =   "SPECIAL_MENU.frx":1CE64
      ScaleHeight     =   3915
      ScaleWidth      =   3915
      TabIndex        =   5
      Top             =   3360
      Width           =   3975
   End
   Begin VB.PictureBox Picture3 
      Height          =   3975
      Left            =   13080
      Picture         =   "SPECIAL_MENU.frx":33E10
      ScaleHeight     =   3915
      ScaleWidth      =   3915
      TabIndex        =   4
      Top             =   3360
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   3840
      Picture         =   "SPECIAL_MENU.frx":58B6F
      ScaleHeight     =   3915
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   3360
      Width           =   4095
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
      Left            =   3840
      TabIndex        =   2
      Top             =   8040
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
      Left            =   8520
      TabIndex        =   1
      Top             =   8040
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
      Left            =   13080
      TabIndex        =   0
      Top             =   8040
      Width           =   4095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "   PIZZA    FACTORY"
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
      Height          =   2055
      Left            =   7920
      TabIndex        =   9
      Top             =   720
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "VEGGIE PARADISE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   7440
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "CHICKEN LOADED"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8400
      TabIndex        =   7
      Top             =   7440
      Width           =   4095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "NONVEG SUPREME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
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
      Top             =   7440
      Width           =   3975
   End
End
Attribute VB_Name = "SPECIAL_MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public base As Double
Dim a, price As Integer
Public flag2 As Integer

Private Sub Form_Load()
flag2 = 0
Me.Height = 11520
Me.Width = 20490
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

Function value(a As Integer)
flag2 = 1
Load CUSTOMIZATION
CUSTOMIZATION.Show
Unload Me
End Function

Private Sub Command1_Click()
value (1)
CUSTOMIZATION.Picture1.Picture = SPECIAL_MENU.Picture1.Picture
price = 300
CUSTOMIZATION.base = price
CHECKOUT.List1.AddItem "Veggie Paradise"
Payment_Summary.List1.AddItem "Veggie Paradise"
End Sub

Private Sub Command2_Click()
value (1)
CUSTOMIZATION.Picture1.Picture = SPECIAL_MENU.Picture2.Picture
price = 200
CUSTOMIZATION.base = price
CHECKOUT.List1.AddItem "Chicken Loaded"
Payment_Summary.List1.AddItem "Chicken Loaded"
End Sub

Private Sub Command3_Click()
value (1)
CUSTOMIZATION.Picture1.Picture = SPECIAL_MENU.Picture3.Picture
price = 300
CUSTOMIZATION.base = price
CHECKOUT.List1.AddItem "Nonveg Supreme"
Payment_Summary.List1.AddItem "Nonveg Supreme"
End Sub



