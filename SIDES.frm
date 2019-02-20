VERSION 5.00
Begin VB.Form SIDES 
   Caption         =   "DELICIOUS SIDES"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "SIDES.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   4530
      Left            =   840
      Picture         =   "SIDES.frx":1CE64
      ScaleHeight     =   4470
      ScaleWidth      =   6405
      TabIndex        =   5
      Top             =   3000
      Width           =   6465
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   4545
      Left            =   7920
      Picture         =   "SIDES.frx":2436C
      ScaleHeight     =   4485
      ScaleWidth      =   6420
      TabIndex        =   4
      Top             =   3000
      Width           =   6480
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   4155
      Left            =   15360
      Picture         =   "SIDES.frx":51964
      ScaleHeight     =   4095
      ScaleWidth      =   4095
      TabIndex        =   3
      Top             =   3000
      Width           =   4155
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
      Left            =   1680
      TabIndex        =   2
      Top             =   8640
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
      Left            =   8880
      TabIndex        =   1
      Top             =   8520
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
      Left            =   15480
      TabIndex        =   0
      Top             =   8520
      Width           =   4095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  PIZZA  FACTORY"
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
      Left            =   7800
      TabIndex        =   9
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "GARLIC BREADSTICK"
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
      Left            =   840
      TabIndex        =   8
      Top             =   7680
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "BURGER PIZZA"
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
      Left            =   7920
      TabIndex        =   7
      Top             =   7680
      Width           =   5895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "BUTTERSOCH "
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
      Left            =   15120
      TabIndex        =   6
      Top             =   7680
      Width           =   4815
   End
End
Attribute VB_Name = "SIDES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public special, sideprice As Double
Private Sub Command1_Click()
'Me.special = Order.sum + 99
sideprice = 99
CHECKOUT.List1.AddItem "Garlic Breadstick"
Payment_Summary.List1.AddItem "Garlic Breadstick"
Load Ordering
Ordering.Show
Unload Me
End Sub

Private Sub Command2_Click()
'Me.special = Order.sum + 99
sideprice = 99
CHECKOUT.List1.AddItem "Burger Pizza"
Payment_Summary.List1.AddItem "Burger Pizza"
Load Ordering
Ordering.Show
Unload Me
End Sub

Private Sub Command3_Click()
'Me.special = Order.sum + 99
sideprice = 99
CHECKOUT.List1.AddItem "BUTTERSCOTCH"
Payment_Summary.List1.AddItem "BUTTERSCOTCH"
Load Ordering
Ordering.Show
Unload Me
End Sub




