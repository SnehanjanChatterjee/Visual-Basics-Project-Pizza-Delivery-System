VERSION 5.00
Begin VB.Form Ordering 
   Caption         =   "ORDER"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "Ordering.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "VIEW ALL"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15600
      TabIndex        =   9
      Top             =   7920
      Width           =   4095
   End
   Begin VB.CommandButton HOME 
      Caption         =   "PIZZA FACTORY"
      BeginProperty Font 
         Name            =   "Adobe Gothic Std B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton VIEWALLVEG 
      Caption         =   "VIEW ALL"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   480
      TabIndex        =   2
      Top             =   7920
      Width           =   4455
   End
   Begin VB.CommandButton VIEWALLNVEG 
      Caption         =   "VIEW ALL"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      MousePointer    =   4  'Icon
      TabIndex        =   1
      Top             =   7920
      Width           =   4575
   End
   Begin VB.CommandButton VIEWALLSPECIAL 
      Caption         =   "VIEW ALL"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   0
      Top             =   7920
      Width           =   4455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "SIDES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   15840
      TabIndex        =   8
      Top             =   7200
      Width           =   3375
   End
   Begin VB.Image Image4 
      Height          =   3840
      Left            =   15480
      Picture         =   "Ordering.frx":1CE64
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   4200
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "       PIZZA                 FACTORY"
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
      Height          =   2175
      Left            =   8160
      TabIndex        =   6
      Top             =   720
      Width           =   4935
   End
   Begin VB.Image Image2 
      Height          =   3855
      Left            =   5400
      Picture         =   "Ordering.frx":219A1
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "VEG PIZZA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   7200
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NON VEG PIZZA"
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
      Height          =   615
      Left            =   5400
      TabIndex        =   4
      Top             =   7200
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   3855
      Left            =   480
      Picture         =   "Ordering.frx":46700
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SPECIAL PIZZA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10440
      TabIndex        =   3
      Top             =   7200
      Width           =   4455
   End
   Begin VB.Image Image3 
      Height          =   3855
      Left            =   10440
      Picture         =   "Ordering.frx":6803B
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   4455
   End
End
Attribute VB_Name = "Ordering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load SIDES
SIDES.Show
Unload Me
End Sub

Private Sub HOME_Click()
CHECKOUT.List1.CLEAR
Payment_Summary.List1.CLEAR
Load Front_page
Front_page.Show
Unload Me
End Sub

Private Sub VIEWALLVEG_Click()
Load VEG_MENU
VEG_MENU.Show
Unload Me
End Sub

Private Sub VIEWALLNVEG_Click()
Load NONVEG_MENU
NONVEG_MENU.Show
Unload Me
End Sub

Private Sub VIEWALLSPECIAL_Click()
Load SPECIAL_MENU
SPECIAL_MENU.Show
Unload Me
End Sub

Private Sub Form_Load()
Me.Height = 10000
Me.Width = 15000
End Sub

