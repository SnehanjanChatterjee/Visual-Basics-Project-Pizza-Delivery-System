VERSION 5.00
Begin VB.Form Payment_Summary 
   Caption         =   "PAYMENT"
   ClientHeight    =   9780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   12
      Top             =   9240
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "RS"
      Top             =   7680
      Width           =   615
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   7680
      Width           =   4395
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2500
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2640
      Width           =   5000
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2500
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Width           =   5000
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2500
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   5000
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   2500
      TabIndex        =   3
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Label Label6 
      Caption         =   "TOTAL PRICE"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "THANK YOU FOR YOU ORDER ! ! YOUR ORDER WILL ARRIVE SHORTLY AT THE GIVEN LOCATION"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   8280
      Width           =   7455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDER"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   4
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   500
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NUMBER"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   500
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   500
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Payment_Summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Front_page
Unload CHECKOUT
Unload CUSTOMIZATION
Unload NONVEG_MENU
Unload VEG_MENU
Unload SPECIAL_MENU
Unload SIDES
Unload Ordering
Unload Payment_Summary
End Sub

Private Sub Form_Load()
Text1.Text = Login_page.billname
Text2.Text = Login_page.billnumber
Text3.Text = Login_page.billaddr
End Sub
