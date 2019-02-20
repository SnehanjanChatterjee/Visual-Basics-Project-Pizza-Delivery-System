VERSION 5.00
Begin VB.Form CHECKOUT 
   Caption         =   "CHECKOUT"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "CHECKOUT.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton MODIFY 
      Caption         =   "MODIFY"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10000
      TabIndex        =   18
      Top             =   6720
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "RS"
      Top             =   8520
      Width           =   615
   End
   Begin VB.TextBox Text3 
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
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "RS"
      Top             =   7680
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   15
      Text            =   "RS"
      Top             =   9480
      Width           =   615
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      Left            =   2160
      TabIndex        =   4
      Top             =   3600
      Width           =   3375
   End
   Begin VB.CommandButton Proceedtopayment 
      Caption         =   "Proceed To Payment"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12720
      MaskColor       =   &H0000FF00&
      TabIndex        =   3
      Top             =   9480
      Width           =   2775
   End
   Begin VB.TextBox name_text 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10000
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   0
      Top             =   3120
      Width           =   5055
   End
   Begin VB.TextBox addr_text 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10000
      Locked          =   -1  'True
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4200
      Width           =   5055
   End
   Begin VB.TextBox number_text 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10000
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   2
      Top             =   5880
      Width           =   5055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "      PIZZA        FACTORY"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   8040
      TabIndex        =   19
      Top             =   600
      Width           =   5175
   End
   Begin VB.Label Label11 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10560
      TabIndex        =   14
      Top             =   9480
      Width           =   1455
   End
   Begin VB.Label Label10 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10560
      TabIndex        =   13
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "DELIVERY"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   6840
      TabIndex        =   12
      Top             =   8520
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "FINAL NET AMOUNT"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   975
      Left            =   6840
      TabIndex        =   11
      Top             =   9360
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Your Order Summary"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Price"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   6795
      TabIndex        =   9
      Top             =   7680
      Width           =   2535
   End
   Begin VB.Label Label2 
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
      Left            =   10560
      TabIndex        =   8
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Your Name"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   735
      Left            =   6795
      TabIndex        =   7
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Your Address"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   975
      Left            =   6795
      TabIndex        =   6
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Mobile No."
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   855
      Left            =   6795
      TabIndex        =   5
      Top             =   6000
      Width           =   2055
   End
End
Attribute VB_Name = "CHECKOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gst, netprice As Double

Private Sub MODIFY_Click()
name_text.Locked = False
addr_text.Locked = False
number_text.Locked = False
name_text.SetFocus
End Sub

Private Sub Form_Load()
name_text = Login_page.billname
addr_text = Login_page.billaddr
number_text = Login_page.billnumber
'Label2.Caption = CUSTOMIZATION.sum
'Label2.Caption = CUSTOMIZATION.totalprice
'Label10.Caption = 40
'netprice = 40 + CUSTOMIZATION.totalprice
'Label11.Caption = netprice
End Sub

Private Sub number_text_LostFocus()
If Len(number_text) < 13 Then
    MsgBox "Mobile Number Has To Be Of 10 Digits!!", vbInformation, "Mobile Number"
    number_text.SetFocus
End If
End Sub

Private Sub Proceedtopayment_Click()
ans = MsgBox("Are you sure you want to proceed?", vbQuestion + vbYesNo, "Payment")
If ans = vbYes Then
    Load Payment_Summary
    Payment_Summary.Show
End If
End Sub

Private Sub name_text_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    addr_text.SetFocus
End If
Select Case KeyAscii
    Case 32 To 64, 91 To 96, 123 To 126
        If Not KeyAscii = 32 Then ''if spacebar is pressed,it's accepted
            MsgBox "MUST BE ALPHABETS ONLY! PLEASE TRY AGAIN", vbOKOnly, "NAME"
            KeyAscii = 0
            Exit Sub
        End If
End Select
End Sub

Private Sub addr_text_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    number_text.SetFocus
End If
End Sub

Private Sub number_text_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Proceedtopayment.SetFocus
End If
If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
    If (Not KeyAscii = 8) And (Not KeyAscii = 13) Then ''if backspace and enter is pressed,it's accepted
    MsgBox "MUST BE NUMBERS ONLY! PLEASE TRY AGAIN", vbOKOnly, "MOBILE NUMBER"
    KeyAscii = 0
    End If
End If
End Sub
