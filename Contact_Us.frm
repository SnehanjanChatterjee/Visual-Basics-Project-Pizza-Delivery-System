VERSION 5.00
Begin VB.Form Contact_Us 
   Caption         =   "CONTACT US :)"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Height          =   585
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "@gmail.com"
      Top             =   6840
      Width           =   1695
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
      Height          =   1335
      Left            =   1600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   8280
      Width           =   11055
   End
   Begin VB.CommandButton Submit 
      Caption         =   "SUBMIT"
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
      Left            =   1600
      TabIndex        =   4
      Top             =   9960
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2565
      MaxLength       =   10
      TabIndex        =   1
      Top             =   5280
      Width           =   10035
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1600
      MaxLength       =   50
      TabIndex        =   0
      Top             =   3720
      Width           =   11000
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1600
      TabIndex        =   2
      Top             =   6840
      Width           =   11000
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1600
      TabIndex        =   15
      Text            =   "+91"
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mandatory Field"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   400
      Left            =   10320
      TabIndex        =   14
      Top             =   7800
      Width           =   3000
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mandatory Field"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   400
      Left            =   10320
      TabIndex        =   13
      Top             =   6120
      Width           =   3000
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mandatory Field"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   10320
      TabIndex        =   12
      Top             =   3240
      Width           =   3000
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MESSAGE *"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   1600
      TabIndex        =   10
      Top             =   7680
      Width           =   1995
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "EMAIL *"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   1600
      TabIndex        =   9
      Top             =   6120
      Width           =   1995
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE NUMBER"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   1605
      TabIndex        =   8
      Top             =   4560
      Width           =   1995
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME *"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   1605
      TabIndex        =   7
      Top             =   3000
      Width           =   1995
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WE'D LOVE TO HEAR FROM YOU!!"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      TabIndex        =   6
      Top             =   2040
      Width           =   8655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT US"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   6120
      TabIndex        =   5
      Top             =   0
      Width           =   7455
   End
   Begin VB.Image Image1 
      Height          =   10920
      Left            =   0
      Picture         =   "Contact_Us.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "Contact_Us"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim res As ADODB.Recordset

Private Sub Form_Load()
Set con = New ADODB.Connection
con.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source=CustomerInfo.accdb;"
con.Open
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
End Sub

Private Sub Submit_Click()

Set res = New ADODB.Recordset
res.Open "select * from CustomerFeedback", con, adOpenDynamic, adLockOptimistic

If Text1.Text = "" Then
    MsgBox "Enter NAME", vbCritical, "Mandatory Field"
    Label7.Visible = True
    Text1.SetFocus
Else
    Label7.Visible = False
End If

If Text3.Text = "" Then
    MsgBox "Enter MAIL ID", vbCritical, "Mandatory Field"
    Label8.Visible = True
    Text3.SetFocus
Else
    Label8.Visible = False
End If

If Text4.Text = "" Then
    MsgBox "Enter MESSAGE", vbCritical, "Mandatory Field"
    Label9.Visible = True
    Text4.SetFocus
Else
    Label9.Visible = False
End If


If (Text1.Text <> "") And (Text3.Text <> "") And (Text4.Text <> "") Then
  ans = MsgBox("ARE YOU SURE U WANT TO SUBMIT?", vbYesNo, "CONTACT US")
  If ans = vbYes Then
    res.AddNew
    res.Fields(0) = Text1.Text
    res.Fields(1) = "+91" & Text2.Text
    res.Fields(2) = Text3.Text & "@gmail.com"
    res.Fields(3) = Text4.Text
    res.Update
    MsgBox ("THANKS FOR YOUR VALUABLE FEEDBACK :)"), vbOKOnly, Text1.Text
    Unload Me
    Load Front_page
    Front_page.Show
  Else
    Text4.SetFocus
   End If
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
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

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Text3.SetFocus
End If
If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
    If (Not KeyAscii = 8) And (Not KeyAscii = 13) Then ''if backspace and enter is pressed,it's accepted
    MsgBox "MUST BE NUMBERS ONLY! PLEASE TRY AGAIN", vbOKOnly, "MOBILE NUMBER"
    KeyAscii = 0
    End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Text4.SetFocus
End If
End Sub

'Private Sub Text4_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
' Submit.SetFocus
'End If
'End Sub

Private Sub Text2_LostFocus()
If (Text2.Text <> "") And (Len(Text2.Text) < 10) Then
    MsgBox "ENTER 10 DIGIT MOBILE NUMBER", vbInformation, "Mobile Number"
    Text2.SetFocus
End If
End Sub

Private Sub Text3_LostFocus()
If InStr(Text3.Text, "@gmail.com") Or InStr(Text3.Text, "@GMAIL.COM") Or InStr(Text3.Text, "@") Then
    MsgBox "INCORRECT MAIL ID", vbInformation, "MAIL ID"
    Text3.SetFocus
End If
End Sub

'Private Sub Text1_LostFocus()
'If (Text1.Text <> "") And (IsNumeric(Text1.Text)) Then
'    MsgBox "Please enter alphabets only.", vbInformation
'    Text1.Text = ""
'    Text1.SetFocus
'End If
'End Sub

'Private Sub Text2_LostFocus()
'If (Text2.Text <> "") And (Not IsNumeric(Text2.Text)) Then
'    MsgBox "Please enter numbers only.", vbInformation
'    Text2.Text = ""
'    Text2.SetFocus
'End If
'End Sub


