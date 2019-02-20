VERSION 5.00
Begin VB.Form Forgot_password 
   Caption         =   "PASSWORD CHANGE"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton HOME 
      Caption         =   "PIZZA FACTORY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Change_Password 
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   1
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   480
      TabIndex        =   9
      Top             =   2760
      Width           =   8295
      Begin VB.CommandButton Show_c_password 
         Caption         =   "SHOW PASSWORD"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6840
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Show_password 
         Caption         =   "SHOW PASSWORD"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6840
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Save 
         BackColor       =   &H0080FF80&
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         Width           =   7695
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         IMEMode         =   3  'DISABLE
         Left            =   3960
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1320
         Width           =   4000
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         IMEMode         =   3  'DISABLE
         Left            =   3960
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   480
         Width           =   4000
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CONFIRMATORY  PASSWORD"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   11
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USERID"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   480
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   7095
      Left            =   0
      Picture         =   "Forgot_password.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9300
   End
End
Attribute VB_Name = "Forgot_password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim res As ADODB.Recordset
Dim flag As Boolean
Dim flag1 As Boolean

Private Sub Form_Load()
Frame1.Visible = False
Me.Height = 3250
Me.Width = 9400
flag = False
flag1 = False
End Sub

Private Sub HOME_Click()
Unload Me
Load Front_page
Front_page.Show
End Sub

Private Sub Show_password_Click()
If Text3.Text = "" Then
    MsgBox "Enter Password", vbCritical, "Password Check"
    Text3.SetFocus
Else
    If flag = False Then
        Text3.PasswordChar = ""
        Show_password.Caption = "HIDE PASSWORD"
        flag = True
    Else
        Text3.PasswordChar = "*"
        Show_password.Caption = "SHOW PASSWORD"
        flag = False
    End If
End If
End Sub

Private Sub Show_c_password_Click()
If Text4.Text = "" Then
    MsgBox "Enter Confirmatory Password", vbCritical, "Password Check"
    Text4.SetFocus
Else
    If flag1 = False Then
        Text4.PasswordChar = ""
        Show_c_password.Caption = "HIDE PASSWORD"
        flag1 = True
    Else
        Text4.PasswordChar = "*"
        Show_c_password.Caption = "SHOW PASSWORD"
        flag1 = False
    End If
End If
End Sub

Private Sub Change_Password_Click()
Set con = New ADODB.Connection
Set res = New ADODB.Recordset
con.ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;data source=CustomerInfo.accdb;"
con.Open
res.Open "Select * from Customer where uname='" & Text1.Text & "' and uid='" & Text2.Text & "'", con, adOpenDynamic, adLockOptimistic
If Text1.Text = "" And Text2.Text = "" Then
   MsgBox "Enter Username and UserID", vbCritical, "Username and UserID Check"
   Text1.SetFocus
ElseIf Text1.Text = "" Then
   MsgBox "Enter Username", vbCritical, "Username Check"
   Text1.SetFocus
ElseIf Text2.Text = "" Then
   MsgBox "Enter UserID", vbCritical, "UserID Check"
   Text2.SetFocus
ElseIf res.EOF Then
   MsgBox "Wrong Username or UserID", vbCritical, "Username and UserID Check"
   Text1.SetFocus
Else
   Frame1.Visible = True
   Me.Height = 6430
   Me.Width = 9400
   Me.WindowState = Normal
   Text3.SetFocus
End If
End Sub

Private Sub Save_Click()

Set res = New ADODB.Recordset
res.Open "Select * from Customer where uname='" & Text1.Text & "' and uid='" & Text2.Text & "'", con, adOpenDynamic, adLockOptimistic

If Text3.Text = "" Then
    MsgBox "Enter Password", vbCritical, "Password Check"
End If
If Text4.Text = "" Then
    MsgBox "Enter Confirmatory Password", vbCritical, "Password Check"
End If
If Text3.Text <> Text4.Text Then
    Text3.Text = ""
    Text4.Text = ""
    Text3.SetFocus
    MsgBox "Passwords Donot Match", vbCritical, "Password Check"
End If
If Text3.Text <> "" And Text4.Text <> "" And Text3.Text = Text4.Text Then
    ans = MsgBox("Are U Sure U Want To Proceed?", vbYesNo, Text1.Text)
    If ans = vbYes Then
        res.Fields(2) = Text4.Text
        res.Update
        MsgBox "Password updated Successfully", vbOKOnly, Text1.Text
        Unload Me
        Load Login_page
        Login_page.Show
    Else
        Text3.SetFocus
    End If
End If

''If Frame1.Visible = True Then
''    If Text5.Text = "" Then
''        MsgBox "Enter Address", vbCritical
''    End If
''    If Text3.Text = "" Then
''        MsgBox "Enter Password", vbCritical
''    End If
''    If Text4.Text = "" Then
''        MsgBox "Enter Confirmatory Password", vbCritical
''    End If
''    If Text3.Text <> Text4.Text Then
''        Text3.Text = ""
''        Text4.Text = ""
''        Text3.SetFocus
''        MsgBox "Passwords Donot Match", vbCritical
''    End If
''    If Text3.Text <> "" And Text4.Text <> "" And Text3.Text = Text4.Text Then
''        res.Delete
''        res.AddNew
''        res.Fields(0) = Text1.Text
''        res.Fields(1) = Text2.Text
''        res.Fields(2) = Text4.Text
''        res.Fields(3) = Text5.Text
''        res.Update
''        MsgBox ("Password Successfully Changed" & " " & Text1.Text), vbInformation
''        Unload Me
''        Load Login_page
''        Login_page.Show
''     End If
''End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Change_Password.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Save.SetFocus
End If
End Sub

Private Sub Text3_LostFocus()
If (Text3.Text <> "") And (Len(Text3.Text) < 8) Then
    MsgBox "PASSWORD HAS TO BE OF MINIMUM 8 CHARACTERS", vbCritical, "Password Check"
    Text3.Text = ""
    Text3.SetFocus
End If
End Sub

Private Sub Text4_LostFocus()
If (Text4.Text <> "") And (Len(Text4.Text) < 8) Then
    MsgBox "PASSWORD HAS TO BE OF MINIMUM 8 CHARACTERS", vbCritical, "Password Check"
    Text4.Text = ""
    Text4.SetFocus
End If
End Sub
