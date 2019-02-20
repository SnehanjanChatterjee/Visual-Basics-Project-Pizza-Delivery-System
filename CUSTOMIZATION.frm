VERSION 5.00
Begin VB.Form CUSTOMIZATION 
   Caption         =   "CUSTOMIZE YOUR PIZZA"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "CUSTOMIZATION.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Reset 
      Caption         =   "RESET"
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
      Left            =   12000
      TabIndex        =   18
      Top             =   8880
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Rs"
      Top             =   9480
      Width           =   735
   End
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
      Left            =   1440
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   0
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   7000
      TabIndex        =   8
      Top             =   2400
      Width           =   8775
      Begin VB.OptionButton Op_CHEESEBURST 
         Caption         =   "CHEESEBURST"
         BeginProperty Font 
            Name            =   "Oswald"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   5400
         TabIndex        =   21
         Top             =   960
         Width           =   2000
      End
      Begin VB.OptionButton Op_FRESHPANPIZZA 
         Caption         =   "FRESHPAN PIZZA"
         BeginProperty Font 
            Name            =   "Oswald"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   3000
         TabIndex        =   20
         Top             =   960
         Width           =   2000
      End
      Begin VB.OptionButton Op_CLASSICHANDTOSSED 
         Caption         =   "CLASSIC HANDTOSSED"
         BeginProperty Font 
            Name            =   "Oswald"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   360
         TabIndex        =   19
         Top             =   960
         Width           =   2000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Your Crust"
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
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   2775
      End
      Begin VB.Image Image2 
         Height          =   1935
         Left            =   0
         Picture         =   "CUSTOMIZATION.frx":1CE64
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8775
      End
   End
   Begin VB.CommandButton checkprice 
      Caption         =   "CHECK    PRICE"
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
      Left            =   7000
      TabIndex        =   7
      Top             =   8880
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   4695
      Left            =   600
      ScaleHeight     =   4635
      ScaleWidth      =   4755
      TabIndex        =   6
      Top             =   3240
      Width           =   4815
   End
   Begin VB.ComboBox size 
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9720
      TabIndex        =   1
      Text            =   "size"
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   7000
      TabIndex        =   4
      Top             =   6120
      Width           =   9900
      Begin VB.CheckBox Ch_FRESHTOMATO 
         Caption         =   "FRESH TOMATO"
         BeginProperty Font 
            Name            =   "Oswald"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7560
         TabIndex        =   25
         Top             =   600
         Width           =   2025
      End
      Begin VB.CheckBox Ch_GRILLEDMUSHROOM 
         Caption         =   "GRILLED MUSHROOM"
         BeginProperty Font 
            Name            =   "Oswald"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4800
         TabIndex        =   24
         Top             =   600
         Width           =   2500
      End
      Begin VB.CheckBox Ch_GOLDENCORN 
         Caption         =   "GOLDENCORN"
         BeginProperty Font 
            Name            =   "Oswald"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2520
         TabIndex        =   23
         Top             =   600
         Width           =   2145
      End
      Begin VB.CheckBox Ch_JALAPENO 
         Caption         =   "JALAPENO"
         BeginProperty Font 
            Name            =   "Oswald"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   2145
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Toppings"
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
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   2355
         Left            =   0
         Picture         =   "CUSTOMIZATION.frx":30069
         Stretch         =   -1  'True
         Top             =   0
         Width           =   10005
      End
   End
   Begin VB.CommandButton MINUS 
      BackColor       =   &H000000FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8400
      Width           =   495
   End
   Begin VB.CommandButton PLUS 
      BackColor       =   &H0000FF00&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8400
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CHECKOUT"
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
      Left            =   12720
      TabIndex        =   0
      Top             =   9480
      Width           =   3135
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PIZZA  FACTORY"
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
      Height          =   1695
      Left            =   8280
      TabIndex        =   14
      Top             =   600
      Width           =   5295
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
      Left            =   7000
      TabIndex        =   13
      Top             =   9600
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Crust Size"
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
      Left            =   7200
      TabIndex        =   12
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      Left            =   10560
      TabIndex        =   10
      Top             =   9480
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   7000
      Picture         =   "CUSTOMIZATION.frx":3EBFA
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   9015
   End
End
Attribute VB_Name = "CUSTOMIZATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim topping1, topping2, topping3, topping4, crust, psize As Integer
Public base, sum, number As Double

Private Function data_validation() As Boolean
If (Op_CLASSICHANDTOSSED = False And Op_FRESHPANPIZZA = False And Op_CHEESEBURST = False) Then
    MsgBox "Choose Crust Type", vbOKOnly, "CRUST TYPES"
End If
If size.Text = "size" Then
    MsgBox "Choose Pizza Types", vbOKOnly, "PIZZA TYPES"
End If
If (Op_CLASSICHANDTOSSED = True) Then
    crust = 0
    Me.Op_CLASSICHANDTOSSED.BackColor = &H8080FF
    Me.Op_FRESHPANPIZZA.BackColor = &H8000000F
    Me.Op_CHEESEBURST.BackColor = &H8000000F
End If
If (Op_FRESHPANPIZZA = True) Then
    crust = 25
    Me.Op_FRESHPANPIZZA.BackColor = &H8080FF
    Me.Op_CLASSICHANDTOSSED.BackColor = &H8000000F
    Me.Op_CHEESEBURST.BackColor = &H8000000F
End If
If (Op_CHEESEBURST = True) Then
    crust = 50
    Me.Op_CHEESEBURST.BackColor = &H8080FF
    Me.Op_CLASSICHANDTOSSED.BackColor = &H8000000F
    Me.Op_FRESHPANPIZZA.BackColor = &H8000000F
End If

If size.Text = "Normal" Then
    psize = 100
End If
If size.Text = "Small" Then
    psize = 0
End If
If size.Text = "Large" Then
    psize = 200
End If

If (Ch_JALAPENO.value = 1) Then
    topping1 = 20
End If
If (Ch_GOLDENCORN.value = 1) Then
    topping2 = 30
End If
If (Ch_GRILLEDMUSHROOM.value = 1) Then
    topping3 = 30
End If
If (Ch_FRESHTOMATO.value = 1) Then
    topping4 = 40
End If
data_validation = True
If ((Op_CLASSICHANDTOSSED = False And Op_FRESHPANPIZZA = False And Op_CHEESEBURST = False) Or size.Text = "size") Then
    data_validation = False
End If
End Function


Private Sub BACK_Click()
ans = MsgBox("Are you sure u want to exit?Everything will be lost", vbYesNo, "EXIT")
If ans = vbYes Then
    Load Ordering
    Ordering.Show
    Unload Me
End If
End Sub

Private Sub Command4_Click()
If data_validation = False Then
MsgBox "COMPLETE YOUR ORDER", vbCritical, "ORDER COMPLETION"
Else
    If Label3.Caption = "" Then
        sum = number * (psize + base + topping1 + topping2 + topping3 + topping4 + crust + SIDES.sideprice)
        Label3.Caption = Label3.Caption + " " & sum
        CHECKOUT.Label2.Caption = sum
        CHECKOUT.Label10.Caption = 40
        CHECKOUT.Label11.Caption = sum + 40
        Payment_Summary.Text4 = sum + 40
        ans = MsgBox("ARE YOU SURE YOU WANT TO CHECKOUT?", vbQuestion + vbYesNo, "TOTAL PRICE=" & sum)
        If ans = vbYes Then
            Load CHECKOUT
            CHECKOUT.Show
        End If
    Else
        ans = MsgBox("ARE YOU SURE YOU WANT TO CHECKOUT?", vbQuestion + vbYesNo, "TOTAL PRICE=" & sum)
        If ans = vbYes Then
            Load CHECKOUT
            CHECKOUT.Show
        End If
    End If
End If
End Sub

Private Sub Form_Load()
size.AddItem "Normal"
size.AddItem "Small"
size.AddItem "Large"
number = 1
sum = 0
topping1 = 0
topping2 = 0
topping3 = 0
topping4 = 0
psize = 0
crust = 0
End Sub

Private Sub MINUS_Click()
If number > 1 Then
    Label3.Caption = ""
    number = number - 1
    Label5.Caption = number
End If
End Sub

Private Sub PLUS_Click()
If number > 9 Then
    MsgBox "You Can Only Order 10 Pizzas At A Time", vbInformation, "PIZZA FACTORY"
Else
    Label3.Caption = ""
    number = number + 1
    Label5.Caption = number
End If
End Sub

Private Sub Reset_Click()
Label3.Caption = ""
Op_CLASSICHANDTOSSED = False
Op_FRESHPANPIZZA = False
Op_CHEESEBURST = False
size.Text = "size"
Ch_JALAPENO.value = 0
Ch_GOLDENCORN.value = 0
Ch_GRILLEDMUSHROOM.value = 0
Ch_FRESHTOMATO.value = 0
number = 1
sum = 0
Op_CLASSICHANDTOSSED.BackColor = &H8000000F
Op_FRESHPANPIZZA.BackColor = &H8000000F
Op_CHEESEBURST.BackColor = &H8000000F
Label5.Caption = 1
End Sub

Private Sub checkprice_Click()
If data_validation = True Then
    Label3.Caption = ""
    sum = number * (psize + base + topping1 + topping2 + topping3 + topping4 + crust + SIDES.sideprice)
    Label3.Caption = Label3.Caption + " " & sum
    CHECKOUT.Label2.Caption = sum
    CHECKOUT.Label10.Caption = 40
    CHECKOUT.Label11.Caption = sum + 40
    Payment_Summary.Text4 = sum + 40
End If
End Sub

