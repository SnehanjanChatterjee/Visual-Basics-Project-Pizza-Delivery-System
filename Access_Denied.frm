VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form Access_Denied 
   Caption         =   "ACCESS DENIED"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   720
      Top             =   360
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   7455
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "ACCESS IS DENIED!!   RETRY AFTER SOMETIME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   6735
      End
   End
   Begin Project1.PictureG PictureG1 
      Height          =   2055
      Left            =   2040
      Top             =   3120
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   3625
      GIF             =   "Access_Denied.frx":0000
   End
End
Attribute VB_Name = "Access_Denied"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Frame1.Visible = True
Timer1.Enabled = True
Timer1.Interval = 100
End Sub

Private Sub Timer1_Timer()

ProgressBar1.value = ProgressBar1.value + 1
Label1.Caption = ProgressBar1.value & "%"

If Label1.Caption = 100 & "%" Then
  Timer1.Enabled = False
  Frame1.Visible = False
  ctr = 0
  Unload Me
  Load Login_page
  Login_page.Show
End If
End Sub

