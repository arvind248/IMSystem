VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplashForm 
   BackColor       =   &H8000000C&
   Caption         =   "Splash form"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4575
   ScaleWidth      =   6030
   Begin VB.Timer Timer1 
      Left            =   1800
      Top             =   3600
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   265
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblName 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "Arvind "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lbldvlpd 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By:-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblTopic 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "Institute Management System"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmSplashForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Timer1.Enabled = True
Timer1.Interval = 500
ProgressBar1.Visible = True
ProgressBar1.Max = 10
End Sub

Private Sub Timer1_Timer()
Static inttime
If IsEmpty(inttime) Then
    inttime = 1
End If

ProgressBar1.Value = inttime
If inttime = ProgressBar1.Max Then
Unload Me

frmLogIn.Show

Else
    inttime = inttime + 1
End If

End Sub
