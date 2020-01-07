VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Help "
   ClientHeight    =   9240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   Picture         =   "frmHelp.frx":0000
   ScaleHeight     =   9240
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":2E752
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   1200
      Width           =   9015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":2E80B
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   8
      Top             =   7680
      Width           =   9015
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":2E895
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   7
      Top             =   4080
      Width           =   9015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":2E96E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   4920
      Width           =   9015
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":2EA1B
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   5
      Top             =   5520
      Width           =   9015
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":2EC06
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   6840
      Width           =   9015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":2ED29
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3480
      Width           =   9495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":2EDF4
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   9135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "This SCHOOL MANAGEMENT SYSTEM is very beneficial in estimating the real position of the system. It has the following activities:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   9015
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
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub VScroll1_Change()

End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

