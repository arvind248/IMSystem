VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSSubject 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Search Subject Details"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   6855
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "category"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3000
      TabIndex        =   12
      Text            =   " "
      Top             =   4560
      Width           =   2505
   End
   Begin VB.ComboBox cmbSubjectCode 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmSSubject.frx":0000
      Left            =   3000
      List            =   "frmSSubject.frx":0002
      TabIndex        =   11
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtSubjectCode 
      BackColor       =   &H00FFFFFF&
      DataField       =   "subjectid"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5760
      TabIndex        =   4
      Text            =   " "
      Top             =   600
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3000
      TabIndex        =   3
      Text            =   " "
      Top             =   1320
      Width           =   2500
   End
   Begin VB.TextBox txtDuration 
      BackColor       =   &H00FFFFFF&
      DataField       =   "duration"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3000
      TabIndex        =   2
      Text            =   " "
      Top             =   2160
      Width           =   2500
   End
   Begin VB.TextBox txtEligibility 
      BackColor       =   &H00FFFFFF&
      DataField       =   "eligibility"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3000
      TabIndex        =   1
      Text            =   " "
      Top             =   3720
      Width           =   2505
   End
   Begin VB.TextBox txtSubjectFees 
      BackColor       =   &H00FFFFFF&
      DataField       =   "subjectfees"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3000
      TabIndex        =   0
      Text            =   " "
      Top             =   2880
      Width           =   2505
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   5520
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:/project/db/IMS.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:/project/db/IMS.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "subject"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblCourseId 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Subject Code"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label lblDuration 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duration"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Label lbPublisher 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eligibility"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   3840
      Width           =   1500
   End
   Begin VB.Label lblState 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Subject Fees"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   3000
      Width           =   1500
   End
   Begin VB.Label lblPhone 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   960
      TabIndex        =   5
      Top             =   4560
      Width           =   1500
   End
End
Attribute VB_Name = "frmSSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbSubjectCode_Click()
Adodc1.Recordset.MoveFirst
While cmbSubjectCode.Text <> txtSubjectCode.Text
Adodc1.Recordset.MoveNext
Wend
End Sub

Private Sub Form_Load()

Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Adodc1.Recordset.MoveFirst
While Adodc1.Recordset.EOF = False
cmbSubjectCode.AddItem (Adodc1.Recordset.Fields(0))
Adodc1.Recordset.MoveNext
Wend

End Sub
