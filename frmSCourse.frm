VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSCourse 
   BackColor       =   &H0080C0FF&
   Caption         =   "Search Coure Details"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   7335
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      DataField       =   "eligibility"
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
      Left            =   240
      TabIndex        =   16
      Text            =   " "
      Top             =   2160
      Width           =   2500
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      DataField       =   "coursetype"
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
      Left            =   3840
      TabIndex        =   15
      Text            =   " "
      Top             =   2160
      Width           =   2500
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "catefory"
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
      Left            =   3840
      TabIndex        =   14
      Text            =   " "
      Top             =   3120
      Width           =   2500
   End
   Begin VB.ComboBox cmbCourseId 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmSCourse.frx":0000
      Left            =   240
      List            =   "frmSCourse.frx":0002
      TabIndex        =   13
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtCourseFees 
      BackColor       =   &H00FFFFFF&
      DataField       =   "coursefees"
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
      Height          =   345
      Left            =   3840
      TabIndex        =   4
      Text            =   " "
      Top             =   1320
      Width           =   2505
   End
   Begin VB.TextBox txtSemester 
      BackColor       =   &H00FFFFFF&
      DataField       =   "noofsemeter"
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
      Height          =   350
      Left            =   240
      TabIndex        =   3
      Text            =   " "
      Top             =   3120
      Width           =   2500
   End
   Begin VB.TextBox txtDuration 
      BackColor       =   &H00FFFFFF&
      DataField       =   "duraion"
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
      Left            =   240
      TabIndex        =   2
      Text            =   " "
      Top             =   1320
      Width           =   2500
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "name"
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
      Left            =   3840
      TabIndex        =   1
      Text            =   " "
      Top             =   480
      Width           =   2500
   End
   Begin VB.TextBox txtCourseId 
      BackColor       =   &H00FFFFFF&
      DataField       =   "courseid"
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
      Left            =   3000
      TabIndex        =   0
      Text            =   " "
      Top             =   480
      Visible         =   0   'False
      Width           =   705
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   3480
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      RecordSource    =   "course"
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
   Begin VB.Label lblFathername 
      BackColor       =   &H0080C0FF&
      Caption         =   "No. of Semester"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   1740
   End
   Begin VB.Label lblPhone 
      BackColor       =   &H0080C0FF&
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3840
      TabIndex        =   11
      Top             =   2880
      Width           =   1500
   End
   Begin VB.Label lblPIN 
      BackColor       =   &H0080C0FF&
      Caption         =   "Course type"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label lblState 
      BackColor       =   &H0080C0FF&
      Caption         =   "Course Fees"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label lbPublisher 
      BackColor       =   &H0080C0FF&
      Caption         =   "Eligibility"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label lblDuration 
      BackColor       =   &H0080C0FF&
      Caption         =   "Duration"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackColor       =   &H0080C0FF&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label lblCourseId 
      BackColor       =   &H0080C0FF&
      Caption         =   "Course ID"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1500
   End
End
Attribute VB_Name = "frmSCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Adodc1.Recordset.MoveFirst
While Adodc1.Recordset.EOF = False
cmbCourseid.AddItem (Adodc1.Recordset.Fields(0))
Adodc1.Recordset.MoveNext
Wend

End Sub
Private Sub cmbCourseid_Click()
Adodc1.Recordset.MoveFirst
While cmbCourseid.Text <> txtCourseId.Text
Adodc1.Recordset.MoveNext
Wend
End Sub
