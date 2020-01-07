VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCompletionCertificate 
   BackColor       =   &H0080FF80&
   Caption         =   "Completion Certificate"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   Picture         =   "frmCompletionCertificate.frx":0000
   ScaleHeight     =   5580
   ScaleWidth      =   7005
   Begin MSComCtl2.DTPicker DTPIssueDate 
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   9371649
      CurrentDate     =   43782
   End
   Begin VB.TextBox txtCCourseName 
      DataField       =   "name"
      DataSource      =   "Course"
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
      Left            =   5280
      TabIndex        =   15
      Text            =   " "
      Top             =   5040
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtSCourseid 
      DataField       =   "courseid"
      DataSource      =   "Student"
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
      Left            =   5880
      TabIndex        =   14
      Text            =   " "
      Top             =   2880
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox txtSName 
      DataField       =   "name"
      DataSource      =   "Student"
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
      TabIndex        =   13
      Text            =   " "
      Top             =   2280
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Certificate 
      Height          =   375
      Left            =   1320
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
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\project\db\IMS.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\project\db\IMS.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "completioncertificate"
      Caption         =   "Certificate"
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
   Begin VB.ComboBox cmbRollno 
      DataField       =   "rollno"
      DataSource      =   "Certificate"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      TabIndex        =   12
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton cmdEnroll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Issue"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Width           =   1440
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4200
      Width           =   1440
   End
   Begin VB.TextBox txtCertificateid 
      DataField       =   "certificateid"
      DataSource      =   "Certificate"
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
      Left            =   3240
      TabIndex        =   7
      Text            =   " "
      Top             =   480
      Width           =   2500
   End
   Begin VB.TextBox txtIssueDate 
      DataField       =   "issuedate"
      DataSource      =   "Certificate"
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
      TabIndex        =   6
      Text            =   " "
      Top             =   1080
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox txtRollno 
      DataField       =   "rollno"
      DataSource      =   "Student"
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
      Left            =   5880
      TabIndex        =   2
      Text            =   " "
      Top             =   1680
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox txtName 
      DataField       =   "name"
      DataSource      =   "Certificate"
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
      Left            =   3240
      TabIndex        =   1
      Text            =   " "
      Top             =   2280
      Width           =   2505
   End
   Begin VB.TextBox txtCourseName 
      DataField       =   "coursename"
      DataSource      =   "Certificate"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   2880
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Course 
      Height          =   375
      Left            =   3360
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Course"
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
   Begin MSAdodcLib.Adodc Student 
      Height          =   330
      Left            =   720
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\project\db\IMS.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\project\db\IMS.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "student"
      Caption         =   "Student"
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
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Certificate id"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label lblRollno 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Rollno"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Label lblCourseName 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   2880
      Width           =   1695
   End
End
Attribute VB_Name = "frmCompletionCertificate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub cmbRollno_Click()
Student.Recordset.MoveFirst
While cmbRollno.Text <> txtRollno.Text
Student.Recordset.MoveNext
Wend

Course.Recordset.MoveFirst
While txtSCourseid.Text <> Course.Recordset.Fields(0)
Course.Recordset.MoveNext
Wend
txtName.Text = txtSName.Text
txtCourseName.Text = txtCCourseName.Text

End Sub

Private Sub cmdEnroll_Click()
txtIssueDate.Text = DTPIssueDate.Value
Certificate.Recordset.Update
MsgBox ("Certificate Issued ")

Certificate.Recordset.MoveLast
n = Mid(Certificate.Recordset.Fields(0), 5, 3)

Certificate.Recordset.AddNew
txtCertificateid.Text = "9812" + CStr(n + 1)


End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Student.Recordset.MoveFirst
While Student.Recordset.EOF = False
cmbRollno.AddItem (Student.Recordset.Fields(0))
Student.Recordset.MoveNext
Wend

If Certificate.Recordset.EOF = True Then
n = 100
Else
Certificate.Recordset.MoveLast
n = Mid(Certificate.Recordset.Fields(0), 5, 3)
End If
Certificate.Recordset.AddNew
txtCertificateid.Text = "9812" + CStr(n + 1)


End Sub
