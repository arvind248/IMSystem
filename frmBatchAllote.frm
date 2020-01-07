VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBatchAllote 
   BackColor       =   &H00FFFF80&
   Caption         =   "Batch Allotment"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   Picture         =   "frmBatchAllote.frx":0000
   ScaleHeight     =   4260
   ScaleWidth      =   10545
   Begin VB.TextBox txtBABatchCode 
      DataField       =   "batchcode"
      DataSource      =   "BatchAllote"
      Height          =   375
      Left            =   6960
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtBARollno 
      DataField       =   "rollno"
      DataSource      =   "BatchAllote"
      Height          =   375
      Left            =   2160
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtBAName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "batchname"
      DataSource      =   "BatchAllote"
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
      Left            =   9600
      TabIndex        =   23
      Text            =   " "
      Top             =   1200
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtBATeacher 
      BackColor       =   &H00FFFFFF&
      DataField       =   "batchteacher"
      DataSource      =   "BatchAllote"
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
      Left            =   9600
      TabIndex        =   22
      Text            =   " "
      Top             =   1800
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtBATime 
      BackColor       =   &H00FFFFFF&
      DataField       =   "batchtime"
      DataSource      =   "BatchAllote"
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
      Left            =   9600
      TabIndex        =   21
      Text            =   " "
      Top             =   2400
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtBAStudentName 
      DataField       =   "studentname"
      DataSource      =   "BatchAllote"
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
      Left            =   4680
      TabIndex        =   20
      Text            =   " "
      Top             =   1200
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.ComboBox cmbRollno 
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
      Left            =   2160
      TabIndex        =   19
      Top             =   600
      Width           =   2535
   End
   Begin VB.ComboBox cmbBatchCode 
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
      Left            =   6960
      TabIndex        =   18
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3240
      Width           =   1380
   End
   Begin VB.CommandButton cmdAllot 
      BackColor       =   &H00FFFFC0&
      Caption         =   "allot"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3240
      Width           =   1380
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
      Left            =   4680
      TabIndex        =   11
      Text            =   " "
      Top             =   600
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txtName 
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
      Left            =   2160
      TabIndex        =   10
      Text            =   " "
      Top             =   1200
      Width           =   2505
   End
   Begin VB.TextBox txtDOB 
      DataField       =   "dob"
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
      Left            =   2160
      TabIndex        =   9
      Text            =   " "
      Top             =   2400
      Width           =   2505
   End
   Begin VB.TextBox txtGender 
      DataField       =   "gender"
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
      Height          =   350
      Left            =   2160
      TabIndex        =   8
      Text            =   " "
      Top             =   1800
      Width           =   2505
   End
   Begin VB.TextBox txtBatchTime 
      BackColor       =   &H00FFFFFF&
      DataField       =   "timing"
      DataSource      =   "Batch"
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
      Left            =   6960
      TabIndex        =   7
      Text            =   " "
      Top             =   2400
      Width           =   2500
   End
   Begin VB.TextBox txtBatchTeacher 
      BackColor       =   &H00FFFFFF&
      DataField       =   "batchteacher"
      DataSource      =   "Batch"
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
      Left            =   6960
      TabIndex        =   2
      Text            =   " "
      Top             =   1800
      Width           =   2500
   End
   Begin VB.TextBox txtBatchName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "batchname"
      DataSource      =   "Batch"
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
      Left            =   6960
      TabIndex        =   1
      Text            =   " "
      Top             =   1200
      Width           =   2500
   End
   Begin VB.TextBox txtBatchCode 
      BackColor       =   &H00FFFFFF&
      DataField       =   "batchcode"
      DataSource      =   "Batch"
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
      Left            =   9480
      TabIndex        =   0
      Text            =   " "
      Top             =   600
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Student 
      Height          =   330
      Left            =   1200
      Top             =   2880
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
   Begin MSAdodcLib.Adodc Batch 
      Height          =   375
      Left            =   6360
      Top             =   2880
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
      RecordSource    =   "batch"
      Caption         =   "Batch"
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
   Begin MSAdodcLib.Adodc BatchAllote 
      Height          =   330
      Left            =   720
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      RecordSource    =   "batchallotement"
      Caption         =   "BatchAllote"
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
   Begin VB.Label lblRollno 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Rollno"
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
      Left            =   840
      TabIndex        =   15
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
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
      Left            =   840
      TabIndex        =   14
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label lblDOB 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of birth"
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
      Left            =   720
      TabIndex        =   13
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label lblGener 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Left            =   720
      TabIndex        =   12
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Label lblTiming 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Timing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5400
      TabIndex        =   6
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label lblDuration 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Teacher"
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
      Left            =   5400
      TabIndex        =   5
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Name"
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
      Left            =   5400
      TabIndex        =   4
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label lblCourseId 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Code"
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
      Left            =   5400
      TabIndex        =   3
      Top             =   600
      Width           =   1500
   End
End
Attribute VB_Name = "frmBatchAllote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbBatchCode_Click()

Batch.Recordset.MoveFirst
While cmbBatchCode.Text <> txtBatchCode.Text
Batch.Recordset.MoveNext
Wend

End Sub

Private Sub cmbRollno_Click()
Student.Recordset.MoveFirst
While cmbRollno.Text <> txtRollno.Text
Student.Recordset.MoveNext
Wend
End Sub



Private Sub cmdAllot_Click()
txtBARollno.Text = txtRollno.Text
txtBABatchCode.Text = txtBatchCode.Text
txtBAStudentName.Text = txtName.Text
txtBAName.Text = txtBatchName.Text
txtBATeacher.Text = txtBatchTeacher.Text
txtBATime.Text = txtBatchTime.Text

BatchAllote.Recordset.Update
BatchAllote.Recordset.AddNew
MsgBox ("Batch Alloted")

Student.Recordset.MoveLast
Student.Recordset.MoveNext
cmbRollno.Text = ""


Batch.Recordset.MoveLast
Batch.Recordset.MoveNext
cmbBatchCode.Text = ""

End Sub

Private Sub cmdCancel_Click()
unlaod Me

End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

Student.Recordset.MoveFirst
While Student.Recordset.EOF = False
cmbRollno.AddItem (Student.Recordset.Fields(0))
Student.Recordset.MoveNext
Wend


Batch.Recordset.MoveFirst
While Batch.Recordset.EOF = False
cmbBatchCode.AddItem (Batch.Recordset.Fields(0))
Batch.Recordset.MoveNext
Wend

BatchAllote.Recordset.AddNew
End Sub
