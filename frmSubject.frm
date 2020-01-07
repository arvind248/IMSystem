VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSubject 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Subject Details"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   Picture         =   "frmSubject.frx":0000
   ScaleHeight     =   7155
   ScaleWidth      =   7200
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2040
      Top             =   6480
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
   Begin VB.ComboBox cmbCategory 
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
      Height          =   405
      ItemData        =   "frmSubject.frx":2EDBC
      Left            =   3000
      List            =   "frmSubject.frx":2EDDB
      TabIndex        =   5
      Top             =   4080
      Width           =   3015
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
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
      Left            =   3720
      TabIndex        =   8
      Top             =   4920
      Width           =   1500
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
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
      Left            =   360
      TabIndex        =   6
      Top             =   4920
      Width           =   1500
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   2040
      TabIndex        =   7
      Top             =   4920
      Width           =   1500
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   5400
      TabIndex        =   9
      Top             =   4920
      Width           =   1500
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
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
      Left            =   360
      TabIndex        =   10
      Top             =   5760
      Width           =   1500
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
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
      Left            =   2040
      TabIndex        =   11
      Top             =   5760
      Width           =   1500
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
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
      Left            =   3720
      TabIndex        =   12
      Top             =   5760
      Width           =   1500
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
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
      Left            =   5400
      TabIndex        =   13
      Top             =   5760
      Width           =   1500
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
      Height          =   405
      Left            =   3000
      TabIndex        =   3
      Text            =   " "
      Top             =   2640
      Width           =   2985
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
      Height          =   405
      Left            =   3000
      TabIndex        =   4
      Text            =   " "
      Top             =   3360
      Width           =   2985
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
      Height          =   405
      Left            =   3000
      TabIndex        =   2
      Text            =   " "
      Top             =   1920
      Width           =   2985
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
      Height          =   405
      Left            =   3000
      TabIndex        =   1
      Text            =   " "
      Top             =   1200
      Width           =   2985
   End
   Begin VB.TextBox txtSubjectCode 
      BackColor       =   &H00FFFFFF&
      DataField       =   "subjectid"
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
      Height          =   405
      Left            =   3000
      TabIndex        =   0
      Text            =   " "
      Top             =   480
      Width           =   2985
   End
   Begin VB.Label lblPhone 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   720
      TabIndex        =   19
      Top             =   4080
      Width           =   1500
   End
   Begin VB.Label lblState 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   720
      TabIndex        =   18
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label lbPublisher 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   720
      TabIndex        =   17
      Top             =   3360
      Width           =   1500
   End
   Begin VB.Label lblDuration 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   720
      TabIndex        =   16
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackColor       =   &H00E0E0E0&
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
      Left            =   720
      TabIndex        =   15
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label lblCourseId 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   720
      TabIndex        =   14
      Top             =   480
      Width           =   1500
   End
End
Attribute VB_Name = "frmSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
Adodc1.Recordset.Delete
MsgBox ("Record Delete")
Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmdFirst_Click()
Adodc1.Recordset.MoveFirst


End Sub

Private Sub cmdLast_Click()
Adodc1.Recordset.MoveLast

End Sub

Private Sub cmdNew_Click()
If Adodc1.Recordset.EOF = True Then
n = 100
Else
Adodc1.Recordset.MoveLast
n = Mid(Adodc1.Recordset.Fields(0), 3, 3)
End If
Adodc1.Recordset.AddNew
txtSubjectCode.Text = "SB" + CStr(n + 1)
txtName.SetFocus


End Sub

Private Sub cmdNext_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
    Adodc1.Recordset.MoveLast
    MsgBox ("this is Last Record")
End If

End Sub

Private Sub cmdPrevious_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
Adodc1.Recordset.MoveFirst
MsgBox ("this is first Record")
End If

End Sub

Private Sub cmdSave_Click()

Adodc1.Recordset.Update
MsgBox ("Record Added")
End Sub

Private Sub cmdUpdate_Click()

Adodc1.Recordset.Update
MsgBox ("Record Updated")
End Sub


Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub


