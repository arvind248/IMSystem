VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBatch 
   BackColor       =   &H80000002&
   Caption         =   "Batch Details"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   7605
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2160
      Top             =   4320
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
   Begin MSComCtl2.DTPicker DTPEDate 
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   90112001
      CurrentDate     =   41640
   End
   Begin MSComCtl2.DTPicker DTPSDate 
      Height          =   345
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   90112001
      CurrentDate     =   41640
   End
   Begin VB.TextBox txtBatchCode 
      BackColor       =   &H00FFFFFF&
      DataField       =   "batchcode"
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
      Left            =   600
      TabIndex        =   0
      Text            =   " "
      Top             =   480
      Width           =   2500
   End
   Begin VB.TextBox txtBatchName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "batchname"
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
      Left            =   4200
      TabIndex        =   1
      Text            =   " "
      Top             =   480
      Width           =   2500
   End
   Begin VB.TextBox txtBatchTeacher 
      BackColor       =   &H00FFFFFF&
      DataField       =   "batchteacher"
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
      Left            =   600
      TabIndex        =   2
      Text            =   " "
      Top             =   1320
      Width           =   2500
   End
   Begin VB.TextBox txtSDate 
      BackColor       =   &H00FFFFFF&
      DataField       =   "startingdate"
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
      Left            =   3600
      TabIndex        =   16
      Text            =   " "
      Top             =   2160
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox txtEDate 
      BackColor       =   &H00FFFFFF&
      DataField       =   "endingdate"
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
      Left            =   7200
      TabIndex        =   15
      Text            =   " "
      Top             =   2160
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H80000010&
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3720
      Width           =   1500
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H80000010&
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3720
      Width           =   1500
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H80000010&
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   1500
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H80000010&
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   1500
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H80000010&
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   1500
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000010&
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1500
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H80000010&
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1500
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H80000010&
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   1500
   End
   Begin VB.TextBox txtTiming 
      BackColor       =   &H00FFFFFF&
      DataField       =   "timing"
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
      Left            =   6840
      TabIndex        =   14
      Text            =   " "
      Top             =   1200
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.ComboBox cmbTiming 
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
      ItemData        =   "frmBatch.frx":0000
      Left            =   4200
      List            =   "frmBatch.frx":0019
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lblCourseId 
      BackColor       =   &H80000002&
      Caption         =   "Batch Code"
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
      Left            =   600
      TabIndex        =   22
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000002&
      Caption         =   "Batch Name"
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
      Left            =   4200
      TabIndex        =   21
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label lblDuration 
      BackColor       =   &H80000002&
      Caption         =   "Batch Teacher"
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
      Left            =   600
      TabIndex        =   20
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label txtStartingDate 
      BackColor       =   &H80000002&
      Caption         =   "Starting Date"
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
      Left            =   600
      TabIndex        =   19
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label lblTiming 
      BackColor       =   &H80000002&
      Caption         =   "Timing"
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
      Left            =   4200
      TabIndex        =   18
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label txtEndingDate 
      BackColor       =   &H80000002&
      Caption         =   "Ending Date"
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
      Left            =   4200
      TabIndex        =   17
      Top             =   1920
      Width           =   1740
   End
End
Attribute VB_Name = "frmBatch"
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
DTPSDate.Value = txtSDate.Text
DTPEDate.Value = txtSDate.Text
cmbTiming.Text = txtTiming.Text


End Sub

Private Sub cmdLast_Click()
Adodc1.Recordset.MoveLast
DTPSDate.Value = txtSDate.Text
DTPEDate.Value = txtSDate.Text
cmbTiming.Text = txtTiming.Text

End Sub

Private Sub cmdNew_Click()
If Adodc1.Recordset.EOF = True Then
n = 100
Else
Adodc1.Recordset.MoveLast
n = Mid(Adodc1.Recordset.Fields(0), 5, 3)
End If
Adodc1.Recordset.AddNew
txtBatchCode.Text = "BTCH" + CStr(n + 1)
txtBatchName.SetFocus

End Sub

Private Sub cmdNext_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
Adodc1.Recordset.MoveFirst
MsgBox ("this is Last Record")
End If

DTPSDate.Value = txtSDate.Text
DTPEDate.Value = txtSDate.Text
cmbTiming.Text = txtTiming.Text

End Sub

Private Sub cmdPrevious_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
Adodc1.Recordset.MoveFirst
MsgBox ("this is first Record")
End If


DTPSDate.Value = txtSDate.Text
DTPEDate.Value = txtSDate.Text
cmbTiming.Text = txtTiming.Text


End Sub

Private Sub cmdSave_Click()
txtSDate.Text = DTPSDate.Value
txtEDate.Text = DTPEDate.Value
txtTiming.Text = cmbTiming.Text

Adodc1.Recordset.Update
MsgBox ("Record Added")

End Sub

Private Sub cmdUpdate_Click()
txtSDate.Text = DTPSDate.Value
txtSDate.Text = DTPEDate.Value
txtTiming.Text = cmbTiming.Text

Adodc1.Recordset.Update
MsgBox ("Record Updated")

End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

End Sub

