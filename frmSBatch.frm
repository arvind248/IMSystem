VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSBatch 
   BackColor       =   &H80000002&
   Caption         =   "Search Batch Details "
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   5250
   Begin VB.ComboBox cmbBatchCode 
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
      Left            =   1920
      TabIndex        =   12
      Top             =   360
      Width           =   2500
   End
   Begin VB.TextBox txtTiming 
      BackColor       =   &H00FFFFFF&
      DataField       =   "timing"
      DataSource      =   "Adodc1"
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
      Left            =   1920
      TabIndex        =   5
      Text            =   " "
      Top             =   3480
      Width           =   2500
   End
   Begin VB.TextBox txtEDate 
      BackColor       =   &H00FFFFFF&
      DataField       =   "endingdate"
      DataSource      =   "Adodc1"
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
      Left            =   1920
      TabIndex        =   4
      Text            =   " "
      Top             =   4200
      Width           =   2500
   End
   Begin VB.TextBox txtSDate 
      BackColor       =   &H00FFFFFF&
      DataField       =   "startingdate"
      DataSource      =   "Adodc1"
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
      Left            =   1920
      TabIndex        =   3
      Text            =   " "
      Top             =   1920
      Width           =   2500
   End
   Begin VB.TextBox txtBatchTeacher 
      BackColor       =   &H00FFFFFF&
      DataField       =   "batchteacher"
      DataSource      =   "Adodc1"
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
      Left            =   1920
      TabIndex        =   2
      Text            =   " "
      Top             =   1080
      Width           =   2500
   End
   Begin VB.TextBox txtBatchName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "batchname"
      DataSource      =   "Adodc1"
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
      Left            =   1920
      TabIndex        =   1
      Text            =   " "
      Top             =   2640
      Width           =   2500
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
      Left            =   4440
      TabIndex        =   0
      Text            =   " "
      Top             =   360
      Visible         =   0   'False
      Width           =   705
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   4920
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
   Begin VB.Label txtEndingDate 
      BackColor       =   &H80000002&
      Caption         =   "Ending Date"
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
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   1740
   End
   Begin VB.Label lblTiming 
      BackColor       =   &H80000002&
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
      Height          =   450
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label txtStartingDate 
      BackColor       =   &H80000002&
      Caption         =   "Starting Date"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label lblDuration 
      BackColor       =   &H80000002&
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
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000002&
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
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label lblCourseId 
      BackColor       =   &H80000002&
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
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1500
   End
End
Attribute VB_Name = "frmSBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbBatchCode_Click()
Adodc1.Recordset.MoveFirst
While cmbBatchCode.Text <> txtBatchCode.Text
Adodc1.Recordset.MoveNext
Wend
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Adodc1.Recordset.MoveFirst
While Adodc1.Recordset.EOF = False
cmbBatchCode.AddItem (Adodc1.Recordset.Fields(0))
Adodc1.Recordset.MoveNext
Wend

End Sub
