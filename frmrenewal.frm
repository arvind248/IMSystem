VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmrenewal 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Renewal "
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbIssueId 
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
      Left            =   2160
      TabIndex        =   12
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdIssue 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Issue"
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4200
      Width           =   1380
   End
   Begin VB.TextBox txtRollno 
      DataField       =   "rollno"
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
      Left            =   2280
      TabIndex        =   8
      Text            =   " "
      Top             =   1800
      Width           =   2385
   End
   Begin VB.TextBox txtBookId 
      BackColor       =   &H00FFFFFF&
      DataField       =   "bookid"
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
      Left            =   2160
      TabIndex        =   6
      Text            =   " "
      Top             =   1080
      Width           =   2505
   End
   Begin VB.TextBox txtIDate 
      DataField       =   "issuedate"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtDDate 
      DataField       =   "duedate"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtIssueId 
      DataField       =   "issueid"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4680
      TabIndex        =   0
      Top             =   360
      Width           =   705
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   4800
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
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:/project/db/IMS.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:/project/db/IMS.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "issuebook"
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
   Begin MSComCtl2.DTPicker DTPDDate 
      Height          =   345
      Left            =   2160
      TabIndex        =   13
      Top             =   3360
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
      Format          =   92667905
      CurrentDate     =   43745
   End
   Begin VB.Label lblRollno 
      BackColor       =   &H00C0FFFF&
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
      Left            =   600
      TabIndex        =   9
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Label lblBookId 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Book ID"
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
      Left            =   600
      TabIndex        =   7
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label lblIssueId 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Issue ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      TabIndex        =   5
      Top             =   360
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Issue Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      TabIndex        =   4
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Due Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      TabIndex        =   3
      Top             =   3360
      Width           =   1125
   End
End
Attribute VB_Name = "frmrenewal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbIssueId_Click()
Adodc1.Recordset.MoveFirst
While cmbIssueId.Text <> txtIssueId.Text
Adodc1.Recordset.MoveNext
Wend
End Sub

Private Sub cmdIssue_Click()
txtDDate.Text = DTPDDate.Value
Adodc1.Recordset.Update
MsgBox ("Renewed book ")
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Adodc1.Recordset.MoveFirst
While Adodc1.Recordset.EOF = False
cmbIssueId.AddItem (Adodc1.Recordset.Fields(0))
Adodc1.Recordset.MoveNext
Wend
End Sub
