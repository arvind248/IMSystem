VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEnquiry 
   BackColor       =   &H0000C0C0&
   Caption         =   "Enquiry"
   ClientHeight    =   9060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   Picture         =   "frmEnquiry.frx":0000
   ScaleHeight     =   9060
   ScaleWidth      =   10095
   Begin VB.ComboBox cmbGender 
      DataField       =   "gender"
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
      ItemData        =   "frmEnquiry.frx":0342
      Left            =   3960
      List            =   "frmEnquiry.frx":034C
      TabIndex        =   4
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtMobileno 
      DataField       =   "mobileno"
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
      Left            =   3960
      TabIndex        =   8
      Text            =   " "
      Top             =   5040
      Width           =   2505
   End
   Begin VB.TextBox txtEmail 
      DataField       =   "emailid"
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
      Left            =   3960
      TabIndex        =   9
      Text            =   " "
      Top             =   5640
      Width           =   2505
   End
   Begin VB.TextBox txtDOB 
      DataField       =   "dob"
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
      Left            =   6480
      TabIndex        =   21
      Text            =   " "
      Top             =   1920
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.ComboBox cmbPassingYear 
      DataField       =   "passingyear"
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
      ItemData        =   "frmEnquiry.frx":035E
      Left            =   3960
      List            =   "frmEnquiry.frx":0383
      TabIndex        =   6
      Top             =   3840
      Width           =   2535
   End
   Begin VB.ComboBox cmbQualification 
      DataField       =   "highestqualification"
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
      ItemData        =   "frmEnquiry.frx":03C9
      Left            =   3960
      List            =   "frmEnquiry.frx":03D6
      TabIndex        =   5
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox txtCandidateName 
      DataField       =   "candidatename"
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
      Left            =   3960
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtEnquiryno 
      DataField       =   "enquiryno"
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
      Height          =   400
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   2500
   End
   Begin VB.TextBox txtEDate 
      DataField       =   "enquirydate"
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
      Left            =   8400
      TabIndex        =   20
      Text            =   " "
      Top             =   720
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtCourseName 
      DataField       =   "coursename"
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
      Left            =   3960
      TabIndex        =   7
      Top             =   4440
      Width           =   2535
   End
   Begin VB.ComboBox cmbRegistrationFee 
      DataField       =   "rgistrationfee"
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
      ItemData        =   "frmEnquiry.frx":03FD
      Left            =   3960
      List            =   "frmEnquiry.frx":0407
      TabIndex        =   10
      Top             =   6240
      Width           =   2535
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00E0E0E0&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8400
      Width           =   1800
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00E0E0E0&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8400
      Width           =   1800
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8400
      Width           =   1800
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8400
      Width           =   1800
   End
   Begin VB.CommandButton cmdEnroll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enroll"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7680
      Width           =   1800
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7680
      Width           =   1800
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7680
      Width           =   1800
   End
   Begin VB.CommandButton cmdNewEnquiry 
      BackColor       =   &H00E0E0E0&
      Caption         =   "New Enquiry"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7680
      Width           =   1800
   End
   Begin VB.ComboBox cmbStatus 
      DataField       =   "status"
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
      ItemData        =   "frmEnquiry.frx":041B
      Left            =   3960
      List            =   "frmEnquiry.frx":0428
      TabIndex        =   11
      Top             =   6840
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6720
      Top             =   5040
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
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:/project/db/IMS.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:/project/db/IMS.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "enquirydetails"
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
   Begin MSComCtl2.DTPicker DTPDOB 
      Height          =   345
      Left            =   3960
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   92405761
      CurrentDate     =   41640
   End
   Begin MSComCtl2.DTPicker DTPEDate 
      Height          =   345
      Left            =   5640
      TabIndex        =   1
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   92405761
      CurrentDate     =   41640
   End
   Begin VB.Label lblphone1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Mobile  no."
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
      Left            =   1800
      TabIndex        =   33
      Top             =   5160
      Width           =   1500
   End
   Begin VB.Label lblEmail 
      BackColor       =   &H0000C0C0&
      Caption         =   "Email ID"
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
      Left            =   1800
      TabIndex        =   32
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lblGener 
      BackColor       =   &H0000C0C0&
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
      Left            =   1800
      TabIndex        =   31
      Top             =   2520
      Width           =   1500
   End
   Begin VB.Label lblDOB 
      BackColor       =   &H0000C0C0&
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
      Left            =   1800
      TabIndex        =   30
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "Passing Year"
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
      Left            =   1800
      TabIndex        =   29
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Highest Qualification"
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
      Left            =   1800
      TabIndex        =   28
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "Candidate Name"
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
      Left            =   1800
      TabIndex        =   27
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "Enquiry No."
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
      Left            =   1440
      TabIndex        =   26
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "Enquiry Date "
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
      Left            =   5640
      TabIndex        =   25
      Top             =   360
      Width           =   1500
   End
   Begin VB.Label lblCourseName 
      BackColor       =   &H0000C0C0&
      Caption         =   "Course Name"
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
      Left            =   1800
      TabIndex        =   24
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "Registration Fee"
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
      Left            =   1800
      TabIndex        =   23
      Top             =   6240
      Width           =   1860
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "Status"
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
      Left            =   1800
      TabIndex        =   22
      Top             =   6840
      Width           =   1860
   End
End
Attribute VB_Name = "frmEnquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnroll_Click()
Unload Me
frmStudent.Show
End Sub

Private Sub cmdFirst_Click()
 DTPEDate.Value = txtEDate.Text
 DTPDOB.Value = txtDOB.Text
Adodc1.Recordset.MoveFirst

End Sub

Private Sub cmdLast_Click()
 DTPEDate.Value = txtEDate.Text
 DTPDOB.Value = txtDOB.Text
Adodc1.Recordset.MoveLast

End Sub

Private Sub cmdNewEnquiry_Click()

If Adodc1.Recordset.BOF = True Then
n = 100
Else
Adodc1.Recordset.MoveLast
n = Mid(Adodc1.Recordset.Fields(0), 3, 3)
End If
Adodc1.Recordset.AddNew
txtEnquiryno.Text = "EN" + CStr(n + 1)
DTPEDate.SetFocus



End Sub

Private Sub cmdNext_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
    Adodc1.Recordset.MoveLast
      MsgBox ("this is Last Record")
End If
DTPEDate.Value = txtEDate.Text
 DTPDOB.Value = txtDOB.Text

End Sub

Private Sub cmdPrevious_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
     Adodc1.Recordset.MoveFirst
    MsgBox ("this is First Record")
End If
DTPEDate.Value = txtEDate.Text
 DTPDOB.Value = txtDOB.Text

End Sub

Private Sub cmdSave_Click()
txtEDate.Text = DTPEDate.Value
txtDOB.Text = DTPDOB.Value

Adodc1.Recordset.Update
MsgBox ("Record Added")

End Sub

Private Sub cmdUpdate_Click()
txtEDate.Text = DTPEDate.Value
txtDOB.Text = DTPDOB.Value

Adodc1.Recordset.Update
MsgBox ("Record Updated")
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

End Sub










