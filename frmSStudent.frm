VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSStudent 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Search Student"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   9945
   Begin VB.TextBox txtCourseID 
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
      Left            =   1920
      TabIndex        =   27
      Text            =   " "
      Top             =   1680
      Width           =   2500
   End
   Begin VB.ComboBox cmbRollno 
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
      TabIndex        =   26
      Top             =   240
      Width           =   2535
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
      Left            =   4320
      TabIndex        =   12
      Text            =   " "
      Top             =   240
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox txtName 
      DataField       =   "name"
      DataSource      =   "Adodc1"
      Height          =   330
      Left            =   7080
      TabIndex        =   11
      Text            =   " "
      Top             =   240
      Width           =   2385
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
      Left            =   1920
      TabIndex        =   10
      Text            =   " "
      Top             =   960
      Width           =   2505
   End
   Begin VB.TextBox txtGender 
      DataField       =   "gender"
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
      Left            =   7080
      TabIndex        =   9
      Text            =   " "
      Top             =   960
      Width           =   2385
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "address"
      DataSource      =   "Adodc1"
      Height          =   350
      Left            =   7080
      TabIndex        =   8
      Text            =   " "
      Top             =   1680
      Width           =   2505
   End
   Begin VB.TextBox txtCity 
      DataField       =   "city"
      DataSource      =   "Adodc1"
      Height          =   350
      Left            =   1920
      TabIndex        =   7
      Text            =   " "
      Top             =   2400
      Width           =   2385
   End
   Begin VB.TextBox txtState 
      DataField       =   "state"
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
      Left            =   7080
      TabIndex        =   6
      Text            =   " "
      Top             =   2400
      Width           =   2505
   End
   Begin VB.TextBox txtPIN 
      DataField       =   "pin"
      DataSource      =   "Adodc1"
      Height          =   350
      Left            =   1920
      TabIndex        =   5
      Text            =   " "
      Top             =   3120
      Width           =   2385
   End
   Begin VB.TextBox txtFatherName 
      DataField       =   "father name"
      DataSource      =   "Adodc1"
      Height          =   350
      Left            =   7080
      TabIndex        =   4
      Text            =   " "
      Top             =   3120
      Width           =   2500
   End
   Begin VB.TextBox txtMotherName 
      DataField       =   "mother name"
      DataSource      =   "Adodc1"
      Height          =   330
      Left            =   1920
      TabIndex        =   3
      Text            =   " "
      Top             =   3840
      Width           =   2385
   End
   Begin VB.TextBox txtEmail 
      DataField       =   "email id"
      DataSource      =   "Adodc1"
      Height          =   350
      Left            =   1800
      TabIndex        =   2
      Text            =   " "
      Top             =   4560
      Width           =   2385
   End
   Begin VB.TextBox txtPhone 
      DataField       =   "phone no"
      DataSource      =   "Adodc1"
      Height          =   350
      Left            =   7080
      TabIndex        =   1
      Text            =   " "
      Top             =   3840
      Width           =   2500
   End
   Begin VB.TextBox txtNationality 
      DataField       =   "nationality"
      DataSource      =   "Adodc1"
      Height          =   350
      Left            =   7080
      TabIndex        =   0
      Text            =   " "
      Top             =   4440
      Width           =   2500
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5280
      Top             =   5160
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Course ID"
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
      Left            =   360
      TabIndex        =   28
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Label lblRollno 
      BackColor       =   &H00C0E0FF&
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
      Left            =   360
      TabIndex        =   25
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0E0FF&
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
      Left            =   5520
      TabIndex        =   24
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label lblDOB 
      BackColor       =   &H00C0E0FF&
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
      Left            =   360
      TabIndex        =   23
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label lblGener 
      BackColor       =   &H00C0E0FF&
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
      Height          =   375
      Left            =   5520
      TabIndex        =   22
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Address"
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
      Left            =   5520
      TabIndex        =   21
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Label lblCity 
      BackColor       =   &H00C0E0FF&
      Caption         =   "City"
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
      Left            =   360
      TabIndex        =   20
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label lblState 
      BackColor       =   &H00C0E0FF&
      Caption         =   "State"
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
      Left            =   5520
      TabIndex        =   19
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label lblPIN 
      BackColor       =   &H00C0E0FF&
      Caption         =   "PIN"
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
      Left            =   360
      TabIndex        =   18
      Top             =   3240
      Width           =   1500
   End
   Begin VB.Label lblEmail 
      BackColor       =   &H00C0E0FF&
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
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   4560
      Width           =   1500
   End
   Begin VB.Label lblNationality 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Nationality"
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
      Left            =   5520
      TabIndex        =   16
      Top             =   4560
      Width           =   1500
   End
   Begin VB.Label lblPhone 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Phone"
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
      Left            =   5520
      TabIndex        =   15
      Top             =   3840
      Width           =   1500
   End
   Begin VB.Label lblMotherName 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Mother Name"
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
      Left            =   360
      TabIndex        =   14
      Top             =   3840
      Width           =   1500
   End
   Begin VB.Label lblFathername 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Father Name"
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
      Left            =   5520
      TabIndex        =   13
      Top             =   3120
      Width           =   1500
   End
End
Attribute VB_Name = "frmSStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbRollno_Click()
Adodc1.Recordset.MoveFirst
While cmbRollno.Text <> txtRollno.Text
Adodc1.Recordset.MoveNext
Wend




End Sub


Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Adodc1.Recordset.MoveFirst
While Adodc1.Recordset.EOF = False
cmbRollno.AddItem (Adodc1.Recordset.Fields(0))
Adodc1.Recordset.MoveNext
Wend
End Sub

Private Sub txtFullAddress_Change()

End Sub
