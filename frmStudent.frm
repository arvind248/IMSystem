VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStudent 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Student Details"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   Picture         =   "frmStudent.frx":0000
   ScaleHeight     =   9075
   ScaleWidth      =   14325
   Begin VB.TextBox txtADate 
      DataField       =   "admissiondate"
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
      Left            =   6360
      TabIndex        =   51
      Text            =   " "
      Top             =   6360
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtNationality 
      DataField       =   "nationality"
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
      Left            =   9480
      TabIndex        =   14
      Top             =   6360
      Width           =   2745
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   960
      Top             =   7200
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
      RecordSource    =   "studentfees"
      Caption         =   "Adodc3"
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
   Begin VB.TextBox txtSFDCourseFees 
      DataField       =   "coursefee"
      DataSource      =   "Adodc3"
      Height          =   495
      Left            =   240
      TabIndex        =   44
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtSFDRollno 
      DataField       =   "rollno"
      DataSource      =   "Adodc3"
      Height          =   375
      Left            =   6360
      TabIndex        =   43
      Top             =   480
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.TextBox txtCID 
      DataField       =   "courseid"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6360
      TabIndex        =   42
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   4920
      Top             =   6840
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
      Caption         =   "Adodc2"
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
   Begin VB.TextBox txtCourseName 
      DataField       =   "name"
      DataSource      =   "Adodc2"
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
      Left            =   9600
      TabIndex        =   5
      Top             =   2160
      Width           =   2775
   End
   Begin VB.ComboBox cmbGender 
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
      ItemData        =   "frmStudent.frx":20438
      Left            =   9600
      List            =   "frmStudent.frx":20442
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtCourseId 
      DataField       =   "courseid"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   6960
      TabIndex        =   40
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbCourseId 
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
      Left            =   3480
      TabIndex        =   4
      Top             =   2160
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPDOB 
      Height          =   345
      Left            =   3480
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.ComboBox cmbState 
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
      ItemData        =   "frmStudent.frx":20454
      Left            =   3480
      List            =   "frmStudent.frx":20464
      TabIndex        =   8
      Top             =   3960
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8280
      Top             =   6960
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8040
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8040
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8040
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8040
      Width           =   1800
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Delete"
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7320
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7320
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7320
      Width           =   1800
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00E0E0E0&
      Caption         =   "New"
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7320
      Width           =   1800
   End
   Begin VB.TextBox txtPhone 
      DataField       =   "phone no"
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
      Left            =   3480
      TabIndex        =   12
      Text            =   " "
      Top             =   5640
      Width           =   2745
   End
   Begin VB.TextBox txtEmail 
      DataField       =   "email id"
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
      Left            =   9600
      TabIndex        =   13
      Text            =   " "
      Top             =   5520
      Width           =   2745
   End
   Begin VB.TextBox txtMotherName 
      DataField       =   "mother name"
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
      Left            =   9600
      TabIndex        =   11
      Text            =   " "
      Top             =   4800
      Width           =   2745
   End
   Begin VB.TextBox txtFatherName 
      DataField       =   "father name"
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
      Left            =   3480
      TabIndex        =   10
      Text            =   " "
      Top             =   4800
      Width           =   2745
   End
   Begin VB.TextBox txtPIN 
      DataField       =   "pin"
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
      Left            =   9600
      TabIndex        =   9
      Text            =   " "
      Top             =   3960
      Width           =   2745
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
      Left            =   6240
      TabIndex        =   38
      Text            =   " "
      Top             =   3960
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txtCity 
      DataField       =   "city"
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
      Left            =   9600
      TabIndex        =   7
      Text            =   " "
      Top             =   3000
      Width           =   2745
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "address"
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
      Left            =   3480
      TabIndex        =   6
      Text            =   " "
      Top             =   3000
      Width           =   2745
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
      Left            =   12720
      TabIndex        =   37
      Text            =   " "
      Top             =   1320
      Visible         =   0   'False
      Width           =   465
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
      Left            =   6720
      TabIndex        =   36
      Text            =   " "
      Top             =   1320
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtName 
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
      Left            =   9600
      TabIndex        =   1
      Text            =   " "
      Top             =   480
      Width           =   2745
   End
   Begin VB.TextBox txtRollno 
      DataField       =   "rollno"
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
      Left            =   3480
      TabIndex        =   0
      Text            =   " "
      Top             =   480
      Width           =   2745
   End
   Begin MSComCtl2.DTPicker DTPADate 
      Height          =   345
      Left            =   3480
      TabIndex        =   53
      Top             =   6360
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Admission Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   52
      Top             =   6360
      Width           =   1980
   End
   Begin VB.Label Label7 
      Caption         =   "SF details"
      Height          =   255
      Left            =   240
      TabIndex        =   50
      Top             =   6360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Student"
      Height          =   255
      Left            =   6360
      TabIndex        =   49
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "course"
      Height          =   255
      Left            =   6960
      TabIndex        =   48
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "SFdetails"
      Height          =   372
      Left            =   240
      TabIndex        =   47
      Top             =   7200
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "course"
      Height          =   375
      Left            =   4200
      TabIndex        =   46
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Student"
      Height          =   255
      Left            =   7680
      TabIndex        =   45
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblCourseName 
      BackColor       =   &H00FFFFC0&
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
      Height          =   255
      Left            =   7800
      TabIndex        =   41
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblCourseId 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Course ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   39
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblFathername 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   35
      Top             =   4800
      Width           =   1500
   End
   Begin VB.Label lblMotherName 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mother Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   34
      Top             =   4800
      Width           =   1500
   End
   Begin VB.Label lblPhone 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   33
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lblNationality 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   32
      Top             =   6360
      Width           =   1500
   End
   Begin VB.Label lblEmail 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   31
      Top             =   5520
      Width           =   1500
   End
   Begin VB.Label lblPIN 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "PIN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   30
      Top             =   3960
      Width           =   1020
   End
   Begin VB.Label lblState 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   29
      Top             =   3960
      Width           =   1500
   End
   Begin VB.Label lblCity 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   28
      Top             =   3000
      Width           =   1500
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   27
      Top             =   3000
      Width           =   2340
   End
   Begin VB.Label lblGener 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   26
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label lblDOB 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of birth"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   25
      Top             =   1320
      Width           =   1620
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFFFC0&
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
      Height          =   255
      Left            =   7800
      TabIndex        =   24
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label lblRollno 
      BackColor       =   &H00FFFFC0&
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
      Height          =   255
      Left            =   1440
      TabIndex        =   23
      Top             =   480
      Width           =   1500
   End
End
Attribute VB_Name = "frmStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdDelete_Click()
Adodc1.Recordset.Delete
MsgBox ("Record Delete")
Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmdFirst_Click()
Adodc1.Recordset.MoveFirst
cmbGender.Text = txtGender.Text
cmbCourseid.Text = txtCID.Text

Adodc2.Recordset.MoveFirst
While cmbCourseid.Text <> txtCourseId.Text
Adodc2.Recordset.MoveNext
Wend

DTPDOB.Value = txtDOB.Text
DTPADate.Value = txtADate.Text
cmbState.Text = txtState.Text

 

End Sub

Private Sub cmdLast_Click()
Adodc1.Recordset.MoveLast
Adodc3.Recordset.MoveLast
cmbGender.Text = txtGender.Text
cmbCourseid.Text = txtCID.Text
Adodc2.Recordset.MoveFirst
While cmbCourseid.Text <> txtCourseId.Text
Adodc2.Recordset.MoveNext
Wend
DTPDOB.Value = txtDOB.Text
DTPADate.Value = txtADate.Text
cmbState.Text = txtState.Text

End Sub

Private Sub cmdNew_Click()
If Adodc1.Recordset.EOF = True Then
n = 100
Else
Adodc1.Recordset.MoveLast
n = Mid(Adodc1.Recordset.Fields(0), 2, 3)
End If
Adodc1.Recordset.AddNew
Adodc3.Recordset.AddNew
txtRollno.Text = "S" + CStr(n + 1)
txtName.SetFocus
cmbCourseid.Text = ""
cmbGender.Text = ""
txtCourseName.Text = ""
cmbState.Text = ""
End Sub

Private Sub cmdNext_Click()
Adodc1.Recordset.MoveNext
Adodc3.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
    Adodc1.Recordset.MoveLast
    Adodc3.Recordset.MoveLast
    MsgBox ("this is Last Record")
End If
cmbGender.Text = txtGender.Text
cmbCourseid.Text = txtCID.Text

Adodc2.Recordset.MoveFirst
While cmbCourseid.Text <> txtCourseId.Text
Adodc2.Recordset.MoveNext
Wend

DTPDOB.Value = txtDOB.Text
DTPADate.Value = txtADate.Text
cmbState.Text = txtState.Text

End Sub

Private Sub cmdPrevious_Click()
Adodc1.Recordset.MovePrevious
Adodc3.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
Adodc1.Recordset.MoveFirst
Adodc3.Recordset.MoveFirst
MsgBox ("this is first Record")
End If

cmbGender.Text = txtGender.Text
cmbCourseid.Text = txtCID.Text
Adodc2.Recordset.MoveFirst
While cmbCourseid.Text <> txtCourseId.Text
Adodc2.Recordset.MoveNext
Wend
DTPDOB.Value = txtDOB.Text
DTPADate.Value = txtADate.Text
cmbState.Text = txtState.Text

End Sub

Private Sub cmdSave_Click()
txtGender.Text = cmbGender.Text
txtCID.Text = cmbCourseid.Text

txtState.Text = cmbState.Text
txtDOB.Text = DTPDOB.Value
txtADate.Text = DTPADate.Value
Adodc1.Recordset.Update


txtSFDRollno.Text = txtRollno.Text
Adodc2.Recordset.MoveFirst
While cmbCourseid.Text <> Adodc2.Recordset.Fields(0)
Adodc2.Recordset.MoveNext
Wend
txtSFDCourseFees.Text = Adodc2.Recordset.Fields(3)

Adodc3.Recordset.Update
MsgBox ("Record Added")
End Sub

Private Sub cmdUpdate_Click()
txtGender.Text = cmbGender.Text
txtCID.Text = cmbCourseid.Text

txtState.Text = cmbState.Text
txtDOB.Text = DTPDOB.Value
txtADate.Text = DTPADate.Value
Adodc1.Recordset.Update

txtSFDRollno.Text = txtRollno.Text
Adodc2.Recordset.MoveFirst
While cmbCourseid.Text <> Adodc2.Recordset.Fields(0)
Adodc2.Recordset.MoveNext
Wend
txtSFDCourseFees.Text = Adodc2.Recordset.Fields(3)
Adodc3.Recordset.Update

MsgBox ("Record Updated")
End Sub

Private Sub cmbCourseid_Click()
Adodc2.Recordset.MoveFirst
While cmbCourseid.Text <> txtCourseId.Text
Adodc2.Recordset.MoveNext
Wend

txtCID.Text = cmbCourseid.Text
 
 
 End Sub


Private Sub Form_Load()

Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

Adodc2.Recordset.MoveFirst
While Adodc2.Recordset.EOF = False
cmbCourseid.AddItem (Adodc2.Recordset.Fields(0))
Adodc2.Recordset.MoveNext
Wend


cmbGender.Text = txtGender.Text
cmbCourseid.Text = txtCID.Text
Adodc2.Recordset.MoveFirst
While cmbCourseid.Text <> txtCourseId.Text
Adodc2.Recordset.MoveNext
Wend


DTPDOB.Value = txtDOB.Text
DTPADate.Value = txtADate.Text
cmbState.Text = txtState.Text

End Sub

