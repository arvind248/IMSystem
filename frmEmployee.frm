VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmployee 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Employee Details"
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   12165
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   6840
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      RecordSource    =   "employee"
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
      Height          =   500
      Left            =   8400
      TabIndex        =   38
      Top             =   8040
      Width           =   1700
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
      Height          =   500
      Left            =   6480
      TabIndex        =   37
      Top             =   8040
      Width           =   1700
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
      Height          =   500
      Left            =   4560
      TabIndex        =   36
      Top             =   8040
      Width           =   1700
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
      Height          =   500
      Left            =   2640
      TabIndex        =   35
      Top             =   8040
      Width           =   1700
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
      Height          =   500
      Left            =   8400
      TabIndex        =   34
      Top             =   7200
      Width           =   1700
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
      Height          =   500
      Left            =   6480
      TabIndex        =   33
      Top             =   7200
      Width           =   1700
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
      Height          =   500
      Left            =   4560
      TabIndex        =   32
      Top             =   7200
      Width           =   1700
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
      Height          =   500
      Left            =   2640
      TabIndex        =   31
      Top             =   7200
      Width           =   1700
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Imployement Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   3615
      Left            =   4680
      TabIndex        =   68
      Top             =   3120
      Width           =   4815
      Begin VB.TextBox txtDOJ 
         DataField       =   "doj"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   4320
         TabIndex        =   82
         Top             =   2640
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtEType 
         DataField       =   "emptype"
         DataSource      =   "Adodc1"
         Height          =   350
         Left            =   4320
         TabIndex        =   69
         Top             =   840
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txtDepartment 
         DataField       =   "department"
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
         Left            =   2040
         TabIndex        =   22
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtEManager 
         DataField       =   "empmanager"
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
         Left            =   2040
         TabIndex        =   23
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtPosition 
         DataField       =   "empposition"
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
         Left            =   2040
         TabIndex        =   20
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtSalary 
         DataField       =   "salary"
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
         Left            =   2040
         TabIndex        =   24
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtWHours 
         DataField       =   "workinghours"
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
         Left            =   2040
         TabIndex        =   26
         Top             =   3120
         Width           =   2190
      End
      Begin VB.ComboBox cmbEType 
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
         ItemData        =   "frmEmployee.frx":0000
         Left            =   2040
         List            =   "frmEmployee.frx":000D
         TabIndex        =   21
         Top             =   840
         Width           =   2190
      End
      Begin MSComCtl2.DTPicker DTPDOJ 
         Height          =   345
         Left            =   2040
         TabIndex        =   25
         Top             =   2640
         Width           =   2175
         _ExtentX        =   3836
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
         Format          =   92405761
         CurrentDate     =   43734
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Employment Type"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   76
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   75
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Employee  Manager "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   74
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Salary"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   73
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Employee Position"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Date of Joining"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   71
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Working Hours"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   70
         Top             =   3120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Bank Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   3615
      Left            =   9600
      TabIndex        =   62
      Top             =   3120
      Width           =   2295
      Begin VB.TextBox txtPAN 
         DataField       =   "pannumber"
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
         Left            =   120
         TabIndex        =   30
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox txtIFSC 
         DataField       =   "ifsccode"
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
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtANumber 
         DataField       =   "accountnumber"
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
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtAHName 
         DataField       =   "accountholdername"
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
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   2070
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0FFC0&
         Caption         =   "PAN Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "IFSC Code"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Account Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Account Holder Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   63
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Qualification 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Qualification"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   2535
      Left            =   4680
      TabIndex        =   53
      Top             =   480
      Width           =   7200
      Begin VB.TextBox txtPassingYear 
         DataField       =   "passingyear"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   6840
         TabIndex        =   81
         Top             =   960
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox txtIDuration 
         DataField       =   "internshipduration"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1800
         TabIndex        =   80
         Top             =   1920
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtDuration 
         DataField       =   "duration"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1800
         TabIndex        =   79
         Top             =   1440
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox txtDiploma 
         DataField       =   "diploma"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1320
         TabIndex        =   78
         Top             =   1440
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txtQualification 
         DataField       =   "highestqualification"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   6480
         TabIndex        =   77
         Top             =   960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtExperience 
         DataField       =   "workexperience"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1800
         TabIndex        =   67
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtIComp 
         DataField       =   "internshipcpmpanyname"
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
         Left            =   4800
         TabIndex        =   19
         Top             =   1920
         Width           =   1690
      End
      Begin VB.ComboBox cmbIDuration 
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
         ItemData        =   "frmEmployee.frx":0031
         Left            =   2160
         List            =   "frmEmployee.frx":003E
         TabIndex        =   18
         Top             =   1920
         Width           =   1215
      End
      Begin VB.ComboBox cmbDuration 
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
         ItemData        =   "frmEmployee.frx":005E
         Left            =   5520
         List            =   "frmEmployee.frx":0071
         TabIndex        =   17
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cmbDiploma 
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
         ItemData        =   "frmEmployee.frx":00A1
         Left            =   2160
         List            =   "frmEmployee.frx":00BD
         TabIndex        =   16
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtWComp 
         DataField       =   "workedcompanyname"
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
         Left            =   4800
         TabIndex        =   13
         Top             =   480
         Width           =   1690
      End
      Begin VB.ComboBox cmbPassingYear 
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
         ItemData        =   "frmEmployee.frx":0134
         Left            =   5520
         List            =   "frmEmployee.frx":0159
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cmbQualification 
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
         ItemData        =   "frmEmployee.frx":019F
         Left            =   2160
         List            =   "frmEmployee.frx":01AC
         TabIndex        =   14
         Top             =   960
         Width           =   2055
      End
      Begin VB.ComboBox cmbExperience 
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
         ItemData        =   "frmEmployee.frx":01D3
         Left            =   2160
         List            =   "frmEmployee.frx":01EC
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   61
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Duration"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   60
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Passing Year"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   59
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   58
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Work Experience"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   57
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Internship"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   56
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Diploma"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   55
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Highest Qualification"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   54
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.Frame fnrPersonal 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Personal Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   6255
      Left            =   360
      TabIndex        =   39
      Top             =   480
      Width           =   4320
      Begin VB.ComboBox cmbState 
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
         ItemData        =   "frmEmployee.frx":022B
         Left            =   1920
         List            =   "frmEmployee.frx":023B
         TabIndex        =   7
         Top             =   3360
         Width           =   2055
      End
      Begin VB.OptionButton optFemale 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton OptMale 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   2
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtNationality 
         DataField       =   "nationality"
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
         Left            =   1920
         TabIndex        =   11
         Top             =   5400
         Width           =   2000
      End
      Begin VB.TextBox txtPhone 
         DataField       =   "phone"
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
         Left            =   1920
         TabIndex        =   10
         Top             =   4800
         Width           =   2000
      End
      Begin VB.TextBox txtEmailId 
         DataField       =   "emailid"
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
         Left            =   1920
         TabIndex        =   9
         Top             =   4320
         Width           =   2000
      End
      Begin VB.TextBox txtFatherName 
         DataField       =   "fathename"
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
         Left            =   1920
         TabIndex        =   8
         Top             =   3840
         Width           =   2000
      End
      Begin VB.TextBox txtPIN 
         DataField       =   "pin"
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
         Left            =   1920
         TabIndex        =   6
         Top             =   2880
         Width           =   2000
      End
      Begin VB.TextBox txtState 
         DataField       =   "state"
         DataSource      =   "Adodc1"
         Height          =   350
         Left            =   960
         TabIndex        =   52
         Top             =   3360
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtCity 
         DataField       =   "city"
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
         Left            =   1920
         TabIndex        =   5
         Top             =   2400
         Width           =   1950
      End
      Begin VB.TextBox txtAddress 
         DataField       =   "address"
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
         Left            =   1920
         TabIndex        =   4
         Top             =   1920
         Width           =   2000
      End
      Begin VB.TextBox txtGender 
         DataField       =   "gender"
         DataSource      =   "Adodc1"
         Height          =   350
         Left            =   1200
         TabIndex        =   51
         Top             =   1440
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtName 
         DataField       =   "name"
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
         Left            =   1920
         TabIndex        =   1
         Top             =   960
         Width           =   2000
      End
      Begin VB.TextBox txtEmpId 
         DataField       =   "empid"
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
         Left            =   1920
         TabIndex        =   0
         Top             =   480
         Width           =   2000
      End
      Begin VB.Label lblEmpId 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Employee ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lblGender 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   1920
         Width           =   1020
      End
      Begin VB.Label lblCity 
         BackColor       =   &H00C0FFC0&
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   2400
         Width           =   1620
      End
      Begin VB.Label lblState 
         BackColor       =   &H00C0FFC0&
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   3480
         Width           =   1260
      End
      Begin VB.Label lblPIN 
         BackColor       =   &H00C0FFC0&
         Caption         =   "PIN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   3000
         Width           =   420
      End
      Begin VB.Label lblFatherName 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Father name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   3960
         Width           =   1380
      End
      Begin VB.Label lblPhone 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   4920
         Width           =   1500
      End
      Begin VB.Label lblEmail 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Email ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   4440
         Width           =   1380
      End
      Begin VB.Label lblNationality 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Nationality"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   5400
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmEmployee"
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
If txtGender.Text = "Male" Then
OptMale.SetFocus
End If
If txtGender.Text = "Female" Then
optFemale.SetFocus
End If
cmbState.Text = txtState.Text

cmbExperience.Text = txtExperience.Text
cmbQualification.Text = txtQualification.Text
cmbPassingYear.Text = txtPassingYear.Text
cmbDiploma.Text = txtDiploma.Text
cmbDuration.Text = txtDuration.Text
cmbIDuration.Text = txtIDuration.Text

cmbEType.Text = txtEType.Text
DTPDOJ.Value = txtDOJ.Text
 


End Sub

Private Sub cmdLast_Click()
Adodc1.Recordset.MoveLast

If txtGender.Text = "Male" Then
OptMale.SetFocus
End If
If txtGender.Text = "Female" Then
optFemale.SetFocus
End If
cmbState.Text = txtState.Text

cmbExperience.Text = txtExperience.Text
cmbQualification.Text = txtQualification.Text
cmbPassingYear.Text = txtPassingYear.Text
cmbDiploma.Text = txtDiploma.Text
cmbDuration.Text = txtDuration.Text
cmbIDuration.Text = txtIDuration.Text

cmbEType.Text = txtEType.Text
DTPDOJ.Value = txtDOJ.Text

End Sub

Private Sub cmdNew_Click()
If Adodc1.Recordset.EOF = True Then
n = 100
Else
Adodc1.Recordset.MoveLast
n = Mid(Adodc1.Recordset.Fields(0), 4, 3)
End If
Adodc1.Recordset.AddNew
txtEmpId.Text = "Emp" + CStr(n + 1)
txtName.SetFocus

End Sub

Private Sub cmdNext_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
        Adodc1.Recordset.MoveLast
        MsgBox ("this is Last Record")

End If

If txtGender.Text = "Male" Then
OptMale.SetFocus
End If
If txtGender.Text = "Female" Then
optFemale.SetFocus
End If
cmbState.Text = txtState.Text

cmbExperience.Text = txtExperience.Text
cmbQualification.Text = txtQualification.Text
cmbPassingYear.Text = txtPassingYear.Text
cmbDiploma.Text = txtDiploma.Text
cmbDuration.Text = txtDuration.Text
cmbIDuration.Text = txtIDuration.Text

cmbEType.Text = txtEType.Text
DTPDOJ.Value = txtDOJ.Text


End Sub

Private Sub cmdPrevious_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
    
    Adodc1.Recordset.MoveFirst
    MsgBox ("this is First Record")
End If

If txtGender.Text = "Male" Then
OptMale.SetFocus
End If
If txtGender.Text = "Female" Then
optFemale.SetFocus
End If
cmbState.Text = txtState.Text

cmbExperience.Text = txtExperience.Text
cmbQualification.Text = txtQualification.Text
cmbPassingYear.Text = txtPassingYear.Text
cmbDiploma.Text = txtDiploma.Text
cmbDuration.Text = txtDuration.Text
cmbIDuration.Text = txtIDuration.Text

cmbEType.Text = txtEType.Text
DTPDOJ.Value = txtDOJ.Text

End Sub

Private Sub cmdSave_Click()
txtState.Text = cmbState.Text
If OptMale.Value = True Then
txtGender.Text = "Male"
End If
If optFemale.Value = True Then
txtGender.Text = "Female"
End If

txtExperience.Text = cmbExperience.Text
txtQualification.Text = cmbQualification.Text
txtPassingYear.Text = cmbPassingYear.Text
txtDiploma.Text = cmbDiploma.Text
txtDuration.Text = cmbDuration.Text
txtIDuration.Text = cmbIDuration.Text

txtEType.Text = cmbEType.Text
txtDOJ.Text = DTPDOJ.Value


Adodc1.Recordset.Update
MsgBox ("Record Added")

End Sub

Private Sub cmdUpdate_Click()
txtState.Text = cmbState.Text
If OptMale.Value = True Then
txtGender.Text = "Male"
End If
If optFemale.Value = True Then
txtGender.Text = "Female"
End If

txtExperience.Text = cmbExperience.Text
txtQualification.Text = cmbQualification.Text
txtPassingYear.Text = cmbPassingYear.Text
txtDiploma.Text = cmbDiploma.Text
txtDuration.Text = cmbDuration.Text
txtIDuration.Text = cmbIDuration.Text

txtEType.Text = cmbEType.Text
txtDOJ.Text = DTPDOJ.Value


Adodc1.Recordset.Update
MsgBox ("Record Updated")

End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

cmbState.Text = txtState.Text

cmbExperience.Text = txtExperience.Text
cmbQualification.Text = txtQualification.Text
cmbPassingYear.Text = txtPassingYear.Text
cmbDiploma.Text = txtDiploma.Text
cmbDuration.Text = txtDuration.Text
cmbIDuration.Text = txtIDuration.Text

cmbEType.Text = txtEType.Text
DTPDOJ.Value = txtDOJ.Text



End Sub
