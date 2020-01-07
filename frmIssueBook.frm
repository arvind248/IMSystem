VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIssueBook 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   13815
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   1560
      Top             =   6360
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Left            =   9720
      TabIndex        =   35
      Top             =   2160
      Width           =   2535
   End
   Begin VB.ComboBox cmbBookId 
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
      Left            =   720
      TabIndex        =   34
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
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
      Height          =   550
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6120
      Width           =   1500
   End
   Begin VB.CommandButton cmdIssue 
      BackColor       =   &H00FF8080&
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
      Height          =   550
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6120
      Width           =   1500
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   7800
      Top             =   5640
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      RecordSource    =   "student"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   5640
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      RecordSource    =   "book"
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Issue Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   480
      TabIndex        =   28
      Top             =   240
      Width           =   12855
      Begin VB.TextBox txtIssueId 
         DataField       =   "issueid"
         DataSource      =   "Adodc3"
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
         Left            =   480
         TabIndex        =   40
         Top             =   600
         Width           =   2500
      End
      Begin VB.TextBox txtDDate 
         DataField       =   "duedate"
         DataSource      =   "Adodc3"
         Height          =   375
         Left            =   11400
         TabIndex        =   39
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtIDate 
         DataField       =   "issuedate"
         DataSource      =   "Adodc3"
         Height          =   375
         Left            =   7440
         TabIndex        =   38
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComCtl2.DTPicker DTPDDate 
         Height          =   345
         Left            =   8880
         TabIndex        =   37
         Top             =   600
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
         Format          =   90767361
         CurrentDate     =   43745
      End
      Begin MSComCtl2.DTPicker DTPIDate 
         Height          =   345
         Left            =   4800
         TabIndex        =   36
         Top             =   600
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
         Format          =   90767361
         CurrentDate     =   43745
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
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
         Left            =   8880
         TabIndex        =   31
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
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
         Left            =   4800
         TabIndex        =   30
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label lblIssueId 
         BackColor       =   &H00FFC0C0&
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
         Left            =   480
         TabIndex        =   29
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Book Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   600
      TabIndex        =   13
      Top             =   1560
      Width           =   6975
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FFFFFF&
         DataField       =   "status"
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
         Left            =   3720
         TabIndex        =   43
         Text            =   " "
         Top             =   3480
         Width           =   2505
      End
      Begin VB.TextBox txtBID 
         DataField       =   "bookid"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   2760
         TabIndex        =   41
         Top             =   480
         Visible         =   0   'False
         Width           =   735
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
         Left            =   2880
         TabIndex        =   20
         Text            =   " "
         Top             =   840
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtBName 
         BackColor       =   &H00FFFFFF&
         DataField       =   "name"
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
         Left            =   3720
         TabIndex        =   19
         Text            =   " "
         Top             =   840
         Width           =   2500
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00FFFFFF&
         DataField       =   "title"
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
         Left            =   120
         TabIndex        =   18
         Text            =   " "
         Top             =   1680
         Width           =   2500
      End
      Begin VB.TextBox txtPublisher 
         BackColor       =   &H00FFFFFF&
         DataField       =   "publisher"
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
         Left            =   120
         TabIndex        =   17
         Text            =   " "
         Top             =   2520
         Width           =   2505
      End
      Begin VB.TextBox txtPrice 
         BackColor       =   &H00FFFFFF&
         DataField       =   "price"
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
         Left            =   3720
         TabIndex        =   16
         Text            =   " "
         Top             =   2520
         Width           =   2500
      End
      Begin VB.TextBox txtISBN 
         BackColor       =   &H00FFFFFF&
         DataField       =   "isbn"
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
         Left            =   120
         TabIndex        =   15
         Text            =   " "
         Top             =   3480
         Width           =   2500
      End
      Begin VB.TextBox txtAuthor 
         BackColor       =   &H00FFFFFF&
         DataField       =   "author"
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
         Left            =   3720
         TabIndex        =   14
         Text            =   " "
         Top             =   1680
         Width           =   2505
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Status"
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
         Left            =   3720
         TabIndex        =   44
         Top             =   3240
         Width           =   1500
      End
      Begin VB.Label lblBookId 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Book ID"
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
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label lblBName 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Name"
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
         Left            =   3720
         TabIndex        =   26
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Title"
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
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   1500
      End
      Begin VB.Label lbPublisher 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Publisher"
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
         Left            =   120
         TabIndex        =   24
         Top             =   2280
         Width           =   1500
      End
      Begin VB.Label lblAuthor 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Author"
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
         Left            =   3720
         TabIndex        =   23
         Top             =   1440
         Width           =   1500
      End
      Begin VB.Label lblPrice 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Price"
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
         Left            =   3720
         TabIndex        =   22
         Top             =   2280
         Width           =   1500
      End
      Begin VB.Label lblISBN 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ISBN no."
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
         Left            =   120
         TabIndex        =   21
         Top             =   3240
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Student Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   7800
      TabIndex        =   0
      Top             =   1440
      Width           =   5535
      Begin VB.TextBox txtRN 
         DataField       =   "rollno"
         DataSource      =   "Adodc3"
         Height          =   375
         Left            =   5160
         TabIndex        =   42
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtRollno 
         DataField       =   "rollno"
         DataSource      =   "Adodc2"
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
         Left            =   4560
         TabIndex        =   6
         Text            =   " "
         Top             =   720
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txtSName 
         DataField       =   "name"
         DataSource      =   "Adodc2"
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
         TabIndex        =   5
         Text            =   " "
         Top             =   1200
         Width           =   2500
      End
      Begin VB.TextBox txtDOB 
         DataField       =   "dob"
         DataSource      =   "Adodc2"
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
         TabIndex        =   4
         Text            =   " "
         Top             =   1800
         Width           =   2505
      End
      Begin VB.TextBox txtGender 
         DataField       =   "gender"
         DataSource      =   "Adodc2"
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
         Left            =   1920
         TabIndex        =   3
         Text            =   " "
         Top             =   2280
         Width           =   2505
      End
      Begin VB.TextBox txtEmail 
         DataField       =   "email id"
         DataSource      =   "Adodc2"
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
         Left            =   1920
         TabIndex        =   2
         Text            =   " "
         Top             =   3360
         Width           =   2500
      End
      Begin VB.TextBox txtPhone 
         DataField       =   "phone no"
         DataSource      =   "Adodc2"
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
         Left            =   1920
         TabIndex        =   1
         Text            =   " "
         Top             =   2880
         Width           =   2500
      End
      Begin VB.Label lblRollno 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rollno"
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
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label lbSlName 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Name"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label lblDOB 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Date of birth"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label lblGener 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   1500
      End
      Begin VB.Label lblEmail 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Email ID"
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
         Left            =   240
         TabIndex        =   8
         Top             =   3360
         Width           =   1500
      End
      Begin VB.Label lblPhone 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Phone"
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
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmIssueBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbBookId_Click()
Adodc1.Recordset.MoveFirst
While cmbBookId.Text <> txtBookId.Text
Adodc1.Recordset.MoveNext
Wend

End Sub

Private Sub cmbRollno_Click()
Adodc2.Recordset.MoveFirst
While cmbRollno.Text <> txtRollno.Text
Adodc2.Recordset.MoveNext
Wend

End Sub
Private Sub cmdIssue_Click()

If txtStatus = "ISSUED" Then
MsgBox ("Sorry! This book is ISSUED")
Adodc1.Recordset.MoveLast
Adodc1.Recordset.MoveNext
cmbBookId.Text = txtBookId.Text
Adodc2.Recordset.MoveLast
Adodc2.Recordset.MoveNext
cmbRollno.Text = txtRollno.Text
Else

 txtIDate.Text = DTPIDate.Value
 txtDDate.Text = DTPDDate.Value
 txtBID.Text = txtBookId
 txtRN.Text = txtRollno.Text
 
Adodc3.Recordset.Update
txtStatus.Text = "ISSUED"
Adodc1.Recordset.Update
MsgBox ("Record added")
Adodc3.Recordset.MoveLast
n = Mid(Adodc3.Recordset.Fields(0), 4, 3)
Adodc3.Recordset.AddNew
txtIssueId.Text = "ISB" + CStr(n + 1)
Adodc1.Recordset.MoveLast
Adodc1.Recordset.MoveNext
cmbBookId.Text = txtBookId.Text
Adodc2.Recordset.MoveLast
Adodc2.Recordset.MoveNext
cmbRollno.Text = txtRollno.Text
End If
End Sub

Private Sub Command2_Click()
Adodc3.Recordset.MoveLast
Adodc3.Recordset.Delete
Unload Me


End Sub

Private Sub Form_Load()

Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Adodc1.Recordset.MoveFirst
While Adodc2.Recordset.EOF = False
cmbRollno.AddItem (Adodc2.Recordset.Fields(0))
Adodc2.Recordset.MoveNext
Wend

Adodc1.Recordset.MoveFirst
While Adodc1.Recordset.EOF = False
cmbBookId.AddItem (Adodc1.Recordset.Fields(0))
Adodc1.Recordset.MoveNext
Wend

If Adodc3.Recordset.EOF = True Then
n = 100
Else
Adodc3.Recordset.MoveLast
n = Mid(Adodc3.Recordset.Fields(0), 4, 3)
End If
Adodc3.Recordset.AddNew
txtIssueId.Text = "ISB" + CStr(n + 1)



End Sub

