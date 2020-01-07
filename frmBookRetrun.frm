VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBookReturn 
   BackColor       =   &H0080FF80&
   Caption         =   "Book Return"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14955
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   14955
   Begin MSAdodcLib.Adodc BookReturn 
      Height          =   330
      Left            =   2880
      Top             =   6600
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "bookreturn"
      Caption         =   "Book return"
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
   Begin MSAdodcLib.Adodc Book 
      Height          =   330
      Left            =   9960
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      RecordSource    =   "book"
      Caption         =   "Book"
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
   Begin MSAdodcLib.Adodc Student 
      Height          =   330
      Left            =   480
      Top             =   5880
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.TextBox txtIssuedDate 
      BackColor       =   &H00FFFFFF&
      DataField       =   "issuedate"
      DataSource      =   "BookIssued"
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
      Left            =   7920
      TabIndex        =   46
      Top             =   4800
      Width           =   2505
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
      Left            =   7920
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
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
      Left            =   2400
      TabIndex        =   0
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
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
      Left            =   480
      TabIndex        =   28
      Top             =   1680
      Width           =   6975
      Begin VB.TextBox txtPhone 
         DataField       =   "phone no"
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
         Left            =   1920
         TabIndex        =   35
         Top             =   2880
         Width           =   2500
      End
      Begin VB.TextBox txtEmail 
         DataField       =   "email id"
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
         Left            =   1920
         TabIndex        =   34
         Top             =   3360
         Width           =   2500
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
         Left            =   1920
         TabIndex        =   33
         Top             =   2280
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
         Left            =   1920
         TabIndex        =   32
         Top             =   1800
         Width           =   2505
      End
      Begin VB.TextBox txtSName 
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
         Left            =   1920
         TabIndex        =   31
         Top             =   1200
         Width           =   2500
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
         Left            =   4560
         TabIndex        =   30
         Text            =   " "
         Top             =   720
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txtRN 
         DataField       =   "rollno"
         DataSource      =   "BookReturn"
         Height          =   645
         Left            =   5160
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPhone 
         BackColor       =   &H0080FF80&
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
         TabIndex        =   41
         Top             =   2880
         Width           =   1500
      End
      Begin VB.Label lblEmail 
         BackColor       =   &H0080FF80&
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
         TabIndex        =   40
         Top             =   3360
         Width           =   1500
      End
      Begin VB.Label lblGener 
         BackColor       =   &H0080FF80&
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
         TabIndex        =   39
         Top             =   2280
         Width           =   1500
      End
      Begin VB.Label lblDOB 
         BackColor       =   &H0080FF80&
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
         TabIndex        =   38
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label lbSlName 
         BackColor       =   &H0080FF80&
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
         TabIndex        =   37
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label lblRollno 
         BackColor       =   &H0080FF80&
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
         TabIndex        =   36
         Top             =   720
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
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
      Height          =   5295
      Left            =   7560
      TabIndex        =   10
      Top             =   480
      Width           =   6975
      Begin VB.TextBox txtIDate 
         DataField       =   "issueddate"
         DataSource      =   "BookReturn"
         Height          =   525
         Left            =   1560
         TabIndex        =   47
         Top             =   4800
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSAdodcLib.Adodc BookIssued 
         Height          =   375
         Left            =   3000
         Top             =   4320
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   "issuebook"
         Caption         =   "Book Issued"
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
      Begin VB.TextBox txtAuthor 
         BackColor       =   &H00FFFFFF&
         DataField       =   "author"
         DataSource      =   "Book"
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
         Left            =   3960
         TabIndex        =   19
         Top             =   1560
         Width           =   2505
      End
      Begin VB.TextBox txtISBN 
         BackColor       =   &H00FFFFFF&
         DataField       =   "isbn"
         DataSource      =   "Book"
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
         Left            =   360
         TabIndex        =   18
         Top             =   3360
         Width           =   2500
      End
      Begin VB.TextBox txtPrice 
         BackColor       =   &H00FFFFFF&
         DataField       =   "price"
         DataSource      =   "Book"
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
         Left            =   3960
         TabIndex        =   17
         Top             =   2400
         Width           =   2500
      End
      Begin VB.TextBox txtPublisher 
         BackColor       =   &H00FFFFFF&
         DataField       =   "publisher"
         DataSource      =   "Book"
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
         Left            =   360
         TabIndex        =   16
         Top             =   2400
         Width           =   2505
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00FFFFFF&
         DataField       =   "title"
         DataSource      =   "Book"
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
         Left            =   360
         TabIndex        =   15
         Top             =   1560
         Width           =   2500
      End
      Begin VB.TextBox txtBName 
         BackColor       =   &H00FFFFFF&
         DataField       =   "name"
         DataSource      =   "Book"
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
         Left            =   3960
         TabIndex        =   14
         Top             =   720
         Width           =   2500
      End
      Begin VB.TextBox txtBookId 
         BackColor       =   &H00FFFFFF&
         DataField       =   "bookid"
         DataSource      =   "Book"
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
         Left            =   3120
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtBID 
         DataField       =   "bookid"
         DataSource      =   "BookReturn"
         Height          =   525
         Left            =   3000
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FFFFFF&
         DataField       =   "status"
         DataSource      =   "Book"
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
         Left            =   3960
         TabIndex        =   11
         Top             =   3360
         Width           =   2505
      End
      Begin VB.Label lblIssuedDate 
         BackColor       =   &H0080FF80&
         Caption         =   "Issued Date"
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
         Left            =   360
         TabIndex        =   45
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label lblISBN 
         BackColor       =   &H0080FF80&
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
         Left            =   360
         TabIndex        =   27
         Top             =   3120
         Width           =   1260
      End
      Begin VB.Label lblPrice 
         BackColor       =   &H0080FF80&
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
         Left            =   3960
         TabIndex        =   26
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label lblAuthor 
         BackColor       =   &H0080FF80&
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
         Left            =   3960
         TabIndex        =   25
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label lbPublisher 
         BackColor       =   &H0080FF80&
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
         Left            =   360
         TabIndex        =   24
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H0080FF80&
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
         Left            =   360
         TabIndex        =   23
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label lblBName 
         BackColor       =   &H0080FF80&
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
         Left            =   3960
         TabIndex        =   22
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label lblBookId 
         BackColor       =   &H0080FF80&
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
         Left            =   360
         TabIndex        =   21
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H0080FF80&
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
         Left            =   3960
         TabIndex        =   20
         Top             =   3120
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FF80&
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
      TabIndex        =   4
      Top             =   480
      Width           =   6975
      Begin VB.TextBox txtDDate 
         DataField       =   "duedate"
         DataSource      =   "Adodc3"
         Height          =   375
         Left            =   11400
         TabIndex        =   42
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtRDate 
         DataField       =   "returndate"
         DataSource      =   "BookReturn"
         Height          =   525
         Left            =   6480
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtReturnId 
         DataField       =   "returnid"
         DataSource      =   "BookReturn"
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
         TabIndex        =   5
         Top             =   600
         Width           =   2500
      End
      Begin MSComCtl2.DTPicker DTPRDate 
         Height          =   345
         Left            =   3840
         TabIndex        =   7
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
         Format          =   9306113
         CurrentDate     =   43745
      End
      Begin MSComCtl2.DTPicker DTPDDate 
         Height          =   345
         Left            =   8880
         TabIndex        =   43
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
         Format          =   9306113
         CurrentDate     =   43745
      End
      Begin VB.Label Label3 
         BackColor       =   &H000040C0&
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
         TabIndex        =   44
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblReturnId 
         BackColor       =   &H0080FF80&
         Caption         =   "Return ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "Return Date"
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
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Return "
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFC0&
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   1500
   End
End
Attribute VB_Name = "frmBookReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbBookId_Click()
Book.Recordset.MoveFirst
While cmbBookId.Text <> txtBookId.Text
Book.Recordset.MoveNext
Wend

BookIssued.Recordset.MoveFirst
While cmbBookId.Text <> BookIssued.Recordset.Fields(3)
BookIssued.Recordset.MoveNext
Wend
End Sub


Private Sub cmbRollno_Click()
Student.Recordset.MoveFirst
While cmbRollno.Text <> txtRollno.Text
Student.Recordset.MoveNext
Wend

End Sub

Private Sub cmdCancel_Click()
BookReturn.Recordset.MoveLast
BookReturn.Recordset.Delete
Unload Me

End Sub

Private Sub cmdReturn_Click()

 txtRDate.Text = DTPRDate.Value
 txtIDate.Text = txtIssuedDate.Text
 txtBID.Text = txtBookId
 txtRN.Text = txtRollno.Text
 
BookReturn.Recordset.Update
txtStatus.Text = "AVAILABLE"
Book.Recordset.Update
MsgBox ("Record added")
BookReturn.Recordset.MoveLast
n = Mid(BookReturn.Recordset.Fields(0), 3, 3)
BookReturn.Recordset.AddNew
txtReturnId.Text = "RB" + CStr(n + 1)
Student.Recordset.MoveLast
Student.Recordset.MoveNext
cmbRollno.Text = txtRollno.Text
Book.Recordset.MoveLast
Book.Recordset.MoveNext
cmbBookId.Text = txtBookId.Text

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Student.Recordset.MoveFirst
While Student.Recordset.EOF = False
cmbRollno.AddItem (Student.Recordset.Fields(0))
Student.Recordset.MoveNext
Wend

Book.Recordset.MoveFirst
While Book.Recordset.EOF = False
cmbBookId.AddItem (Book.Recordset.Fields(0))
Book.Recordset.MoveNext
Wend

If BookReturn.Recordset.EOF = True Then
n = 100
Else
BookReturn.Recordset.MoveLast
n = Mid(BookReturn.Recordset.Fields(0), 3, 3)
End If
BookReturn.Recordset.AddNew
txtReturnId.Text = "RB" + CStr(n + 1)

End Sub
