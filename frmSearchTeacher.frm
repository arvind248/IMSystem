VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSTeacher 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Search Teacher Details"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14520
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   14520
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
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
      Left            =   11880
      TabIndex        =   77
      Top             =   3960
      Width           =   2295
      Begin VB.TextBox txtPAN 
         DataField       =   "pannumber"
         DataSource      =   "teacher"
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
         TabIndex        =   81
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox txtIFSC 
         DataField       =   "ifsccode"
         DataSource      =   "teacher"
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
         TabIndex        =   80
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtANumber 
         DataField       =   "acccountnumber"
         DataSource      =   "teacher"
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
         TabIndex        =   79
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtAHName 
         DataField       =   "accoutholdername"
         DataSource      =   "teacher"
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
         TabIndex        =   78
         Top             =   840
         Width           =   2070
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   85
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   84
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   83
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   82
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Qualification 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Teaching Experience "
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
      Height          =   2295
      Left            =   4560
      TabIndex        =   63
      Top             =   3960
      Width           =   7200
      Begin VB.TextBox txtWComp 
         DataField       =   "1place"
         DataSource      =   "teacher"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4680
         TabIndex        =   69
         Top             =   480
         Width           =   2415
      End
      Begin VB.ComboBox cmbExperience 
         DataField       =   "1experience"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":0000
         Left            =   1080
         List            =   "frmSearchTeacher.frx":0019
         TabIndex        =   68
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox Combo11 
         DataField       =   "2experience"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":0058
         Left            =   1080
         List            =   "frmSearchTeacher.frx":0071
         TabIndex        =   67
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         DataField       =   "2palce"
         DataSource      =   "teacher"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4680
         TabIndex        =   66
         Top             =   1080
         Width           =   2415
      End
      Begin VB.ComboBox Combo12 
         DataField       =   "3experience"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":00B0
         Left            =   1080
         List            =   "frmSearchTeacher.frx":00C9
         TabIndex        =   65
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         DataField       =   "3place"
         DataSource      =   "teacher"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4680
         TabIndex        =   64
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "College/School/Institute Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   76
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Experience"
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
         Left            =   120
         TabIndex        =   75
         Top             =   480
         Width           =   975
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
         Left            =   -3480
         TabIndex        =   74
         Top             =   -3960
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Experience"
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
         Left            =   120
         TabIndex        =   73
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "College/School/Institute Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   72
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Experience"
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
         Left            =   120
         TabIndex        =   71
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "College/School/Institute Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   70
         Top             =   1680
         Width           =   2055
      End
   End
   Begin VB.Frame fnrPersonal 
      BackColor       =   &H00C0E0FF&
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
      Height          =   6015
      Left            =   240
      TabIndex        =   40
      Top             =   240
      Width           =   4320
      Begin VB.ComboBox cmbTeacherId 
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
         ItemData        =   "frmSearchTeacher.frx":0108
         Left            =   1920
         List            =   "frmSearchTeacher.frx":010A
         TabIndex        =   86
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox cmbGender 
         DataField       =   "gender"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":010C
         Left            =   1920
         List            =   "frmSearchTeacher.frx":0116
         TabIndex        =   51
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtteacherId 
         DataField       =   "teacherid"
         DataSource      =   "teacher"
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
         Left            =   1440
         TabIndex        =   50
         Top             =   480
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtName 
         DataField       =   "name"
         DataSource      =   "teacher"
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
         TabIndex        =   49
         Top             =   960
         Width           =   2000
      End
      Begin VB.TextBox txtAddress 
         DataField       =   "address"
         DataSource      =   "teacher"
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
         TabIndex        =   48
         Top             =   1920
         Width           =   2000
      End
      Begin VB.TextBox txtCity 
         DataField       =   "city"
         DataSource      =   "teacher"
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
         TabIndex        =   47
         Top             =   2400
         Width           =   1950
      End
      Begin VB.TextBox txtPIN 
         DataField       =   "pin"
         DataSource      =   "teacher"
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
         TabIndex        =   46
         Top             =   2880
         Width           =   2000
      End
      Begin VB.TextBox txtFatherName 
         DataField       =   "fathername"
         DataSource      =   "teacher"
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
         TabIndex        =   45
         Top             =   3840
         Width           =   2000
      End
      Begin VB.TextBox txtEmailId 
         DataField       =   "emailid"
         DataSource      =   "teacher"
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
         TabIndex        =   44
         Top             =   4320
         Width           =   2000
      End
      Begin VB.TextBox txtPhone 
         DataField       =   "phone"
         DataSource      =   "teacher"
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
         TabIndex        =   43
         Top             =   4800
         Width           =   2000
      End
      Begin VB.TextBox txtNationality 
         DataField       =   "nationality"
         DataSource      =   "teacher"
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
         TabIndex        =   42
         Top             =   5400
         Width           =   2000
      End
      Begin VB.ComboBox cmbState 
         DataField       =   "state"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":0128
         Left            =   1920
         List            =   "frmSearchTeacher.frx":0138
         TabIndex        =   41
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label lblNationality 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   62
         Top             =   5400
         Width           =   1500
      End
      Begin VB.Label lblEmail 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   61
         Top             =   4440
         Width           =   1380
      End
      Begin VB.Label lblPhone 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   60
         Top             =   4920
         Width           =   1500
      End
      Begin VB.Label lblFatherName 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   59
         Top             =   3960
         Width           =   1380
      End
      Begin VB.Label lblPIN 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   58
         Top             =   3000
         Width           =   420
      End
      Begin VB.Label lblState 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   57
         Top             =   3480
         Width           =   1260
      End
      Begin VB.Label lblCity 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   56
         Top             =   2400
         Width           =   1620
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   55
         Top             =   1920
         Width           =   1020
      End
      Begin VB.Label lblGender 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   54
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   53
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lblId 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Teacher Id"
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
         TabIndex        =   52
         Top             =   480
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   1335
      Left            =   240
      TabIndex        =   30
      Top             =   6240
      Width           =   11535
      Begin VB.TextBox txtDOJ 
         DataField       =   "dateofjoining"
         DataSource      =   "teacher"
         Height          =   405
         Left            =   10200
         TabIndex        =   34
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtBasicPay 
         DataField       =   "basicpay"
         DataSource      =   "teacher"
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
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Width           =   2055
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
         TabIndex        =   32
         Top             =   3120
         Width           =   2190
      End
      Begin VB.ComboBox cmbTeachingHours 
         DataField       =   "teachinghours"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":016B
         Left            =   3600
         List            =   "frmSearchTeacher.frx":0178
         TabIndex        =   31
         Top             =   600
         Width           =   2550
      End
      Begin MSComCtl2.DTPicker DTPDOJ 
         Height          =   345
         Left            =   8040
         TabIndex        =   35
         Top             =   720
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
         Format          =   92471297
         CurrentDate     =   43734
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Teaching Hours"
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
         Left            =   3600
         TabIndex        =   39
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Basic Pay"
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
         TabIndex        =   38
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Date of Joining"
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
         Left            =   8040
         TabIndex        =   37
         Top             =   480
         Width           =   1695
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
         TabIndex        =   36
         Top             =   3120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Qualification "
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
      Height          =   3735
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   9615
      Begin VB.ComboBox cmb5degree 
         DataField       =   "5degree"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":019C
         Left            =   720
         List            =   "frmSearchTeacher.frx":0248
         TabIndex        =   20
         Top             =   3000
         Width           =   3735
      End
      Begin VB.ComboBox cmb5passingyear 
         DataField       =   "5passingyear"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":0AE5
         Left            =   7200
         List            =   "frmSearchTeacher.frx":0B52
         TabIndex        =   19
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txt5marks 
         DataField       =   "5marks"
         DataSource      =   "teacher"
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
         Left            =   8520
         TabIndex        =   18
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txt5college 
         DataField       =   "5college"
         DataSource      =   "teacher"
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
         Left            =   4560
         TabIndex        =   17
         Top             =   3000
         Width           =   2535
      End
      Begin VB.ComboBox cmb4degree 
         DataField       =   "4degree"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":0C24
         Left            =   720
         List            =   "frmSearchTeacher.frx":0CD0
         TabIndex        =   16
         Top             =   2520
         Width           =   3735
      End
      Begin VB.ComboBox cmb4passingyear 
         DataField       =   "4passingyear"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":156D
         Left            =   7200
         List            =   "frmSearchTeacher.frx":15DA
         TabIndex        =   15
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txt4marks 
         DataField       =   "4marks"
         DataSource      =   "teacher"
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
         Left            =   8520
         TabIndex        =   14
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txt4college 
         DataField       =   "4college"
         DataSource      =   "teacher"
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
         Left            =   4560
         TabIndex        =   13
         Top             =   2520
         Width           =   2535
      End
      Begin VB.ComboBox cmb3degree 
         DataField       =   "3degree"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":16AC
         Left            =   720
         List            =   "frmSearchTeacher.frx":1758
         TabIndex        =   12
         Top             =   2040
         Width           =   3735
      End
      Begin VB.ComboBox cmb3passingyear 
         DataField       =   "3passingyear"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":1FF5
         Left            =   7200
         List            =   "frmSearchTeacher.frx":2062
         TabIndex        =   11
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txt3marks 
         DataField       =   "3marks"
         DataSource      =   "teacher"
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
         Left            =   8520
         TabIndex        =   10
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txt3college 
         DataField       =   "3college"
         DataSource      =   "teacher"
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
         Left            =   4560
         TabIndex        =   9
         Top             =   2040
         Width           =   2535
      End
      Begin VB.ComboBox cmb2degree 
         DataField       =   "2degree"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":2134
         Left            =   720
         List            =   "frmSearchTeacher.frx":21E0
         TabIndex        =   8
         Top             =   1560
         Width           =   3735
      End
      Begin VB.ComboBox cmb2passingyear 
         DataField       =   "2passingyear"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":2A7D
         Left            =   7200
         List            =   "frmSearchTeacher.frx":2AEA
         TabIndex        =   7
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txt2marks 
         DataField       =   "2marks"
         DataSource      =   "teacher"
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
         Left            =   8520
         TabIndex        =   6
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txt2college 
         DataField       =   "2college"
         DataSource      =   "teacher"
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
         Left            =   4560
         TabIndex        =   5
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txt1college 
         DataField       =   "1college"
         DataSource      =   "teacher"
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
         Left            =   4560
         TabIndex        =   4
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtmarks 
         DataField       =   "1marks"
         DataSource      =   "teacher"
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
         Left            =   8520
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox cmb1passingyear 
         DataField       =   "1passingyear"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":2BBC
         Left            =   7200
         List            =   "frmSearchTeacher.frx":2C29
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cmb1degree 
         DataField       =   "1degree"
         DataSource      =   "teacher"
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
         ItemData        =   "frmSearchTeacher.frx":2CFB
         Left            =   720
         List            =   "frmSearchTeacher.frx":2DA7
         TabIndex        =   1
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C0E0FF&
         Caption         =   "5."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0E0FF&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0E0FF&
         Caption         =   "3."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0E0FF&
         Caption         =   "2."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0E0FF&
         Caption         =   "1."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Marks"
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
         Left            =   8640
         TabIndex        =   24
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0E0FF&
         Caption         =   "School/College/Institute"
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
         Left            =   4680
         TabIndex        =   23
         Top             =   480
         Width           =   2445
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Passing year"
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
         Left            =   7200
         TabIndex        =   22
         Top             =   480
         Width           =   1365
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0E0FF&
         Caption         =   " Degrees/Diploma/Certificate"
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
         Left            =   720
         TabIndex        =   21
         Top             =   480
         Width           =   3045
      End
   End
   Begin MSAdodcLib.Adodc teacher 
      Height          =   330
      Left            =   120
      Top             =   7560
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
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:/project/db/IMS.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:/project/db/IMS.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "teacher"
      Caption         =   "teacher"
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
End
Attribute VB_Name = "frmSTeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
teacher.Recordset.MoveFirst
While teacher.Recordset.EOF = False
cmbTeacherId.AddItem (teacher.Recordset.Fields(0))
teacher.Recordset.MoveNext
Wend

End Sub
Private Sub cmbTeacherId_Click()
teacher.Recordset.MoveFirst
While cmbTeacherId.Text <> txtteacherId.Text
teacher.Recordset.MoveNext
Wend
End Sub
