VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTeacher 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Teacher Details"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14745
   LinkTopic       =   "Form1"
   Picture         =   "frmTeacher.frx":0000
   ScaleHeight     =   9420
   ScaleWidth      =   14745
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
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
      Left            =   4680
      TabIndex        =   84
      Top             =   360
      Width           =   9615
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
         ItemData        =   "frmTeacher.frx":2EDBC
         Left            =   720
         List            =   "frmTeacher.frx":2EE68
         TabIndex        =   12
         Top             =   1080
         Width           =   3735
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
         ItemData        =   "frmTeacher.frx":2F705
         Left            =   7200
         List            =   "frmTeacher.frx":2F772
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
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
         TabIndex        =   15
         Top             =   1080
         Width           =   855
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
         TabIndex        =   13
         Top             =   1080
         Width           =   2535
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
         TabIndex        =   17
         Top             =   1560
         Width           =   2535
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
         TabIndex        =   19
         Top             =   1560
         Width           =   855
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
         ItemData        =   "frmTeacher.frx":2F844
         Left            =   7200
         List            =   "frmTeacher.frx":2F8B1
         TabIndex        =   18
         Top             =   1560
         Width           =   1095
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
         ItemData        =   "frmTeacher.frx":2F983
         Left            =   720
         List            =   "frmTeacher.frx":2FA2F
         TabIndex        =   16
         Top             =   1560
         Width           =   3735
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
         TabIndex        =   21
         Top             =   2040
         Width           =   2535
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
         TabIndex        =   23
         Top             =   2040
         Width           =   855
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
         ItemData        =   "frmTeacher.frx":302CC
         Left            =   7200
         List            =   "frmTeacher.frx":30339
         TabIndex        =   22
         Top             =   2040
         Width           =   1095
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
         ItemData        =   "frmTeacher.frx":3040B
         Left            =   720
         List            =   "frmTeacher.frx":304B7
         TabIndex        =   20
         Top             =   2040
         Width           =   3735
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
         TabIndex        =   25
         Top             =   2520
         Width           =   2535
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
         TabIndex        =   27
         Top             =   2520
         Width           =   855
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
         ItemData        =   "frmTeacher.frx":30D54
         Left            =   7200
         List            =   "frmTeacher.frx":30DC1
         TabIndex        =   26
         Top             =   2520
         Width           =   1095
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
         ItemData        =   "frmTeacher.frx":30E93
         Left            =   720
         List            =   "frmTeacher.frx":30F3F
         TabIndex        =   24
         Top             =   2520
         Width           =   3735
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
         TabIndex        =   29
         Top             =   3000
         Width           =   2535
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
         TabIndex        =   31
         Top             =   3000
         Width           =   855
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
         ItemData        =   "frmTeacher.frx":317DC
         Left            =   7200
         List            =   "frmTeacher.frx":31849
         TabIndex        =   30
         Top             =   3000
         Width           =   1095
      End
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
         ItemData        =   "frmTeacher.frx":3191B
         Left            =   720
         List            =   "frmTeacher.frx":319C7
         TabIndex        =   28
         Top             =   3000
         Width           =   3735
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   93
         Top             =   480
         Width           =   3045
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   92
         Top             =   480
         Width           =   1365
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   91
         Top             =   480
         Width           =   2445
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   90
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   89
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   88
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   87
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   86
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   85
         Top             =   3000
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
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
      Left            =   360
      TabIndex        =   52
      Top             =   6360
      Width           =   11535
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
         ItemData        =   "frmTeacher.frx":32264
         Left            =   3600
         List            =   "frmTeacher.frx":32271
         TabIndex        =   43
         Top             =   600
         Width           =   2550
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
         TabIndex        =   54
         Top             =   3120
         Width           =   2190
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
         Height          =   330
         Left            =   240
         TabIndex        =   42
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtDOJ 
         DataField       =   "dateofjoining"
         DataSource      =   "teacher"
         Height          =   405
         Left            =   10200
         TabIndex        =   53
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPDOJ 
         Height          =   345
         Left            =   8040
         TabIndex        =   44
         Top             =   600
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
         Format          =   92078081
         CurrentDate     =   43734
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
         TabIndex        =   58
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   57
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   56
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   55
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame fnrPersonal 
      BackColor       =   &H00FFFFC0&
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
      Left            =   360
      TabIndex        =   72
      Top             =   360
      Width           =   4320
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
         ItemData        =   "frmTeacher.frx":32295
         Left            =   1920
         List            =   "frmTeacher.frx":322A5
         TabIndex        =   7
         Top             =   3360
         Width           =   2055
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
         TabIndex        =   11
         Top             =   5400
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
         TabIndex        =   10
         Top             =   4800
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
         TabIndex        =   9
         Top             =   4320
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
         TabIndex        =   8
         Top             =   3840
         Width           =   2000
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
         TabIndex        =   5
         Top             =   2880
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
         TabIndex        =   4
         Top             =   2400
         Width           =   1950
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
         TabIndex        =   3
         Top             =   1920
         Width           =   2000
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
         TabIndex        =   1
         Top             =   960
         Width           =   2000
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
         Left            =   1920
         TabIndex        =   0
         Top             =   480
         Width           =   2000
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
         ItemData        =   "frmTeacher.frx":322D8
         Left            =   1920
         List            =   "frmTeacher.frx":322E2
         TabIndex        =   2
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblId 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   83
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   82
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lblGender 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   81
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   80
         Top             =   1920
         Width           =   1020
      End
      Begin VB.Label lblCity 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   79
         Top             =   2400
         Width           =   1620
      End
      Begin VB.Label lblState 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   78
         Top             =   3480
         Width           =   1260
      End
      Begin VB.Label lblPIN 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   77
         Top             =   3000
         Width           =   420
      End
      Begin VB.Label lblFatherName 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   76
         Top             =   3960
         Width           =   1380
      End
      Begin VB.Label lblPhone 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   75
         Top             =   4920
         Width           =   1500
      End
      Begin VB.Label lblEmail 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   74
         Top             =   4440
         Width           =   1380
      End
      Begin VB.Label lblNationality 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   73
         Top             =   5400
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00FFC0C0&
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   7920
      Width           =   1700
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0C0&
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   7920
      Width           =   1700
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFC0C0&
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7920
      Width           =   1700
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFC0C0&
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   7920
      Width           =   1700
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00FFC0C0&
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   8760
      Width           =   1700
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00FFC0C0&
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   8760
      Width           =   1700
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFC0C0&
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   8760
      Width           =   1700
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00FFC0C0&
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   8760
      Width           =   1700
   End
   Begin MSAdodcLib.Adodc teacher 
      Height          =   330
      Left            =   480
      Top             =   7920
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
   Begin VB.Frame Qualification 
      BackColor       =   &H00FFFFC0&
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
      Left            =   4680
      TabIndex        =   64
      Top             =   4080
      Width           =   7200
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
         TabIndex        =   37
         Top             =   1680
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
         ItemData        =   "frmTeacher.frx":322F4
         Left            =   1080
         List            =   "frmTeacher.frx":3230D
         TabIndex        =   36
         Top             =   1680
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
         TabIndex        =   35
         Top             =   1080
         Width           =   2415
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
         ItemData        =   "frmTeacher.frx":3234C
         Left            =   1080
         List            =   "frmTeacher.frx":32365
         TabIndex        =   34
         Top             =   1080
         Width           =   1215
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
         ItemData        =   "frmTeacher.frx":323A4
         Left            =   1080
         List            =   "frmTeacher.frx":323BD
         TabIndex        =   32
         Top             =   480
         Width           =   1215
      End
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
         TabIndex        =   33
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   71
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   70
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   69
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   68
         Top             =   1080
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
         Left            =   -3480
         TabIndex        =   67
         Top             =   -3960
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   66
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   65
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   59
      Top             =   4080
      Width           =   2415
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
         TabIndex        =   38
         Top             =   840
         Width           =   2070
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
         TabIndex        =   39
         Top             =   1560
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
         TabIndex        =   40
         Top             =   2280
         Width           =   2055
      End
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
         TabIndex        =   41
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFC0&
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
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   62
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   61
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   60
         Top             =   2880
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmTeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
MsgBox ("Record Deleted")
teacher.Recordset.Delete

End Sub

Private Sub cmdFirst_Click()
 DTPDOJ.Value = txtDOJ.Text
teacher.Recordset.MoveFirst

End Sub

Private Sub cmdLast_Click()
DTPDOJ.Value = txtDOJ.Text
teacher.Recordset.MoveLast

End Sub

Private Sub cmdNew_Click()
If teacher.Recordset.EOF = True Then
n = 100
Else
teacher.Recordset.MoveLast
n = Mid(teacher.Recordset.Fields(0), 4, 3)
End If
teacher.Recordset.AddNew
txtteacherId.Text = "Tch" + CStr(n + 1)
txtName.SetFocus

End Sub

Private Sub cmdNext_Click()
teacher.Recordset.MoveNext
If teacher.Recordset.EOF = True Then
teacher.Recordset.MoveFirst
MsgBox ("this is  last Record")
End If
DTPDOJ.Value = txtDOJ.Text

End Sub

Private Sub cmdPrevious_Click()
teacher.Recordset.MovePrevious
If teacher.Recordset.BOF = True Then
teacher.Recordset.MoveFirst
MsgBox ("this is first Record")
End If
DTPDOJ.Value = txtDOJ.Text
End Sub

Private Sub cmdSave_Click()
txtDOJ.Text = DTPDOJ.Value
teacher.Recordset.Update
MsgBox ("Record Added")
End Sub

Private Sub cmdUpdate_Click()
txtDOJ.Text = DTPDOJ.Value
teacher.Recordset.Update
MsgBox ("Record Updated")
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

