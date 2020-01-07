VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDI Form"
   ClientHeight    =   5910
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13245
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuMaster 
      Caption         =   "&Master"
      Index           =   1
      Begin VB.Menu mnuStudent 
         Caption         =   "Student"
         Index           =   12
      End
      Begin VB.Menu mnuEmployee 
         Caption         =   "Employee"
         Index           =   13
      End
      Begin VB.Menu mnuTeacher 
         Caption         =   "Teacher"
         Index           =   11
      End
      Begin VB.Menu mnuBook 
         Caption         =   "Book"
         Index           =   14
      End
      Begin VB.Menu mnuSubject 
         Caption         =   "Subject"
         Index           =   16
      End
      Begin VB.Menu mnucourse 
         Caption         =   "Course"
         Index           =   15
      End
      Begin VB.Menu mnuBatch 
         Caption         =   "Batch"
         Index           =   17
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "&Transaction"
      Index           =   2
      Begin VB.Menu mnuEnquiry 
         Caption         =   "Enquiry"
         Index           =   25
      End
      Begin VB.Menu mnu 
         Caption         =   "Book Issue"
         Index           =   21
      End
      Begin VB.Menu mnuBookReturn 
         Caption         =   "Book Return"
         Index           =   22
      End
      Begin VB.Menu mnuFees 
         Caption         =   "Pay Fees"
         Index           =   24
      End
      Begin VB.Menu mnuStudentAttendance 
         Caption         =   "Student Attendance"
         Index           =   26
      End
      Begin VB.Menu mnuTeaherAttendance 
         Caption         =   "Teacher Attendance"
         Index           =   27
      End
      Begin VB.Menu mnuEmployeeAttendance 
         Caption         =   "Employee Attendance"
         Index           =   28
      End
      Begin VB.Menu mnuTest 
         Caption         =   "Test"
         Index           =   28
      End
      Begin VB.Menu mnuStudentMarks 
         Caption         =   "Student Marks"
         Index           =   29
      End
      Begin VB.Menu mnuBatchAllotment 
         Caption         =   "Batch Allotmnet"
         Index           =   210
      End
      Begin VB.Menu mnuCompletionCertificate 
         Caption         =   "Completion Certificate "
         Index           =   211
      End
      Begin VB.Menu mnuPlacement 
         Caption         =   "Placement "
         Index           =   212
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Index           =   8
      Begin VB.Menu mnuSearchStudent 
         Caption         =   "Student"
         Index           =   31
      End
      Begin VB.Menu mnuSearchEmployee 
         Caption         =   "Employee"
         Index           =   32
      End
      Begin VB.Menu mnuSearchTeacher 
         Caption         =   "Teacher"
         Index           =   36
      End
      Begin VB.Menu mnusearchBook 
         Caption         =   "Book"
         Index           =   33
      End
      Begin VB.Menu mnuSearchSubject 
         Caption         =   "Subject"
         Index           =   35
      End
      Begin VB.Menu mnuSearchCourse 
         Caption         =   "Course"
         Index           =   34
      End
      Begin VB.Menu mnuSearchBatch 
         Caption         =   "Batch"
         Index           =   37
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Index           =   3
      Begin VB.Menu mnuSudentReport 
         Caption         =   "Student"
         Index           =   41
         Begin VB.Menu mnuStudentPersonalReport 
            Caption         =   "Student Report"
            Index           =   411
         End
         Begin VB.Menu mnuStudentFeesReport 
            Caption         =   "Student Fees Report"
            Index           =   412
         End
         Begin VB.Menu mnuStudentMarksReport 
            Caption         =   "Student Marks Report"
            Index           =   420
         End
      End
      Begin VB.Menu mnuEmpReport 
         Caption         =   "Employee  Report"
         Begin VB.Menu mnuPesonal 
            Caption         =   "Personal"
         End
         Begin VB.Menu mnuEmplyement 
            Caption         =   "Employement"
         End
         Begin VB.Menu mnuBank 
            Caption         =   "Bank"
         End
      End
      Begin VB.Menu mnuBookReport 
         Caption         =   "Book Report"
         Index           =   45
      End
      Begin VB.Menu mnuSubjectRepoer 
         Caption         =   "Subject Report"
         Index           =   47
      End
      Begin VB.Menu mnuTeacherReprot 
         Caption         =   "Teacher Report "
         Index           =   46
      End
      Begin VB.Menu mnuCourseReport 
         Caption         =   "Course Report"
         Index           =   43
      End
      Begin VB.Menu mnuEnquiryReport 
         Caption         =   "Enquiry Report"
         Index           =   413
      End
      Begin VB.Menu mnuBookIssueReport 
         Caption         =   "Book Issue Report"
         Index           =   414
      End
      Begin VB.Menu mnuBookReturnReport 
         Caption         =   "Book Return Report"
         Index           =   415
      End
      Begin VB.Menu mnuTestReport 
         Caption         =   "Test Report"
         Index           =   416
      End
      Begin VB.Menu mnuPlacementReport 
         Caption         =   "Placement Report"
         Index           =   417
      End
      Begin VB.Menu mnuBatchAllotementReport 
         Caption         =   "Batch Allotement Report"
         Index           =   418
      End
      Begin VB.Menu mnuCompletionCertificateReport 
         Caption         =   "Completion Certificate Report"
         Index           =   419
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Too&ls"
      Index           =   4
      Begin VB.Menu mnucalculator 
         Caption         =   "Calculator"
         Index           =   51
      End
      Begin VB.Menu mnuSearchNotepad 
         Caption         =   "Notepad"
         Index           =   52
      End
   End
   Begin VB.Menu mnuContactUs 
      Caption         =   "&Contact us"
      Index           =   5
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Index           =   6
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
      Index           =   7
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnu_Click(Index As Integer)
frmIssueBook.Show
End Sub

Private Sub mnuBank_Click()
rptEmpBankDetails.Show
End Sub

Private Sub mnuBatch_Click(Index As Integer)
frmBatch.Show
End Sub



Private Sub mnuBatchAllotementReport_Click(Index As Integer)
rptBatchAllotement.Show
End Sub

Private Sub mnuBatchAllotment_Click(Index As Integer)
frmBatchAllote.Show
End Sub

Private Sub mnuBook_Click(Index As Integer)
frmBook.Show
End Sub

Private Sub mnuBookIssueReport_Click(Index As Integer)
rptIssueBook.Show
End Sub

Private Sub mnuBookReport_Click(Index As Integer)
rptBook.Show
End Sub

Private Sub mnuBookReturn_Click(Index As Integer)
frmBookReturn.Show
End Sub

Private Sub mnuBookReturnReport_Click(Index As Integer)
rptBookReturn.Show
End Sub

Private Sub mnucalculator_Click(Index As Integer)
Shell ("C:\Windows\System32\calc.exe")
End Sub





Private Sub mnuCompletionCertificate_Click(Index As Integer)
frmCompletionCertificate.Show
End Sub

Private Sub mnuCompletionCertificateReport_Click(Index As Integer)
rptCompletionCertificate.Show
End Sub

Private Sub mnuContactUs_Click(Index As Integer)
frmContactUs.Show
End Sub

Private Sub mnuCourse_Click(Index As Integer)
frmCourse.Show
End Sub

Private Sub mnuCourseReport_Click(Index As Integer)
rptcourse.Show
End Sub

Private Sub mnuEmployee_Click(Index As Integer)
frmEmployee.Show
End Sub



Private Sub mnuEmployeeAttendance_Click(Index As Integer)
frmEmployeeAttendance.Show
End Sub

Private Sub mnuEmplyement_Click()
rptEmpEmployementDetails.Show
End Sub

Private Sub mnuEnquiry_Click(Index As Integer)
frmEnquiry.Show
End Sub

Private Sub mnuEnquiryReport_Click(Index As Integer)
rptEnquiry.Show
End Sub

Private Sub mnuExit_Click(Index As Integer)
End
End Sub

Private Sub mnuFees_Click(Index As Integer)
frmFees.Show
End Sub

Private Sub mnuHelp_Click(Index As Integer)
frmHelp.Show
End Sub

Private Sub mnuPesonal_Click()
rptEmpPersonalDetails.Show
End Sub

Private Sub mnuPlacement_Click(Index As Integer)
frmplacement.Show
End Sub

Private Sub mnuPlacementReport_Click(Index As Integer)
rptPlacement.Show
End Sub

Private Sub mnuRenewal_Click(Index As Integer)
frmrenewal.Show
End Sub

Private Sub mnuSearchBatch_Click(Index As Integer)
frmSBatch.Show
End Sub

Private Sub mnusearchBook_Click(Index As Integer)
frmSBook.Show
End Sub

Private Sub mnuSearchCourse_Click(Index As Integer)
frmSCourse.Show
End Sub

Private Sub mnuSearchEmployee_Click(Index As Integer)
frmSEmployee.Show
End Sub

Private Sub mnuSearchEnquiry_Click(Index As Integer)
frmSEnquiry.Show
End Sub

Private Sub mnuSearchNotepad_Click(Index As Integer)
Shell ("C:\Windows\System32\notepad.exe")
End Sub

Private Sub mnuSearchStudent_Click(Index As Integer)
frmSStudent.Show
End Sub



Private Sub mnuSearchSubject_Click(Index As Integer)
frmSSubject.Show
End Sub

Private Sub mnuSearchTeacher_Click(Index As Integer)
frmSTeacher.Show
End Sub

Private Sub mnuStudent_Click(Index As Integer)
frmStudent.Show
End Sub

Private Sub mnuStudentAttendance_Click(Index As Integer)
frmStudentAttendance.Show
End Sub

Private Sub mnuStudentFeesReport_Click(Index As Integer)
rptStudentFeesDetails.Show
End Sub

Private Sub mnuStudentMarks_Click(Index As Integer)
frmStudentMarks.Show
End Sub

Private Sub mnuStudentMarksReport_Click(Index As Integer)
rptStudentMarks.Show
End Sub

Private Sub mnuStudentPersonalReport_Click(Index As Integer)
rptStudent.Show
End Sub



Private Sub mnuSubject_Click(Index As Integer)
frmSubject.Show
End Sub

Private Sub mnuSubjectRepoer_Click(Index As Integer)
rptSubject.Show
End Sub


Private Sub mnuTeacher_Click(Index As Integer)
frmTeacher.Show
End Sub

Private Sub mnuTeacherReprot_Click(Index As Integer)
rptTeacherDetails.Show
End Sub

Private Sub mnuTeaherAttendance_Click(Index As Integer)
frmTeacherAttendance.Show
End Sub

Private Sub mnuTest_Click(Index As Integer)
frmTest.Show
End Sub

Private Sub mnuTestReport_Click(Index As Integer)
rptTest.Show
End Sub
