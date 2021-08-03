VERSION 5.00
Begin VB.MDIForm MDIuserdashboard 
   BackColor       =   &H8000000C&
   Caption         =   "Student Dashboard | Arcade Business College"
   ClientHeight    =   4030
   ClientLeft      =   195
   ClientTop       =   507
   ClientWidth     =   9230
   Icon            =   "MDIuserdashboard.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIuserdashboard.frx":1084A
   WindowState     =   2  'Maximized
   Begin VB.Menu studentfile 
      Caption         =   "File"
      Begin VB.Menu studentDoubt 
         Caption         =   "Doubt Class"
      End
      Begin VB.Menu studentLibraryPass 
         Caption         =   "Library Pass"
      End
      Begin VB.Menu studentSemesterResult 
         Caption         =   "Semester Result"
      End
   End
   Begin VB.Menu studentview 
      Caption         =   "View"
      Begin VB.Menu studentAttendence 
         Caption         =   "Attendence"
      End
      Begin VB.Menu studentnotification 
         Caption         =   "Notification"
      End
      Begin VB.Menu studentroutine 
         Caption         =   "Routine"
      End
   End
   Begin VB.Menu studentproject 
      Caption         =   "Submit Project"
   End
   Begin VB.Menu studentfee 
      Caption         =   "Fee Status"
   End
   Begin VB.Menu studentfine 
      Caption         =   "Fine Status"
   End
   Begin VB.Menu studentfacultyDetails 
      Caption         =   "Faculty Details"
      Begin VB.Menu studentfacultyAvailability 
         Caption         =   "Availability"
      End
      Begin VB.Menu studentfacultyQualification 
         Caption         =   "Qualification"
      End
   End
   Begin VB.Menu studentaccount 
      Caption         =   "Account"
      Begin VB.Menu studentprofile 
         Caption         =   "View Profile"
      End
      Begin VB.Menu studentChangeDetails 
         Caption         =   "Change Details"
      End
   End
   Begin VB.Menu studentlogout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "MDIuserdashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub studentAttendence_Click()
attendenceDialog.Show
End Sub

Private Sub studentChangeDetails_Click()
UdateuserDialog.Show

End Sub

Private Sub studentDoubt_Click()
doubtclassdialog.Show

End Sub

Private Sub studentfacultyAvailability_Click()
facultyDataReport.Show
End Sub

Private Sub studentfacultyQualification_Click()

fqulaificationDataReport.Show

End Sub

Private Sub studentfee_Click()
FeeStatusDialog.Show

End Sub

Private Sub studentfine_Click()
fineDialog.Show

End Sub

Private Sub studentLibraryPass_Click()
studentdbpassDialog.Show

End Sub


Private Sub studentlogout_Click()
End
End Sub

Private Sub studentnotification_Click()
notificationDataReport.Show
End Sub

Private Sub studentprofile_Click()

Userprofile.Show

End Sub


Private Sub studentproject_Click()
submirrojectform.Show

End Sub

Private Sub studentroutine_Click()
studentroutineDialog.Show

End Sub

Private Sub studentSemesterResult_Click()
semresultform.Show

End Sub
