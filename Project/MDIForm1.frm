VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000F&
   Caption         =   "GIIS Student Marks Database System"
   ClientHeight    =   9255
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13230
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MenuAdminTools 
      Caption         =   "Admin Tools"
      Begin VB.Menu MenuChangePW 
         Caption         =   "Change Password"
      End
      Begin VB.Menu MenuBnR 
         Caption         =   "Backup and Restore"
      End
   End
   Begin VB.Menu MenuStuMgr 
      Caption         =   "Student Management"
      Begin VB.Menu MenuAddStu 
         Caption         =   "Add Student"
      End
      Begin VB.Menu MenuModStu 
         Caption         =   "Modify Student"
      End
   End
   Begin VB.Menu MenuPrint 
      Caption         =   "Print"
      Begin VB.Menu MenuPrintStu 
         Caption         =   "Print Student Details"
      End
      Begin VB.Menu MenuPrintRep 
         Caption         =   "Print Report Card"
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "Help"
      Begin VB.Menu MenuDocs 
         Caption         =   "Documentation"
      End
      Begin VB.Menu MenuManual 
         Caption         =   "Manual"
      End
      Begin VB.Menu MenuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MenuAbout_Click()
About.Show
End Sub

Private Sub MenuAddStu_Click()
AddStudent.Show
End Sub

Private Sub MenuBnR_Click()
BackupRestore.Show
End Sub

Private Sub MenuChangePW_Click()
ChangePW.Show
End Sub

Private Sub MenuModStu_Click()
ModifyStudent.Show
End Sub

Private Sub MenuPrintRep_Click()
PrintReportCard.Show
End Sub

Private Sub MenuPrintStu_Click()
PrintDetails.Show
End Sub
