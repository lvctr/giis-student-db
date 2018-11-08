VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GIIS Student Mark Database System"
   ClientHeight    =   4950
   ClientLeft      =   10515
   ClientTop       =   5475
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   14175
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox EngGrade 
      Height          =   285
      Left            =   13320
      TabIndex        =   82
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox L2Grade 
      Height          =   285
      Left            =   13320
      TabIndex        =   81
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox MathGrade 
      Height          =   285
      Left            =   13320
      TabIndex        =   80
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox ScienceGrade 
      Height          =   285
      Left            =   13320
      TabIndex        =   79
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox SSTGrade 
      Height          =   285
      Left            =   13320
      TabIndex        =   78
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox ComputingGrade 
      Height          =   285
      Left            =   13320
      TabIndex        =   77
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Update 
      Caption         =   "Update"
      Height          =   855
      Left            =   11880
      TabIndex        =   74
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox ComputingFA3 
      Height          =   285
      Left            =   10200
      TabIndex        =   71
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox ComputingSA1 
      Height          =   285
      Left            =   9480
      TabIndex        =   70
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox SSTFA3 
      Height          =   285
      Left            =   10200
      TabIndex        =   69
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox SSTSA1 
      Height          =   285
      Left            =   9480
      TabIndex        =   68
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox ScienceFA3 
      Height          =   285
      Left            =   10200
      TabIndex        =   67
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox ScienceSA1 
      Height          =   285
      Left            =   9480
      TabIndex        =   66
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox MathFA3 
      Height          =   285
      Left            =   10200
      TabIndex        =   65
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox MathSA1 
      Height          =   285
      Left            =   9480
      TabIndex        =   64
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox L2FA3 
      Height          =   285
      Left            =   10200
      TabIndex        =   63
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox L2SA1 
      Height          =   285
      Left            =   9480
      TabIndex        =   62
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox EngFA3 
      Height          =   285
      Left            =   10200
      TabIndex        =   61
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox EngSA1 
      Height          =   285
      Left            =   9480
      TabIndex        =   60
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox ComputingFinal 
      Height          =   285
      Left            =   12480
      TabIndex        =   47
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox ComputingSA2 
      Height          =   285
      Left            =   11640
      TabIndex        =   46
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox ComputingFA4 
      Height          =   285
      Left            =   10920
      TabIndex        =   45
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox ComputingFA2 
      Height          =   285
      Left            =   8760
      TabIndex        =   44
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox ComputingFA1 
      Height          =   285
      Left            =   8040
      TabIndex        =   43
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox SSTFinal 
      Height          =   285
      Left            =   12480
      TabIndex        =   42
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox SSTSA2 
      Height          =   285
      Left            =   11640
      TabIndex        =   41
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox SSTFA4 
      Height          =   285
      Left            =   10920
      TabIndex        =   40
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox SSTFA2 
      Height          =   285
      Left            =   8760
      TabIndex        =   39
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox SSTFA1 
      Height          =   285
      Left            =   8040
      TabIndex        =   38
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox ScienceFinal 
      Height          =   285
      Left            =   12480
      TabIndex        =   37
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox ScienceSA2 
      Height          =   285
      Left            =   11640
      TabIndex        =   36
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox ScienceFA4 
      Height          =   285
      Left            =   10920
      TabIndex        =   35
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox ScienceFA2 
      Height          =   285
      Left            =   8760
      TabIndex        =   34
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox ScienceFA1 
      Height          =   285
      Left            =   8040
      TabIndex        =   33
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox MathFinal 
      Height          =   285
      Left            =   12480
      TabIndex        =   32
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox MathSA2 
      Height          =   285
      Left            =   11640
      TabIndex        =   31
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox MathFA4 
      Height          =   285
      Left            =   10920
      TabIndex        =   30
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox MathFA2 
      Height          =   285
      Left            =   8760
      TabIndex        =   29
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox MathFA1 
      Height          =   285
      Left            =   8040
      TabIndex        =   28
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox L2Final 
      Height          =   285
      Left            =   12480
      TabIndex        =   27
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox L2SA2 
      Height          =   285
      Left            =   11640
      TabIndex        =   26
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox L2FA4 
      Height          =   285
      Left            =   10920
      TabIndex        =   25
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox L2FA2 
      Height          =   285
      Left            =   8760
      TabIndex        =   24
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox L2FA1 
      Height          =   285
      Left            =   8040
      TabIndex        =   23
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox EngFinal 
      Height          =   285
      Left            =   12480
      TabIndex        =   22
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox EngSA2 
      Height          =   285
      Left            =   11640
      TabIndex        =   21
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox EngFA4 
      Height          =   285
      Left            =   10920
      TabIndex        =   20
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox EngFA2 
      Height          =   285
      Left            =   8760
      TabIndex        =   19
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox EngFA1 
      Height          =   285
      Left            =   8040
      TabIndex        =   18
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton ModifyButton 
      Caption         =   "Edit"
      Height          =   855
      Left            =   4560
      TabIndex        =   17
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton SearchButton 
      Caption         =   "Look Up"
      Height          =   855
      Left            =   4560
      TabIndex        =   16
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton DeleteButton 
      Caption         =   "Delete"
      Height          =   855
      Left            =   4560
      TabIndex        =   15
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton AddButton 
      Caption         =   "Add"
      Height          =   855
      Left            =   4560
      TabIndex        =   7
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox StudentID 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   4095
   End
   Begin VB.TextBox StudentName 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox DoB 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "1/1/2000"
      Top             =   2640
      Width           =   4095
   End
   Begin VB.TextBox EnrollDate 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "1/1/2006"
      Top             =   4200
      Width           =   4095
   End
   Begin VB.ComboBox Class 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Text            =   "1"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ComboBox Section 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Text            =   "A"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ComboBox Gender 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "M"
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Label RecalcGrades 
      Alignment       =   1  'Right Justify
      Caption         =   "Recalculate Grades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   9240
      TabIndex        =   86
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label RecalcFinals 
      Alignment       =   1  'Right Justify
      Caption         =   "Recalculate Final Marks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   9240
      TabIndex        =   85
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Caption         =   "Student Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   84
      Top             =   360
      Width           =   7095
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "Grade"
      Height          =   255
      Left            =   13320
      TabIndex        =   83
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label CGPALabel 
      Alignment       =   1  'Right Justify
      Caption         =   "N/A"
      Height          =   255
      Left            =   8040
      TabIndex        =   76
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label RecalcCGPA 
      Alignment       =   1  'Right Justify
      Caption         =   "Recalculate CGPA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   9120
      TabIndex        =   75
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "FA3"
      Height          =   255
      Left            =   10200
      TabIndex        =   73
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "SA1"
      Height          =   255
      Left            =   9480
      TabIndex        =   72
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "CGPA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   59
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "Final"
      Height          =   255
      Left            =   12480
      TabIndex        =   58
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "SA2"
      Height          =   255
      Left            =   11640
      TabIndex        =   57
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "FA4"
      Height          =   255
      Left            =   10920
      TabIndex        =   56
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "FA2"
      Height          =   255
      Left            =   8760
      TabIndex        =   55
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "FA1"
      Height          =   255
      Left            =   8040
      TabIndex        =   54
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Computing"
      Height          =   255
      Left            =   6840
      TabIndex        =   53
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Math"
      Height          =   255
      Left            =   6840
      TabIndex        =   52
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Science"
      Height          =   255
      Left            =   6840
      TabIndex        =   51
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Social Science"
      Height          =   255
      Left            =   6840
      TabIndex        =   50
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "2nd Language"
      Height          =   255
      Left            =   6840
      TabIndex        =   49
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "English"
      Height          =   255
      Left            =   6840
      TabIndex        =   48
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Student ID"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Student Name"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Class"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Section"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Gender"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "Enrollment Date (format: M/D/YYYY)"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   "Date of Birth (format: M/D/YYYY)"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Menu AdminTools 
      Caption         =   "Admin Tools"
      Begin VB.Menu ChangePW 
         Caption         =   "Change Password"
      End
      Begin VB.Menu BackupMenu 
         Caption         =   "Backup"
      End
      Begin VB.Menu RestoreMenu 
         Caption         =   "Restore"
      End
   End
   Begin VB.Menu Print 
      Caption         =   "Print"
      Begin VB.Menu PrintDetails 
         Caption         =   "Print Student Details"
      End
      Begin VB.Menu PrintReport 
         Caption         =   "Print Report Card"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_Click()
formAbout.Show
End Sub

Private Sub BackupMenu_Click()
formBackup.Show
End Sub

Private Sub ChangePW_Click()
formChangePW.Show
End Sub

Private Sub DeleteButton_Click()

'Deletes the specified student from the database.

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
Set asd = db.OpenRecordset("StudentInfo")
Set asx = db.OpenRecordset("MarkInfo")

With asd
    While Not .EOF
        If StudentID.Text = !StudentID Then
            asd.Delete
        End If
    
        If Not .EOF Then
            .MoveNext
        End If
    Wend
End With

With asx
    While Not .EOF
        If StudentID.Text = !StudentID Then
            asx.Delete
        End If
    
        If Not .EOF Then
            .MoveNext
        End If
    Wend
    
    asd.Close
    asx.Close
    
    MsgBox "Deleted successfully."
End With

Set asd = Nothing
Set asx = Nothing
db.Close
Set db = Nothing

End Sub

Private Sub Form_Load()

'Adds the drop down menu items.

   Class.AddItem "1"
   Class.AddItem "2"
   Class.AddItem "3"
   Class.AddItem "4"
   Class.AddItem "5"
   Class.AddItem "6"
   Class.AddItem "7"
   Class.AddItem "8"
   Class.AddItem "9"
   Class.AddItem "10"
   Class.AddItem "11"
   Class.AddItem "12"
   
   Section.AddItem "A"
   Section.AddItem "B"
   Section.AddItem "C"
   Section.AddItem "D"
   
   Gender.AddItem "M"
   Gender.AddItem "F"
End Sub

Private Sub AddButton_Click()

'Adds a new student to the database under the specified student ID.

Dim dup As Boolean
Dim dupa As Boolean
Dim dupb As Boolean

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
Set asd = db.OpenRecordset("StudentInfo")
Set asx = db.OpenRecordset("MarkInfo")

'Duplication check
dup = False
dupa = False
dupb = False

With asd
    While Not .EOF
        If !StudentID = StudentID Then
            dupa = True
        End If
        If Not .EOF Then
            .MoveNext
        End If
    Wend
End With

With asx
     While Not .EOF
        If !StudentID = StudentID Then
            dupb = True
        End If
        If Not .EOF Then
            .MoveNext
        End If
    Wend
End With

If dupa = True Then
    dup = True
ElseIf dupb = True Then
    dup = True
Else
    dup = False
End If
'End of duplication check

If dup = False Then
With asd
asd.AddNew
        asd!StudentID = StudentID
        asd!StudentName = StudentName
        asd!Class = Class
        asd!Section = Section
        asd!DoB = DoB
        asd!Gender = Gender
        asd!EnrollDate = EnrollDate
asd.Update
asd.Close
End With

With asx
asx.AddNew
        asx!StudentID = StudentID
        asx!CGPA = 0
        
        asx!EngFA1 = 0
        asx!EngFA2 = 0
        asx!EngSA1 = 0
        asx!EngFA3 = 0
        asx!EngFA4 = 0
        asx!EngSA2 = 0
        asx!EngFinal = 0
        asx!EngGrade = "F"
                
        asx!L2FA1 = 0
        asx!L2FA2 = 0
        asx!L2SA1 = 0
        asx!L2FA3 = 0
        asx!L2FA4 = 0
        asx!L2SA2 = 0
        asx!L2Final = 0
        asx!L2Grade = "F"
                
        asx!MathFA1 = 0
        asx!MathFA2 = 0
        asx!MathSA1 = 0
        asx!MathFA3 = 0
        asx!MathFA4 = 0
        asx!MathSA2 = 0
        asx!MathFinal = 0
        asx!MathGrade = "F"
                
        asx!ScienceFA1 = 0
        asx!ScienceFA2 = 0
        asx!ScienceSA1 = 0
        asx!ScienceFA3 = 0
        asx!ScienceFA4 = 0
        asx!ScienceSA2 = 0
        asx!ScienceFinal = 0
        asx!ScienceGrade = "F"
                
        asx!SSTFA1 = 0
        asx!SSTFA2 = 0
        asx!SSTSA1 = 0
        asx!SSTFA3 = 0
        asx!SSTFA4 = 0
        asx!SSTSA2 = 0
        asx!SSTFinal = 0
        asx!SSTGrade = "F"
                
        asx!ComputingFA1 = 0
        asx!ComputingFA2 = 0
        asx!ComputingSA1 = 0
        asx!ComputingFA3 = 0
        asx!ComputingFA4 = 0
        asx!ComputingSA2 = 0
        asx!ComputingFinal = 0
        asx!ComputingGrade = "F"
        
        MsgBox "Added successfully."
asx.Update
asx.Close
End With
Else
    MsgBox ("Student ID already exists.")
End If

Set test = Nothing
db.Close
Set db = Nothing

End Sub

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub ModifyButton_Click()

'Edits the details of the specified student.

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
Set asd = db.OpenRecordset("StudentInfo")

If StudentID.Text = "" Then
    MsgBox ("Please enter Student ID.")
    Exit Sub
End If

With asd
    While Not .EOF
    If StudentID.Text = !StudentID Then
        asd.Edit
        asd!StudentName = StudentName
        asd!Class = Class
        asd!Section = Section
        asd!DoB = DoB
        asd!Gender = Gender
        asd!EnrollDate = EnrollDate
        asd.Update
    End If
    
    If Not .EOF Then
        .MoveNext
    End If
    Wend
    
    asd.Close
    
    MsgBox "Edited successfully."
End With

Set asd = Nothing
db.Close
Set db = Nothing

End Sub

Private Sub PrintDetails_Click()
formPrintDetails.Show
End Sub

Private Sub PrintReport_Click()
formPrintReport.Show
End Sub

Private Sub RecalcCGPA_Click()

'Calculates the cumulative grade point average from the report data.

Dim RawCGPA As Single
Dim Sub1 As Integer
Dim Sub2 As Integer
Dim Sub3 As Integer
Dim Sub4 As Integer
Dim Sub5 As Integer
Dim Sub6 As Integer

Sub1 = EngFinal.Text
Sub2 = L2Final.Text
Sub3 = MathFinal.Text
Sub4 = SSTFinal.Text
Sub5 = ComputingFinal.Text
Sub6 = ScienceFinal.Text

RawCGPA = (Sub1 + Sub2 + Sub3 + Sub4 + Sub5 + Sub6) / 60

CGPALabel.Caption = Format(RawCGPA, "Standard")

End Sub

Private Sub RecalcFinals_Click()

'Calculates the final marks from the report data.

Dim aEngFA1 As Integer
Dim aEngFA2 As Integer
Dim aEngSA1 As Integer
Dim aEngFA3 As Integer
Dim aEngFA4 As Integer
Dim aEngSA2 As Integer

Dim aL2FA1 As Integer
Dim aL2FA2 As Integer
Dim aL2SA1 As Integer
Dim aL2FA3 As Integer
Dim aL2FA4 As Integer
Dim aL2SA2 As Integer

Dim aMathFA1 As Integer
Dim aMathFA2 As Integer
Dim aMathSA1 As Integer
Dim aMathFA3 As Integer
Dim aMathFA4 As Integer
Dim aMathSA2 As Integer

Dim aScienceFA1 As Integer
Dim aScienceFA2 As Integer
Dim aScienceSA1 As Integer
Dim aScienceFA3 As Integer
Dim aScienceFA4 As Integer
Dim aScienceSA2 As Integer

Dim aSSTFA1 As Integer
Dim aSSTFA2 As Integer
Dim aSSTSA1 As Integer
Dim aSSTFA3 As Integer
Dim aSSTFA4 As Integer
Dim aSSTSA2 As Integer

Dim aComputingFA1 As Integer
Dim aComputingFA2 As Integer
Dim aComputingSA1 As Integer
Dim aComputingFA3 As Integer
Dim aComputingFA4 As Integer
Dim aComputingSA2 As Integer

Dim aEngFinal As Single
Dim aL2Final As Single
Dim aMathFinal As Single
Dim aScienceFinal As Single
Dim aSSTFinal As Single
Dim aComputingFinal As Single

aEngFA1 = EngFA1.Text
aEngFA2 = EngFA2.Text
aEngSA1 = EngSA1.Text
aEngFA3 = EngFA3.Text
aEngFA4 = EngFA4.Text
aEngSA2 = EngSA2.Text

aL2FA1 = L2FA1.Text
aL2FA2 = L2FA2.Text
aL2SA1 = L2SA1.Text
aL2FA3 = L2FA3.Text
aL2FA4 = L2FA4.Text
aL2SA2 = L2SA2.Text

aMathFA1 = MathFA1.Text
aMathFA2 = MathFA2.Text
aMathSA1 = MathSA1.Text
aMathFA3 = MathFA3.Text
aMathFA4 = MathFA4.Text
aMathSA2 = MathSA2.Text

aScienceFA1 = ScienceFA1.Text
aScienceFA2 = ScienceFA2.Text
aScienceSA1 = ScienceSA1.Text
aScienceFA3 = ScienceFA3.Text
aScienceFA4 = ScienceFA4.Text
aScienceSA2 = ScienceSA2.Text

aSSTFA1 = SSTFA1.Text
aSSTFA2 = SSTFA2.Text
aSSTSA1 = SSTSA1.Text
aSSTFA3 = SSTFA3.Text
aSSTFA4 = SSTFA4.Text
aSSTSA2 = SSTSA2.Text

aComputingFA1 = ComputingFA1.Text
aComputingFA2 = ComputingFA2.Text
aComputingSA1 = ComputingSA1.Text
aComputingFA3 = ComputingFA3.Text
aComputingFA4 = ComputingFA4.Text
aComputingSA2 = ComputingSA2.Text

aEngFinal = (aEngFA1 + aEngFA2 + aEngSA1 + aEngFA3 + aEngFA4 + aEngSA2) / 6
aL2Final = (aL2FA1 + aL2FA2 + aL2SA1 + aL2FA3 + aL2FA4 + aL2SA2) / 6
aMathFinal = (aMathFA1 + aMathFA2 + aMathSA1 + aMathFA3 + aMathFA4 + aMathSA2) / 6
aScienceFinal = (aScienceFA1 + aScienceFA2 + aScienceSA1 + aScienceFA3 + aScienceFA4 + aScienceSA2) / 6
aSSTFinal = (aSSTFA1 + aSSTFA2 + aSSTSA1 + aSSTFA3 + aSSTFA4 + aSSTSA2) / 6
aComputingFinal = (aComputingFA1 + aComputingFA2 + aComputingSA1 + aComputingFA3 + aComputingFA4 + aComputingSA2) / 6

aEngFinal = Int(aEngFinal)
aL2Final = Int(aL2Final)
aMathFinal = Int(aMathFinal)
aScienceFinal = Int(aScienceFinal)
aSSTFinal = Int(aSSTFinal)
aComputingFinal = Int(aComputingFinal)

EngFinal.Text = aEngFinal
L2Final.Text = aL2Final
MathFinal.Text = aMathFinal
ScienceFinal.Text = aScienceFinal
SSTFinal.Text = aSSTFinal
ComputingFinal.Text = aComputingFinal

End Sub

Private Sub RecalcGrades_Click()

'Calculates the grades from the report data.

Select Case EngFinal
    Case Is < 40
        EngGrade = "F"
    Case 41 To 50
        EngGrade = "E"
    Case 51 To 60
        EngGrade = "D"
    Case 61 To 70
        EngGrade = "C"
    Case 71 To 80
        EngGrade = "B"
    Case 81 To 90
        EngGrade = "A"
    Case 91 To 100
        EngGrade = "A+"
    Case Else
        MsgBox ("Incorrect grade for English.")
End Select

Select Case L2Final
    Case Is < 40
        L2Grade = "F"
    Case 41 To 50
        L2Grade = "E"
    Case 51 To 60
        L2Grade = "D"
    Case 61 To 70
        L2Grade = "C"
    Case 71 To 80
        L2Grade = "B"
    Case 81 To 90
        L2Grade = "A"
    Case 91 To 100
        L2Grade = "A+"
    Case Else
        MsgBox ("Incorrect grade for 2nd Language.")
End Select

Select Case MathFinal
    Case Is < 40
        MathGrade = "F"
    Case 41 To 50
        MathGrade = "E"
    Case 51 To 60
        MathGrade = "D"
    Case 61 To 70
        MathGrade = "C"
    Case 71 To 80
        MathGrade = "B"
    Case 81 To 90
        MathGrade = "A"
    Case 91 To 100
        MathGrade = "A+"
    Case Else
        MsgBox ("Incorrect grade for Mathematics.")
End Select

Select Case ScienceFinal
    Case Is < 40
        ScienceGrade = "F"
    Case 41 To 50
        ScienceGrade = "E"
    Case 51 To 60
        ScienceGrade = "D"
    Case 61 To 70
        ScienceGrade = "C"
    Case 71 To 80
        ScienceGrade = "B"
    Case 81 To 90
        ScienceGrade = "A"
    Case 91 To 100
        ScienceGrade = "A+"
    Case Else
        MsgBox ("Incorrect grade for Science.")
End Select

Select Case SSTFinal
    Case Is < 40
        SSTGrade = "F"
    Case 41 To 50
        SSTGrade = "E"
    Case 51 To 60
        SSTGrade = "D"
    Case 61 To 70
        SSTGrade = "C"
    Case 71 To 80
        SSTGrade = "B"
    Case 81 To 90
        SSTGrade = "A"
    Case 91 To 100
        SSTGrade = "A+"
    Case Else
        MsgBox ("Incorrect grade for Social Science.")
End Select

Select Case ComputingFinal
    Case Is < 40
        ComputingGrade = "F"
    Case 41 To 50
        ComputingGrade = "E"
    Case 51 To 60
        ComputingGrade = "D"
    Case 61 To 70
        ComputingGrade = "C"
    Case 71 To 80
        ComputingGrade = "B"
    Case 81 To 90
        ComputingGrade = "A"
    Case 91 To 100
        ComputingGrade = "A+"
    Case Else
        MsgBox ("Incorrect grade for Computing.")
End Select

EngGrade.Text = EngGrade
L2Grade.Text = L2Grade
MathGrade.Text = MathGrade
ScienceGrade.Text = ScienceGrade
SSTGrade.Text = SSTGrade
ComputingGrade.Text = ComputingGrade

End Sub

Private Sub RestoreMenu_Click()
formRestore.Show
End Sub

Private Sub SearchButton_Click()

'Searches for the specified student in the database.

Dim FoundA As Boolean
Dim FoundB As Boolean

FoundA = False
FoundB = False

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
Set rs = db.OpenRecordset("StudentInfo")
Set rd = db.OpenRecordset("MarkInfo")

If StudentID.Text = "" Then
    MsgBox "Enter ID."
    Exit Sub
End If

With rs
    While Not .EOF And FoundA = False
        If StudentID.Text = !StudentID Then
            FoundA = True
            StudentName.Text = !StudentName
            Class.Text = !Class
            Section.Text = !Section
            DoB.Text = !DoB
            Gender.Text = !Gender
            EnrollDate.Text = !EnrollDate
        End If
        
        If Not .EOF Then
        .MoveNext
        End If
    Wend
    rs.Close
End With

With rd
    While Not .EOF And FoundB = False
        If StudentID.Text = !StudentID Then
        FoundB = True
        CGPALabel.Caption = !CGPA
        
        EngFA1.Text = !EngFA1
        EngFA2.Text = !EngFA2
        EngSA1.Text = !EngSA1
        EngFA3.Text = !EngFA3
        EngFA4.Text = !EngFA4
        EngSA2.Text = !EngSA2
        EngFinal.Text = !EngFinal
        EngGrade.Text = !EngGrade
                
        L2FA1.Text = !L2FA1
        L2FA2.Text = !L2FA2
        L2SA1.Text = !L2SA1
        L2FA3.Text = !L2FA3
        L2FA4.Text = !L2FA4
        L2SA2.Text = !L2SA2
        L2Final.Text = !L2Final
        L2Grade.Text = !L2Grade
                
        MathFA1.Text = !MathFA1
        MathFA2.Text = !MathFA2
        MathSA1.Text = !MathSA1
        MathFA3.Text = !MathFA3
        MathFA4.Text = !MathFA4
        MathSA2.Text = !MathSA2
        MathFinal.Text = !MathFinal
        MathGrade.Text = !MathGrade
                
        ScienceFA1.Text = !ScienceFA1
        ScienceFA2.Text = !ScienceFA2
        ScienceSA1.Text = !ScienceSA1
        ScienceFA3.Text = !ScienceFA3
        ScienceFA4.Text = !ScienceFA4
        ScienceSA2.Text = !ScienceSA2
        ScienceFinal.Text = !ScienceFinal
        ScienceGrade.Text = !ScienceGrade
                
        SSTFA1.Text = !SSTFA1
        SSTFA2.Text = !SSTFA2
        SSTSA1.Text = !SSTSA1
        SSTFA3.Text = !SSTFA3
        SSTFA4.Text = !SSTFA4
        SSTSA2.Text = !SSTSA2
        SSTFinal.Text = !SSTFinal
        SSTGrade.Text = !SSTGrade
                
        ComputingFA1.Text = !ComputingFA1
        ComputingFA2.Text = !ComputingFA2
        ComputingSA1.Text = !ComputingSA1
        ComputingFA3.Text = !ComputingFA3
        ComputingFA4.Text = !ComputingFA4
        ComputingSA2.Text = !ComputingSA2
        ComputingFinal.Text = !ComputingFinal
        ComputingGrade.Text = !ComputingGrade
        
        MsgBox ("Found student.")
        End If
        
        If Not .EOF Then
        .MoveNext
        End If
    Wend
    rd.Close
End With

Set rs = Nothing
Set rd = Nothing
db.Close
Set db = Nothing

End Sub

Private Sub Update_Click()

'Updates the student report.

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
Set asd = db.OpenRecordset("MarkInfo")

If StudentID.Text = "" Then
    MsgBox ("Please enter Student ID.")
    Exit Sub
End If

With asd
    While Not .EOF
    If StudentID.Text = asd!StudentID Then
    asd.Edit
        asd!CGPA = CGPALabel.Caption
        
        asd!EngFA1 = EngFA1
        asd!EngFA2 = EngFA2
        asd!EngSA1 = EngSA1
        asd!EngFA3 = EngFA3
        asd!EngFA4 = EngFA4
        asd!EngSA2 = EngSA2
        asd!EngFinal = EngFinal
        asd!EngGrade = EngGrade
                
        asd!L2FA1 = L2FA1
        asd!L2FA2 = L2FA2
        asd!L2SA1 = L2SA1
        asd!L2FA3 = L2FA3
        asd!L2FA4 = L2FA4
        asd!L2SA2 = L2SA2
        asd!L2Final = L2Final
        asd!L2Grade = L2Grade
                
        asd!MathFA1 = MathFA1
        asd!MathFA2 = MathFA2
        asd!MathSA1 = MathSA1
        asd!MathFA3 = MathFA3
        asd!MathFA4 = MathFA4
        asd!MathSA2 = MathSA2
        asd!MathFinal = MathFinal
        asd!MathGrade = MathGrade
                
        asd!ScienceFA1 = ScienceFA1
        asd!ScienceFA2 = ScienceFA2
        asd!ScienceSA1 = ScienceSA1
        asd!ScienceFA3 = ScienceFA3
        asd!ScienceFA4 = ScienceFA4
        asd!ScienceSA2 = ScienceSA2
        asd!ScienceFinal = ScienceFinal
        asd!ScienceGrade = ScienceGrade
                
        asd!SSTFA1 = SSTFA1
        asd!SSTFA2 = SSTFA2
        asd!SSTSA1 = SSTSA1
        asd!SSTFA3 = SSTFA3
        asd!SSTFA4 = SSTFA4
        asd!SSTSA2 = SSTSA2
        asd!SSTFinal = SSTFinal
        asd!SSTGrade = SSTGrade
                
        asd!ComputingFA1 = ComputingFA1
        asd!ComputingFA2 = ComputingFA2
        asd!ComputingSA1 = ComputingSA1
        asd!ComputingFA3 = ComputingFA3
        asd!ComputingFA4 = ComputingFA4
        asd!ComputingSA2 = ComputingSA2
        asd!ComputingFinal = ComputingFinal
        asd!ComputingGrade = ComputingGrade
    asd.Update
    End If
    
    If Not .EOF Then
    .MoveNext
    End If
    Wend
    
    MsgBox "Update successful."
    asd.Close
    
End With

Set asd = Nothing
db.Close
Set db = Nothing


End Sub

