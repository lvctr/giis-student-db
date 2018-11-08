VERSION 5.00
Begin VB.Form formPrintReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Report Card"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PreviewBox 
      Height          =   8295
      Left            =   240
      ScaleHeight     =   8235
      ScaleWidth      =   7515
      TabIndex        =   3
      Top             =   720
      Width           =   7575
   End
   Begin VB.CommandButton Preview 
      Caption         =   "Preview"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton PrintButton 
      Caption         =   "Print"
      Height          =   1215
      Left            =   2040
      TabIndex        =   1
      Top             =   9240
      Width           =   3975
   End
   Begin VB.TextBox StudentID 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Student ID"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "formPrintReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Preview_Click()

'Previews the student report card document that is to be printed.

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
Set rs = db.OpenRecordset("StudentInfo")

Dim Found As Boolean

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
Set asd = db.OpenRecordset("MarkInfo")
Set asx = db.OpenRecordset("StudentInfo")

PreviewBox.Cls

Found = False

If StudentID.Text = "" Then
    MsgBox "Enter ID."
    Exit Sub
End If

With asx
    While Not .EOF
        If StudentID.Text = !StudentID Then
            thename = !StudentName
        End If
        
        If Not .EOF Then
            .MoveNext
        End If
    Wend
    asx.Close
End With

With asd
    While Not .EOF Or Found = False
        If StudentID.Text = !StudentID Then
            Found = True
            
            PreviewBox.Print "Report Card for Student #" & !StudentID; ", " & thename
            PreviewBox.Print
            PreviewBox.Print "Subject"; Tab(22); "FA1"; Tab(27); "FA2"; Tab(32); "SA1"; Tab(37); "FA3"; Tab(42); "FA4"; Tab(47); "SA2"; Tab(52); "Final Marks"; Tab(64); "Overall Grade"
            PreviewBox.Print "English"; Tab(22); !EngFA1; Tab(27); !EngFA2; Tab(32); !EngSA1; Tab(37); !EngFA3; Tab(42); !EngFA4; Tab(47); !EngSA2; Tab(52); !EngFinal; Tab(64); !EngGrade
            PreviewBox.Print "Second Language"; Tab(22); !L2FA1; Tab(27); !L2FA2; Tab(32); !L2SA1; Tab(37); !L2FA3; Tab(42); !L2FA4; Tab(47); !L2SA2; Tab(52); !L2Final; Tab(64); !L2Grade
            PreviewBox.Print "Mathematics"; Tab(22); !MathFA1; Tab(27); !MathFA2; Tab(32); !MathSA1; Tab(37); !MathFA3; Tab(42); !MathFA4; Tab(47); !MathSA2; Tab(52); !MathFinal; Tab(64); !MathGrade
            PreviewBox.Print "Science"; Tab(22); !ScienceFA1; Tab(27); !ScienceFA2; Tab(32); !ScienceSA1; Tab(37); !ScienceFA3; Tab(42); !ScienceFA4; Tab(47); !ScienceSA2; Tab(52); !ScienceFinal; Tab(64); !ScienceGrade
            PreviewBox.Print "Social Science"; Tab(22); !SSTFA1; Tab(27); !SSTFA2; Tab(32); !SSTSA1; Tab(37); !SSTFA3; Tab(42); !SSTFA4; Tab(47); !SSTSA2; Tab(52); !SSTFinal; Tab(64); !SSTGrade
            PreviewBox.Print "Computing"; Tab(22); !ComputingFA1; Tab(27); !ComputingFA2; Tab(32); !ComputingSA1; Tab(37); !ComputingFA3; Tab(42); !ComputingFA4; Tab(47); !ComputingSA2; Tab(52); !ComputingFinal; Tab(64); !ComputingGrade
        End If
        
        If Not .EOF Or Found = False Then
        .MoveNext
        End If
    Wend
    asd.Close
End With

Set asd = Nothing
Set asx = Nothing
db.Close
Set db = Nothing

End Sub

Private Sub PrintButton_Click()

'Prints a hard copy of the student details for a specifed student ID.

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
Set rs = db.OpenRecordset("StudentInfo")

Dim Found As Boolean
Dim thename As String

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
Set asd = db.OpenRecordset("MarkInfo")
Set asx = db.OpenRecordset("StudentInfo")

PreviewBox.Cls

Found = False

If StudentID.Text = "" Then
    MsgBox "Enter ID."
    Exit Sub
End If

With asx
    While Not .EOF
        If StudentID.Text = !StudentID Then
            thename = !StudentName
        End If
        
        If Not .EOF Then
            .MoveNext
        End If
    Wend
    asx.Close
End With

With asd
    While Not .EOF
        If StudentID.Text = !StudentID Then
            Found = True
            
            Printer.ScaleMode = vbInches
            Printer.FontName = "Arial"
            Printer.CurrentX = 3
            Printer.CurrentY = 0.2
            
            Printer.Print ""
            Printer.FontBold = True
            Printer.FontSize = 24
            Printer.Print "REPORT CARD - HARD COPY"
            
            Printer.FontBold = False
            Printer.FontSize = 16
            
            PreviewBox.Print "Report Card for Student #" & !StudentID; ", " & thename
            PreviewBox.Print
            PreviewBox.Print "Subject"; Tab(22); "FA1"; Tab(27); "FA2"; Tab(32); "SA1"; Tab(37); "FA3"; Tab(42); "FA4"; Tab(47); "SA2"; Tab(52); "Final Marks"; Tab(64); "Overall Grade"
            PreviewBox.Print "English"; Tab(22); !EngFA1; Tab(27); !EngFA2; Tab(32); !EngSA1; Tab(37); !EngFA3; Tab(42); !EngFA4; Tab(47); !EngSA2; Tab(52); !EngFinal; Tab(64); !EngGrade
            PreviewBox.Print "Second Language"; Tab(22); !L2FA1; Tab(27); !L2FA2; Tab(32); !L2SA1; Tab(37); !L2FA3; Tab(42); !L2FA4; Tab(47); !L2SA2; Tab(52); !L2Final; Tab(64); !L2Grade
            PreviewBox.Print "Mathematics"; Tab(22); !MathFA1; Tab(27); !MathFA2; Tab(32); !MathSA1; Tab(37); !MathFA3; Tab(42); !MathFA4; Tab(47); !MathSA2; Tab(52); !MathFinal; Tab(64); !MathGrade
            PreviewBox.Print "Science"; Tab(22); !ScienceFA1; Tab(27); !ScienceFA2; Tab(32); !ScienceSA1; Tab(37); !ScienceFA3; Tab(42); !ScienceFA4; Tab(47); !ScienceSA2; Tab(52); !ScienceFinal; Tab(64); !ScienceGrade
            PreviewBox.Print "Social Science"; Tab(22); !SSTFA1; Tab(27); !SSTFA2; Tab(32); !SSTSA1; Tab(37); !SSTFA3; Tab(42); !SSTFA4; Tab(47); !SSTSA2; Tab(52); !SSTFinal; Tab(64); !SSTGrade
            PreviewBox.Print "Computing"; Tab(22); !ComputingFA1; Tab(27); !ComputingFA2; Tab(32); !ComputingSA1; Tab(37); !ComputingFA3; Tab(42); !ComputingFA4; Tab(47); !ComputingSA2; Tab(52); !ComputingFinal; Tab(64); !ComputingGrade
        End If
        
        If Not .EOF Or Found = False Then
        .MoveNext
        End If
    Wend
    asd.Close
End With

Set asd = Nothing
Set asx = Nothing
db.Close
Set db = Nothing

End Sub
