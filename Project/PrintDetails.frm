VERSION 5.00
Begin VB.Form formPrintDetails 
   Caption         =   "Print Student Details"
   ClientHeight    =   10695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   10695
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox StudentID 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton PrintButton 
      Caption         =   "Print"
      Height          =   1215
      Left            =   4440
      TabIndex        =   2
      Top             =   9240
      Width           =   3975
   End
   Begin VB.CommandButton Preview 
      Caption         =   "Preview"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox PreviewBox 
      Height          =   8295
      Left            =   240
      ScaleHeight     =   8235
      ScaleWidth      =   12315
      TabIndex        =   0
      Top             =   720
      Width           =   12375
   End
   Begin VB.Label Label1 
      Caption         =   "Student ID"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "formPrintDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Preview_Click()

'Previews the student details document that is to be printed.

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
Set rs = db.OpenRecordset("StudentInfo")

Dim Found As Boolean

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
Set rs = db.OpenRecordset("StudentInfo")

PreviewBox.Cls

Found = False

If StudentID.Text = "" Then
    MsgBox "Enter ID."
    Exit Sub
End If

With rs
    While Not .EOF
        If StudentID.Text = !StudentID Then
            PreviewBox.Print "Student Details for " & !StudentName
            PreviewBox.Print
            PreviewBox.Print "Full Name: " & !StudentName
            PreviewBox.Print "Student ID: " & !StudentID
            PreviewBox.Print "Class: " & !Class & !Section
            PreviewBox.Print "Date of Birth: " & !DoB
            PreviewBox.Print "Gender: " & !Gender
            PreviewBox.Print "Date Enrolled: " & !EnrollDate
        End If
        
        If Not .EOF Then
            .MoveNext
        End If
    Wend
    rs.Close
End With

Set rs = Nothing
db.Close
Set db = Nothing

End Sub

Private Sub PrintButton_Click()

'Prints a hard copy of the student details for a specifed student ID.

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
Set rs = db.OpenRecordset("StudentInfo")

Found = False

If StudentID.Text = "" Then
    MsgBox "Enter ID."
    Exit Sub
End If

With rs
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
            Printer.Print "STUDENT DETAILS - HARD COPY"
            
            Printer.FontBold = False
            Printer.FontSize = 16
            Printer.Print "Full Name: " & !StudentName
            Printer.Print "Student ID: " & !StudentID
            Printer.Print "Class: " & !Class & !Section
            Printer.Print "Date of Birth: " & !DoB
            Printer.Print "Gender: " & !Gender
            Printer.Print "Date Enrolled: " & !EnrollDate
            Printer.EndDoc
            PrintButton.Enabled = False

        End If
        
        If Not .EOF Then
            .MoveNext
        End If
    Wend
    rs.Close
End With

Set rs = Nothing
db.Close
Set db = Nothing
End Sub
