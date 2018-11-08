VERSION 5.00
Begin VB.Form formChangePW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   1695
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Submit 
      Caption         =   "Submit"
      Height          =   975
      Left            =   4920
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox ConfirmPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "O"
      TabIndex        =   5
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox NewPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "O"
      TabIndex        =   4
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox CurrentPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "O"
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm Password"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "New Password"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Current Password"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "formChangePW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Submit_Click()

'Changes password to login to the database.

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
Set ors = db.OpenRecordset("AdminDetails")

With ors
    If CurrentPass.Text = !Password Then 'Checks if current password matches with the database.
        If NewPass.Text = ConfirmPass.Text Then 'Checks if new password matches with confirm password.
            ors.Edit
            ors!Password = NewPass
            ors.Update
            ors.Close
            MsgBox ("Password changed.")
        Else
            MsgBox ("Mismatched passwords. Please try again.")
        End If
    Else
        MsgBox ("Incorrect password. Please try again.")
    End If
End With
    
Set ors = Nothing
db.Close
Set db = Nothing

End Sub
