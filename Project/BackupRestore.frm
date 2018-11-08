VERSION 5.00
Begin VB.Form formBackup 
   Caption         =   "Backup"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Backup 
      Caption         =   "Backup"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   4815
   End
   Begin VB.DirListBox Directory 
      Height          =   2115
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4815
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.TextBox FolderName 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "GIIS Database Backup"
      Top             =   2640
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Folder Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
End
Attribute VB_Name = "formBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Backup_Click()

'Copies the database to the specified location.

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
db.Close

Dim FSO As New FileSystemObject
If Dir(Directory.Path & "/" & FolderName.Text) <> "" Then
FSO.CopyFile (App.Path & "/studentdatabase.mdb"), (Directory.Path & "/" & FolderName.Text & "/studentdatabase.mdb")
Else
MkDir (Directory.Path & "/" & FolderName.Text)
FSO.CopyFile (App.Path & "/studentdatabase.mdb"), (Directory.Path & "/" & FolderName.Text & "/studentdatabase.mdb")
End If
MsgBox ("Backup successful.")
End Sub

Private Sub Drive_Change()
Directory.Path = Drive.Drive
End Sub

