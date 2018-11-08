VERSION 5.00
Begin VB.Form formRestore 
   Caption         =   "Restore"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox FileSelect 
      Height          =   2040
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   4815
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
   Begin VB.DirListBox Directory 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
   Begin VB.CommandButton Restore 
      Caption         =   "Restore"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please make sure the name of the file is ""studentdatabase.mdb"""
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   4815
   End
End
Attribute VB_Name = "formRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Directory_Change()
FileSelect.Path = Directory.Path
End Sub

Private Sub Drive_Change()
Directory.Path = Drive.Drive
End Sub

Private Sub Form_Load()
FileSelect.Path = Directory.Path
FileSelect.FileName = "studentdatabase.mdb"
End Sub

Private Sub Restore_Click()

'Deletes the existing database and copies a backed up copy from another location into the database.

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
db.Close

Kill (App.Path & "/studentdatabase.mdb")
FileCopy (FileSelect.Path & "/studentdatabase.mdb"), (App.Path & "/studentdatabase.mdb")
MsgBox ("Restore complete.")
    
End Sub
