VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1110
   ClientLeft      =   9540
   ClientTop       =   5835
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   6975
   Begin VB.CommandButton LoginButton 
      Caption         =   "Login"
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox UserPWText 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "O"
      TabIndex        =   3
      Text            =   "1785"
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox UserLoginText 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "Admin"
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label LoginPWLabel 
      Caption         =   "Password"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label LoginUNLabel 
      Caption         =   "Username"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoginButton_Click()

Set db = OpenDatabase(App.Path & "/studentdatabase.mdb")
Set ors = db.OpenRecordset("AdminDetails")

With ors
    If UserLoginText.Text = !UserName And UserPWText.Text = !Password Then
        MainForm.Show
        Unload Me
    Else
        MsgBox ("Incorrect Password. Try again.")
    End If
End With

End Sub
