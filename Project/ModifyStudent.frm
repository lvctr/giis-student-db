VERSION 5.00
Begin VB.Form ModifyStudent 
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   855
      Left            =   600
      TabIndex        =   8
      Top             =   6360
      Width           =   2895
   End
   Begin VB.CommandButton ModifyButton 
      Caption         =   "Modify"
      Height          =   855
      Left            =   600
      TabIndex        =   7
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox StudentID 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Text            =   "20"
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox StudentName 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "Testing"
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox DoB 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "1/1/2007"
      Top             =   3360
      Width           =   3615
   End
   Begin VB.TextBox EnrollDate 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "2/2/2007"
      Top             =   4800
      Width           =   3615
   End
   Begin VB.ComboBox Class 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Text            =   "Select an item..."
      Top             =   1920
      Width           =   3615
   End
   Begin VB.ComboBox Section 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "Select an item..."
      Top             =   2640
      Width           =   3615
   End
   Begin VB.ComboBox Gender 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Select an item..."
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Student ID"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Student Name"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Class"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Section"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Gender"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "Enrollment Date"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   "Date of Birth (format: M/D/YYYY)"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   2895
   End
End
Attribute VB_Name = "ModifyStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
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

Private Sub CancelButton_Click()
Unload Me
End Sub
