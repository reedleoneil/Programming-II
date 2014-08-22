VERSION 5.00
Begin VB.Form Form0 
   Caption         =   "Select Version"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Version 2"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Version 1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Which version do you want to use?"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load Form1
Form1.Enabled = True
Form1.Show
Unload Me
End Sub

Private Sub Command2_Click()
Load Form2
Form2.Enabled = True
Form2.Show
Unload Me
End Sub
