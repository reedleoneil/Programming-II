VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Calculator"
   ClientHeight    =   4215
   ClientLeft      =   3015
   ClientTop       =   2205
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdsolve 
      Caption         =   "&Solve"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Answer"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   3495
      Begin VB.TextBox txtanswer 
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Answer:"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Option"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmdclear 
         Caption         =   "&Clear"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   3015
      End
      Begin VB.OptionButton optdivide 
         Caption         =   "Divide"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton optmultiply 
         Caption         =   "Multiply"
         Height          =   495
         Left            =   1680
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton optminus 
         Caption         =   "Subtract"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   2295
      End
      Begin VB.OptionButton optadd 
         Caption         =   "Add"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtnum2 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtnum1 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Second Number:"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "First Number:"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msg As String

Private Sub cmdclear_Click()
txtnum1.Text = " "
txtnum2.Text = " "
optadd.Value = False
optminus.Value = False
optmultiply.Value = False
optdivide.Value = False
txtanswer.Text = " "
End Sub

Private Sub cmdexit_Click()
msg = MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, " Exit Confirmation")
If msg = vbYes Then
Unload Me
Else

End If

End Sub

Private Sub cmdsolve_Click()
If optadd = True Then
txtanswer = Val(txtnum1.Text) + Val(txtnum2.Text)
ElseIf optminus = True Then
txtanswer = Val(txtnum1.Text) - Val(txtnum2.Text)
ElseIf optmultiply = True Then
txtanswer = Val(txtnum1.Text) * Val(txtnum2.Text)
ElseIf optdivide = True Then
txtanswer = Val(txtnum1.Text) / Val(txtnum2.Text)
ElseIf (txtnum1.Text = " " Or txtnum2.Text = "") Or (txtnum1.Text = " " And txtnum2.Text = " ") Then
msg = MsgBox("Input numbers.", vbExclamation, "Message!")
Else
msg = MsgBox("Select an option.", vbExclamation, "Message!")

End If
End Sub




