VERSION 5.00
Begin VB.Form FormLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1200
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   709
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   2
      Top             =   660
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   3
      Top             =   660
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1050
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   165
      Width           =   2565
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim trialsleft As Integer
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    If txtPassword = " " Then
        Unload Me
        FormMain.Show
    Else
        msg = MsgBox("Invalid password, try again!", vbCritical, "Error")
        txtPassword.SetFocus
    End If
End Sub
