VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1725
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   1725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Command1_Click()
Do While i <= 20
    If i Mod 2 = 0 Then
        Print i
    Else
    
    End If
    i = i + 1
Loop
End Sub
