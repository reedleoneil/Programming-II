VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5940
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   5940
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerSplash 
      Interval        =   700
      Left            =   2520
      Top             =   1680
   End
   Begin MSComctlLib.ProgressBar ProgressBarSplash 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   7225
      _ExtentX        =   12753
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "FormSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TimerSplash_Timer()
If TimerSplash.Interval = 700 Then
 Call Splash
Else
End If
End Sub
Private Sub Splash()
For i = 0 To 100 Step 0.01
Me.ProgressBarSplash.Value = i
Next i
Unload Me
FormLogin.Show
End Sub

