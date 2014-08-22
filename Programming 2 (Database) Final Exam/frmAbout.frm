VERSION 5.00
Begin VB.Form FormAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4560
   ClientLeft      =   3360
   ClientTop       =   2925
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3147.393
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Warning:"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   5535
      Begin VB.Label lblDisclaimer 
         Caption         =   "This software is protected by reed. Copying, distributing, editng, modifying or reverse engineering is punishable by reed."
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   4080
      Width           =   1260
   End
   Begin VB.Label Label3 
      Caption         =   $"frmAbout.frx":030A
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   $"frmAbout.frx":042C
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":04E7
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label lblDescription 
      Caption         =   "Source code is completly from scratch."
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5565
   End
   Begin VB.Label lblTitle 
      Caption         =   "Reed Database Management System"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 3.0"
      Height          =   225
      Left            =   720
      TabIndex        =   4
      Top             =   360
      Width           =   3885
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()                                                                                          'start here >>>
    Me.lblDescription.Caption = "Source code is completely from scratch."
End Sub

