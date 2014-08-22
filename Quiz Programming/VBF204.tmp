VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1560
      TabIndex        =   16
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1560
      TabIndex        =   15
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   14
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   13
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   12
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   11
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   10
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Salary"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Position"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Conpany"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Birthday"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Contact"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Addess"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "CID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y
Private Sub Text1_Change(Index As Integer)
    If Index = 3 Then
    If IsNumeric(Me.Text1(3).Text) = True Then
        Me.Text1(3).Text = x
    Else
        Me.Text1(3).Text = Empty
        Me.Text1(3).Text = x
    End If
    End If
    
    If Index = 7 Then
    If IsNumeric(Me.Text1(7).Text) = True Then
        Me.Text1(7).Text = y
    Else
        Me.Text1(7).Text = Empty
        Me.Text1(7).Text = y
    End If
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = True And Index = 3 Then
        x = Me.Text1(3).Text & Chr(KeyAscii)
    Else
        
    End If
    If IsNumeric(Chr(KeyAscii)) = True And Index = 7 Then
        y = Me.Text1(7).Text & Chr(KeyAscii)
    Else
        
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    For n = 0 To 7
        Me.Text1(n).Text = Trim(Me.Text1(n).Text)
    Next n
    
    Me.Text1(1).Text = StrConv(Me.Text1(1).Text, vbProperCase)
    Me.Text1(2).Text = StrConv(Me.Text1(2).Text, vbProperCase)
    Me.Text1(5).Text = StrConv(Me.Text1(5).Text, vbProperCase)
    Me.Text1(6).Text = StrConv(Me.Text1(6).Text, vbProperCase)
    
    'If IsNumeric(Me.Text1(1).Text) = True And Index = 1 Or (IsNumeric(Me.Text1(6).Text) = True And Index = 6) Or (IsNumeric(Me.Text1(5).Text) = True And Index = 5) Then
        'msg = MsgBox(Me.Text1(Index).Text & " contains a number!", vbInformation, "Check Input")
    'End If
    
    If IsDate(Me.Text1(4).Text) = False And Index = 4 Then
        msg = MsgBox("Birthday cannot be " & Me.Text1(4).Text & ", it must be in date format!", vbInformation, "Check Input")
    End If
End Sub
