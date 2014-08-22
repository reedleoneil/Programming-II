VERSION 5.00
Begin VB.Form FormDelete 
   Caption         =   "Delete Confirmation"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   13785
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   8
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2520
      Width           =   4455
   End
   Begin VB.CommandButton CommandNo 
      Caption         =   "No"
      Height          =   495
      Left            =   6960
      TabIndex        =   17
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton CommandYes 
      Caption         =   "Yes"
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3240
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Image ImageUpdate 
      BorderStyle     =   1  'Fixed Single
      Height          =   3375
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label09 
      Height          =   255
      Left            =   9240
      TabIndex        =   20
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Label Label07 
      Height          =   255
      Left            =   9240
      TabIndex        =   15
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label08 
      Height          =   255
      Left            =   9240
      TabIndex        =   14
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label04 
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label Label05 
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   3000
      Width           =   4455
   End
   Begin VB.Label Label06 
      Height          =   255
      Left            =   9240
      TabIndex        =   11
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label03 
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label02 
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label01 
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "FormDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandNo_Click()
    Unload Me
    FormMain.CommandInsert.Enabled = True
    FormMain.CommandUpdate.Enabled = True
    FormMain.TabStripRecordSets.Enabled = True
End Sub

Private Sub CommandYes_Click()
    If FormMain.TabStripRecordSets.SelectedItem = "Products" Then
        x = "Product"
    ElseIf FormMain.TabStripRecordSets.SelectedItem = "Customers" Then
        x = "Customer"
    End If

    Qry = "DELETE FROM " + FormMain.TabStripRecordSets.SelectedItem + " WHERE " + x + "ID = '" & Me.Text1(0).Text & "'"
    FormMain.TextQuery.Text = Qry
    Dtbs.Execute Qry
    
    Call CnnctRcrdSt
    Qry = "SELECT * FROM " + FormMain.TabStripRecordSets.SelectedItem
    RcrdSt.Open Qry, Dtbs
    Set FormMain.DataGridDefault.DataSource = RcrdSt
    
    For n = 0 To 8
        FormMain.Text1(n).Text = Empty
    Next n
    
    Unload Me
    
    FormMain.CommandUpdate.Enabled = True
    FormMain.CommandInsert.Enabled = True
    FormMain.TabStripRecordSets.Enabled = True
End Sub

Private Sub Form_Load()
    Me.Label1.Caption = "Are you sure you want to delete this " + Left(FormMain.TabStripRecordSets.SelectedItem.Caption, Len(FormMain.TabStripRecordSets.SelectedItem.Caption) - 1) + "?"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormMain.CommandInsert.Enabled = True
    FormMain.CommandUpdate.Enabled = True
    FormMain.TabStripRecordSets.Enabled = True
End Sub
