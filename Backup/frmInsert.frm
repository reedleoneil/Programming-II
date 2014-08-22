VERSION 5.00
Begin VB.Form FormInsert 
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   13785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CommandCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6960
      TabIndex        =   9
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton CommandInsert 
      Caption         =   "Insert"
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   3000
      Width           =   2175
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
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   4680
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   4680
      TabIndex        =   2
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   4680
      TabIndex        =   3
      Top             =   2520
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   9240
      TabIndex        =   4
      Top             =   360
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   9240
      TabIndex        =   5
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   9240
      TabIndex        =   6
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   9240
      TabIndex        =   7
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Image ImageInsert 
      BorderStyle     =   1  'Fixed Single
      Height          =   3375
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label01 
      Height          =   255
      Left            =   4680
      TabIndex        =   17
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label02 
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label03 
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label06 
      Height          =   255
      Left            =   9240
      TabIndex        =   14
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label05 
      Height          =   255
      Left            =   9240
      TabIndex        =   13
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label04 
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label Label08 
      Height          =   255
      Left            =   9240
      TabIndex        =   11
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label Label07 
      Height          =   255
      Left            =   9240
      TabIndex        =   10
      Top             =   1560
      Width           =   4455
   End
End
Attribute VB_Name = "FormInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub InsertProductCustomer()
        If FormMain.TabStripRecordSets.SelectedItem = "Products" Then
            x = "Product"
            Qry = "INSERT INTO " + FormMain.TabStripRecordSets.SelectedItem + "(ProductID, ProductName, Model, Manufacturer, Specifications, Type, Price, Warranty, Stocks) VALUES ('" & Me.Text1(0).Text & "', '" & Me.Text1(1).Text & "', '" & Me.Text1(2).Text & "', '" & Me.Text1(3).Text & "', '" & Me.Text1(4).Text & "', '" & Me.Text1(5).Text & "', '" & Me.Text1(6).Text & "', '" & Me.Text1(7).Text & "', 0)"
            backupQry = "INSERT INTO Backup" + FormMain.TabStripRecordSets.SelectedItem + "(ProductID, ProductName, Model, Manufacturer, Specifications, Type, Price, Warranty, Stocks) VALUES ('" & Me.Text1(0).Text & "', '" & Me.Text1(1).Text & "', '" & Me.Text1(2).Text & "', '" & Me.Text1(3).Text & "', '" & Me.Text1(4).Text & "', '" & Me.Text1(5).Text & "', '" & Me.Text1(6).Text & "', '" & Me.Text1(7).Text & "', 0)"
        ElseIf FormMain.TabStripRecordSets.SelectedItem = "Customers" Then
            x = "Customer"
            Qry = "INSERT INTO " + FormMain.TabStripRecordSets.SelectedItem + "(CustomerID, CustomerName, Address, Contact, Discount) VALUES ('" & Me.Text1(0).Text & "', '" & Me.Text1(1).Text & "', '" & Me.Text1(2).Text & "', '" & Me.Text1(3).Text & "', 0)"
            backupQry = "INSERT INTO Backup" + FormMain.TabStripRecordSets.SelectedItem + "(CustomerID, CustomerName, Address, Contact, Discount) VALUES ('" & Me.Text1(0).Text & "', '" & Me.Text1(1).Text & "', '" & Me.Text1(2).Text & "', '" & Me.Text1(3).Text & "', 0)"
        End If
        
        FormMain.TextQuery.Text = Qry
        Dtbs.Execute Qry
        Dtbs.Execute backupQry
   
        Call CnnctRcrdSt
        Qry = "SELECT * FROM " + FormMain.TabStripRecordSets.SelectedItem
        RcrdSt.Open Qry, Dtbs
        Set FormMain.DataGridDefault.DataSource = RcrdSt
   
        Call CnnctRcrdSt
        Qry = "SELECT * FROM " + FormMain.TabStripRecordSets.SelectedItem + " WHERE " + x + "ID = '" & Me.Text1(0).Text & "'"
        RcrdSt.Open Qry, Dtbs
        Set FormMain.DataGridDummy.DataSource = RcrdSt
        
        If Me.Caption = "Insert Product" Then
            i = 8
        ElseIf Me.Caption = "Insert Customer" Then
            i = 4
        End If
        
        For n = 0 To i
            FormMain.Text1(n).Text = FormMain.DataGridDummy.Columns(n).Text
        Next n
        
        Unload Me
        
        FormMain.CommandUpdate.Enabled = True
        FormMain.CommandDelete.Enabled = True
        FormMain.TabStripRecordSets.Enabled = True
End Sub

Private Sub CommandCancel_Click()
    Unload Me
    FormMain.CommandUpdate.Enabled = True
    FormMain.CommandDelete.Enabled = True
    FormMain.TabStripRecordSets.Enabled = True
End Sub

Private Sub CommandInsert_Click()
    'check if text1 is empty
    If Me.Caption = "Insert Product" Then
        b = 7
    ElseIf Me.Caption = "Insert Customer" Then
        b = 3
    End If
    
    For a = 1 To b
        If Me.Text1(a) = Empty Then
            If b = 7 Then
                Select Case a
                    Case 1
                        c = c + " Name,"
                    Case 2
                        c = c + " Model,"
                    Case 3
                        c = c + " Manufacturer,"
                    Case 4
                        c = c + " Specifications,"
                    Case 5
                        c = c + " Type,"
                    Case 6
                        c = c + " Price,"
                    Case 7
                        c = c + " Warranty,"
                End Select
            ElseIf b = 3 Then
                Select Case a
                    Case 1
                        c = c + " Name,"
                    Case 2
                        c = c + " Address,"
                    Case 3
                        c = c + " Contact Number,"
                End Select
            End If
        End If
    Next a
    
    If Me.Caption = "Insert Product" Then
        If Me.Text1(1) = Empty Or Me.Text1(2) = Empty Or Me.Text1(3) = Empty Or Me.Text1(4) = Empty Or Me.Text1(5) = Empty Or Me.Text1(6) = Empty Or Me.Text1(7) = Empty Then
            msg = MsgBox(Left("Enter" + c, Len("Enter" + c) - 1) + ".", vbInformation, Me.Caption)
        Else
            Call InsertProductCustomer
        End If
    ElseIf Me.Caption = "Insert Customer" Then
        If Me.Text1(1) = Empty Or Me.Text1(2) = Empty Or Me.Text1(3) = Empty Then
            msg = MsgBox(Left("Enter" + c, Len("Enter" + c) - 1) + ".", vbInformation, Me.Caption)
        Else
            Call InsertProductCustomer
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormMain.CommandDelete.Enabled = True
    FormMain.CommandUpdate.Enabled = True
    FormMain.TabStripRecordSets.Enabled = True
End Sub

Private Sub ImageInsert_Click()
    With FormMain.CommonDialog1
        .DialogTitle = "Select Photos"
        .InitDir = "C:\"
        .Filter = "JPEGs|*.jpg|GIFs|*.gif|Bitmaps|*.bmp|All Files|*.*"
        .FilterIndex = 1
        .ShowOpen
        Me.ImageInsert.Picture = LoadPicture(.FileName)
    End With
End Sub
