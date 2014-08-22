VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Prelim Exam"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   13380
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   255
      Left            =   5880
      TabIndex        =   26
      Top             =   3840
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command6 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   25
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Total Price"
      Height          =   615
      Left            =   4800
      TabIndex        =   22
      Top             =   4320
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3855
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4095
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Order"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Order"
      Height          =   555
      Left            =   9120
      TabIndex        =   23
      Top             =   4320
      Width           =   4095
   End
   Begin MSDataGridLib.DataGrid dgDefault 
      Height          =   4455
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search"
      Height          =   1095
      Left            =   9120
      TabIndex        =   19
      Top             =   120
      Width           =   4095
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   3855
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label6 
         Caption         =   "Search By:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   9120
      TabIndex        =   1
      Top             =   1200
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   6
         Top             =   1680
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   5
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   4
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Update"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Insert"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clear"
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1296
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Customers"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Products"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Orders"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim qry, feilds As String
Dim orderon As Boolean
Private Sub Form_Load()
    Call ConnectDB
    
    'Defaults at Load
    orderon = False
    
    Me.Label1.Caption = "CID:"
    Me.Label2.Caption = "Name:"
    Me.Label3.Caption = "Address:"
    Me.Label4.Caption = "Contact:"
    Me.Label5.Caption = Empty
    
    Me.DataCombo1.Text = "CID"
    Call ConnectTable
    rs.Open "SELECT CustomersDataCombo FROM AppData", db
    Set Me.DataCombo1.RowSource = rs
    Me.DataCombo1.ListField = "CustomersDataCombo"
    
    
    Call ConnectTable
    qry = "SELECT * FROM Customers"
    rs.Open "SELECT * FROM Customers", db
    Set dgDefault.DataSource = rs
    
End Sub
'search, insert, update, delete
Private Sub Text2_Change()

    'Search
    Call ConnectTable
    qry = "SELECT * FROM " + Me.TabStrip1.SelectedItem + " WHERE " + Me.DataCombo1.Text + " LIKE  '%" & Me.Text2.Text & "%'"
    rs.Open qry, db
    Set Me.dgDefault.DataSource = rs
    
End Sub
Private Sub Command2_Click()
    
    If Me.TabStrip1.SelectedItem = "Customers" Then
        x = "c"
    ElseIf Me.TabStrip1.SelectedItem = "Products" Then
        x = "p"
    Else
        x = "o"
    End If
    qry = "SELECT " + x + "id FROM " + Me.TabStrip1.SelectedItem + " WHERE " + x + "id = " + "'" + Me.Text1(0).Text + "'"
    Call ConnectTable
    rs.Open qry, db
    Set Me.DataGrid2.DataSource = rs
    
    If Me.DataGrid2.Columns(0).Text = Me.Text1(0).Text Then
        msg = MsgBox("error")
    Else

    'Insert
        If Me.TabStrip1.SelectedItem = "Customers" Or Me.TabStrip1.SelectedItem = "Products" Then
            If Me.TabStrip1.SelectedItem = "Customers" Then
                feilds = "(CID, Name, Address, Contact)"
            ElseIf Me.TabStrip1.SelectedItem = "Products" Then
                feilds = "(PID, Name, Unit, Price)"
            End If
        qry = "INSERT INTO " + Me.TabStrip1.SelectedItem + feilds + " VALUES ('" & Me.Text1(0).Text & "', '" & Me.Text1(1).Text & "', '" & Me.Text1(2).Text & "', '" & Me.Text1(3).Text & "')"
        ElseIf Me.TabStrip1.SelectedItem = "Orders" Then
            feilds = "(OID, CID, PID, Quantity, TotalPrice)"
            qry = "INSERT INTO " + Me.TabStrip1.SelectedItem + feilds + " VALUES ('" & Me.Text1(0).Text & "', '" & Me.Text1(1).Text & "', '" & Me.Text1(2).Text & "', '" & Me.Text1(3).Text & "', '" & Me.Text1(4).Text & "')"
        End If
        db.Execute qry
        Call SetDg
    
    End If
End Sub
Private Sub Command3_Click()

    'Delete
    Select Case Me.TabStrip1.SelectedItem
    Case "Customers"
        x = "C"
    Case "Products"
        x = "P"
    Case "Orders"
        x = "O"
    End Select
    qry = "DELETE FROM " + Me.TabStrip1.SelectedItem + " WHERE " + x + "ID = '" & Me.Text1(0).Text & "'"
    db.Execute qry
    Call SetDg
    
End Sub
Private Sub Command4_Click()

    'Update
    If Me.TabStrip1.SelectedItem = "Customers" Then
        qry = "UPDATE " + Me.TabStrip1.SelectedItem + " SET Name = '" & Me.Text1(1).Text & "', Address = '" & Me.Text1(2).Text & "', Contact = '" & Me.Text1(3).Text & "' WHERE CID = '" & Me.Text1(0).Text & "'"
    ElseIf Me.TabStrip1.SelectedItem = "Products" Then
        qry = "UPDATE " + Me.TabStrip1.SelectedItem + " SET Name = '" & Me.Text1(1).Text & "', Unit = '" & Me.Text1(2).Text & "', Price = '" & Me.Text1(3).Text & "' WHERE PID = '" & Me.Text1(0).Text & "'"
    ElseIf Me.TabStrip1.SelectedItem = "Orders" Then
        qry = "UPDATE " + Me.TabStrip1.SelectedItem + " SET CID = '" & Me.Text1(1).Text & "', PID = '" & Me.Text1(2).Text & "', Quantity = '" & Me.Text1(3).Text & "', TotalPrice = '" & Me.Text1(4).Text & "' WHERE OID = '" & Me.Text1(0).Text & "'"
    End If
    db.Execute qry
    Call SetDg
    
End Sub
'preview pane
Private Sub TabStrip1_Click()

    'set object properties for preview pane
    If orderon = False Then
        Me.Text1(0).Text = Empty
        Me.Text1(1).Text = Empty
        Me.Text1(2).Text = Empty
        Me.Text1(3).Text = Empty
        Me.Text1(4).Text = Empty
    End If
    
    If Me.TabStrip1.SelectedItem = "Customers" Then
        Me.DataCombo1.Text = "CID"
        If orderon = False Then
            Me.Label1.Caption = "CID:"
            Me.Label2.Caption = "Name:"
            Me.Label3.Caption = "Address:"
            Me.Label4.Caption = "Contact:"
            Me.Label5.Caption = Empty
            Me.Text1(4).Visible = False
        End If
    ElseIf Me.TabStrip1.SelectedItem = "Products" Then
        Me.DataCombo1.Text = "PID"
        If orderon = False Then
            Me.Label1.Caption = "PID:"
            Me.Label2.Caption = "Name:"
            Me.Label3.Caption = "Unit:"
            Me.Label4.Caption = "Price:"
            Me.Label5.Caption = Empty
            Me.Text1(4).Visible = False
        End If
    ElseIf Me.TabStrip1.SelectedItem = "Orders" Then
        Me.DataCombo1.Text = "OID"
        If orderon = False Then
            Me.Label1.Caption = "OID:"
            Me.Label2.Caption = "CID:"
            Me.Label3.Caption = "PID:"
            Me.Label4.Caption = "Quantity:"
            Me.Label5.Caption = "Total Price:"
            Me.Text1(4).Visible = True
        End If
    End If
    
    'set searchby datacombo
    Call ConnectTable
    qry = "SELECT " + Me.TabStrip1.SelectedItem + "DataCombo FROM AppData"
    rs.Open qry, db
    Set Me.DataCombo1.RowSource = rs
    Me.DataCombo1.ListField = Me.TabStrip1.SelectedItem + "DataCombo"
    Call SetDg
    
End Sub
Private Sub dgDefault_Click()
    
    'set data for preview pane
    If orderon = False Then
        Select Case Me.TabStrip1.SelectedItem
        Case "Customers"
            y = 3
        Case "Products"
            y = 3
        Case "Orders"
            y = 4
        End Select
        For x = 0 To y
        With Me.dgDefault
            Me.Text1(x).Text = .Columns(x).Text
        End With
        Next x
    ElseIf orderon = True Then
        Select Case Me.TabStrip1.SelectedItem
        Case "Customers"
            Me.Text1(1).Text = Me.dgDefault.Columns(0).Text
        Case "Products"
            Me.Text1(2).Text = Me.dgDefault.Columns(0).Text
        End Select
    End If
    
End Sub
Private Sub Command5_Click()

    'clear preview pane
    For x = 0 To 4
        Me.Text1(x).Text = Empty
    Next x
    
End Sub
'functions
Public Sub SetDg()

    'set datagrid
    Call ConnectTable
    qry = "SELECT * FROM " + Me.TabStrip1.SelectedItem
    rs.Open qry, db
    Set Me.dgDefault.DataSource = rs
    
End Sub
'order
Private Sub Command1_Click()
Dim nextid As String
    
    'start order
    If Me.Command1.Caption = "Order" Then
    
        'set preview pane
        If orderon = False Then
            Me.Text1(0).Text = Empty
            Me.Text1(1).Text = Empty
            Me.Text1(2).Text = Empty
            Me.Text1(3).Text = Empty
            Me.Text1(4).Text = Empty
        ElseIf orderon = True Then
            nextid = Val(Right(Me.Text1(0).Text, 5)) + 1
            If Len(nextid) = 1 Then
                nextid = "0000" + nextid
            ElseIf Len(nextid) = 2 Then
                nextid = "000" + nextid
            ElseIf Len(nextid) = 3 Then
                nextid = "00" + nextid
            ElseIf Len(nextid) = 4 Then
                nextid = "0" + nextid
            End If
            Me.Text1(0).Text = "o" + nextid
            Me.Text1(2).Text = Empty
            Me.Text1(3).Text = Empty
            Me.Text1(4).Text = Empty
        End If
        
        orderon = True
        
        'set objects properties
        Me.Command1.Caption = "Process Order"
        Me.Command2.Enabled = False
        Me.Command3.Enabled = False
        Me.Command4.Enabled = False
        Me.Label1.Caption = "OID:"
        Me.Label2.Caption = "CID:"
        Me.Label3.Caption = "PID:"
        Me.Label4.Caption = "Quantity:"
        Me.Label5.Caption = "Total Price:"
        Me.Text1(4).Visible = True
        Me.DataGrid1.Visible = False
        Me.dgDefault.Visible = True
        Me.Frame3.Visible = False
        Me.Command6.Visible = False
        Me.Frame2.Enabled = True
        
    'process order
    ElseIf Me.Command1.Caption = "Process Order" Then
    
        'set object properties
        Me.Command1.Caption = "Order"
        Me.DataGrid1.Visible = True
        Me.dgDefault.Visible = False
        Me.Frame3.Visible = True
        Me.Command6.Visible = True
        Me.Frame2.Enabled = False
        
        'set datagrid and insert data into database
        qry = "INSERT INTO Orders (OID, CID, PID, Quantity, TotalPrice) VALUES ('" & Me.Text1(0).Text & "', '" & Me.Text1(1).Text & "', '" & Me.Text1(2).Text & "', '" & Me.Text1(3).Text & "', '" & Me.Text1(4).Text & "' )"
        db.Execute qry
        Call ConnectTable
        qry = "SELECT * FROM Orders INNER JOIN Customers ON Orders.CID = Customers.CID WHERE Orders.CID = '" & Me.Text1(1).Text & "'"
        rs.Open qry, db
        Set Me.DataGrid1.DataSource = rs
        
        'set total price
        Call ConnectTable
        rs.Open "SELECT SUM(TotalPrice) FROM Orders WHERE CID = '" & Me.Text1(1).Text & "'", db
        Set Me.DataGrid2.DataSource = rs
        Me.Text3.Text = Me.DataGrid2.Columns(0).Value
        'Set Me.Text3.DataSource = rs
        'Me.Text3.DataField = "TotalPrice"
        End If
        
End Sub
Private Sub Command6_Click()

    'end order
    orderon = False
    Call TabStrip1_Click
    
    'set object properties
    Me.DataGrid1.Visible = False
    Me.dgDefault.Visible = True
    Me.Frame3.Visible = False
    Me.Command6.Visible = False
    Me.Command2.Enabled = True
    Me.Command3.Enabled = True
    Me.Command4.Enabled = True
    Me.Frame2.Enabled = True
    
End Sub
Private Sub Text1_Change(Index As Integer)

    'set total price =  quantity * price
    If orderon = True Then
        If Index = 3 Then
            Call ConnectTable
            rs.Open "SELECT Price FROM Products WHERE PID = '" & Me.Text1(2).Text & " '", db
            Set Me.Text1(4).DataSource = rs
            Me.Text1(4).DataField = "Price"
            Me.Text1(4).Text = Val(Me.Text1(3).Text) * Val(Me.Text1(4).Text)
        End If
    End If
    
End Sub





























