VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   9120
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Update"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Backup"
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Insert"
      Height          =   375
      Left            =   7440
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5040
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   720
      Width           =   1935
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   4920
      TabIndex        =   3
      Top             =   5040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   8535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Execute"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dgEmployees 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call ConnectTable
    
    rs.Open "SELECT custid AS ID, custname AS Name, custaddress AS Address  FROM customers UNION SELECT * FROM suppliers", db
   
    Set Me.dgEmployees.DataSource = rs
    
End Sub

Private Sub Command2_Click()
    'db.Execute "INSERT INTO customers (custid,custname, custaddress) VALUES ('" & Me.Text2.Text & "','" & Me.Text3.Text & "','" & Me.Text4.Text & "')"
    
    Call ConnectTable
    rs.Open "INSERT INTO customers (custid,custname, custaddress) VALUES ('" & Me.Text2.Text & "','" & Me.Text3.Text & "','" & Me.Text4.Text & "')", db
    
    Call ConnectTable
    rs.Open "SELECT * FROM customers", db
    Set Me.dgEmployees.DataSource = rs

End Sub

Private Sub Command3_Click()
    db.Execute "INSERT INTO cbackup SELECT * FROM customers"
    msg = MsgBox("Nalipat na.")
End Sub

Private Sub Command4_Click()
    db.Execute "UPDATE customers SET custname = '" & Me.Text3.Text & "', custaddress ='" & Me.Text4.Text & "' WHERE custid = '" & Me.Text2.Text & "'"

    Call ConnectTable
    rs.Open "SELECT * FROM customers", db
    Set Me.dgEmployees.DataSource = rs
    
End Sub

Private Sub Command5_Click()
    db.Execute "DELETE FROM cbackup WHERE custaddress = 'Manila'"
End Sub

Private Sub dgEmployees_DblClick()
    With Me.dgEmployees
        Me.Text2.Text = .Columns(0).Text
        Me.Text3.Text = .Columns(1).Text
        Me.Text4.Text = .Columns(2).Text
    End With
End Sub

Private Sub Form_Load()
    Call ConnectDB
    
    Call ConnectTable
    rs.Open "SELECT DISTINCT (address) FROM employees WHERE NOT ISNULL(address) ORDER BY address ASC", db
    
    Me.DataCombo1.ListField = "address"
    Set Me.DataCombo1.RowSource = rs
    
    Call ConnectTable
    rs.Open "SELECT * FROM customers", db
    Set Me.dgEmployees.DataSource = rs
    
End Sub

Private Sub Text1_Change()
    Call ConnectTable
    rs.Open "SELECT * FROM employees WHERE empname LIKE  '%" & Me.Text1.Text & "%'", db
    Set Me.dgEmployees.DataSource = rs
    
End Sub
