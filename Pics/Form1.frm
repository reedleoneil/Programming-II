VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1800
      TabIndex        =   18
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtContact 
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtCompany 
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   2160
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog cdlCustomers 
      Left            =   6000
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2775
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4895
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "cid"
         Caption         =   "ID Number"
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
         DataField       =   "cname"
         Caption         =   "Customer's Name"
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
      BeginProperty Column02 
         DataField       =   "caddress"
         Caption         =   "Address"
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
      BeginProperty Column03 
         DataField       =   "ccontact"
         Caption         =   "Contact"
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
      BeginProperty Column04 
         DataField       =   "ccompany"
         Caption         =   "Company"
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
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Contact:"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Company:"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Address:"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer Name:"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ID:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Call InitRS
    rs.Open "SELECT * FROM customers WHERE cid = '" & Me.txtID.Text & "'", db

    If rs.RecordCount = 0 Then
    Set msstream = New ADODB.Stream
    msstream.Type = adTypeBinary
    msstream.Open
    
    Call InitRS
    rs.Open "SELECT * FROM customers", db
    
    With rs
            .AddNew
                .Fields(0) = Me.txtID.Text
                .Fields(1) = Me.txtName.Text
                .Fields(2) = Me.txtAddress.Text
                .Fields(4) = Me.txtContact.Text
                .Fields(5) = Me.txtCompany.Text
                
                If Me.cdlCustomers.FileName = "" Then
                Else
                     msstream.LoadFromFile Me.cdlCustomers.FileName
                .Fields(3).Value = msstream.Read
                End If
                
                
            .Update
            
        End With
        Call RefGrid
        Else
            msg = MsgBox("Duplicate ID")
        End If
        End Sub

Private Sub cmdDelete_Click()
        Call InitRS
        rs.Open "SELECT * FROM customers WHERE cid = '" & Me.txtID.Text & "' ", db
        If rs.RecordCount <> 0 Then
            msg = MsgBox("Are you sure you want delete?", vbYesNo + vbQuestion)
            
            If msg = vbYes Then
                rs.Delete
                Call ClearForm
                Call RefGrid
            Else
            End If
            
        Else
        End If
End Sub

Private Sub cmdEdit_Click()
    Set msstream = New ADODB.Stream
    msstream.Type = adTypeBinary
    msstream.Open
    
    Call InitRS
        rs.Open "SELECT * FROM customers WHERE cid = '" & Me.txtID.Text & "' ", db
        If rs.RecordCount <> 0 Then
            'put editing here
            With rs
                
                .Fields(0) = Me.txtID.Text
                .Fields(1) = Me.txtName.Text
                .Fields(2) = Me.txtAddress.Text
                .Fields(4) = Me.txtContact.Text
                .Fields(5) = Me.txtCompany.Text
                
                If Me.cdlCustomers.FileName = "" Then
                Else
                     msstream.LoadFromFile Me.cdlCustomers.FileName
                .Fields(3).Value = msstream.Read
                End If
                
                
            .Update
            End With
        
        
        Else
        End If
        Call RefGrid
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    Set msstream = New ADODB.Stream
    msstream.Type = adTypeBinary
    msstream.Open
    
    Call InitRS
    rs.Open "SELECT * FROM customers", db
    
    If isEdit = False Then    'add mode
        With rs
            .AddNew
                .Fields(0) = Me.txtID.Text
                .Fields(1) = Me.txtName.Text
                .Fields(2) = Me.txtAddress.Text
                
                If Me.cdlCustomers.FileName = "" Then
                Else
                     msstream.LoadFromFile Me.cdlCustomers.FileName
                .Fields(3).Value = msstream.Read
                End If
                
                
            .Update
            
        End With
    
    Else 'edit mode
        Call InitRS
        rs.Open "SELECT * FROM customers WHERE cid = '" & Me.txtID.Text & "' ", db
        If rs.RecordCount <> 0 Then
            'put editing here
            With rs
                
                .Fields(0) = Me.txtID.Text
                .Fields(1) = Me.txtName.Text
                .Fields(2) = Me.txtAddress.Text
                
                If Me.cdlCustomers.FileName = "" Then
                Else
                     msstream.LoadFromFile Me.cdlCustomers.FileName
                .Fields(3).Value = msstream.Read
                End If
                
                
            .Update
            End With
        
        
        Else
        End If
    
    End If
    Call RefGrid
   
    
    
End Sub

Private Sub Command6_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Me.txtAddress.Text = Empty
    Me.txtID.Text = Empty
    Me.txtCompany.Text = Empty
    Me.txtContact.Text = Empty
    Me.txtName.Text = Empty
    Me.Image1 = Nothing
End Sub

Private Sub Command7_Click()
    With Me.cdlCustomers
        .DialogTitle = "Select Customer's Picture"
        .InitDir = "C:\"
        .Filter = "JPEGs|*.jpg|GIFs|*.gif|BITMAPs|*.bmp"
        .FilterIndex = 0
        .ShowOpen
        
        'load picture to the imagebox
        
        Me.Image1.Picture = LoadPicture(.FileName)
        
        
    End With
End Sub

Private Sub cmdFind_Click()

    Set msstream = New ADODB.Stream
    msstream.Type = adTypeBinary
    msstream.Open
    
    Call InitRS
    rs.Open "SELECT * FROM customers WHERE cid = '" & Me.txtID.Text & "'", db

    If rs.RecordCount <> 0 Then
        With rs
            Me.txtID.Text = rs.Fields(0)
            Me.txtName.Text = rs.Fields(1)
            Me.txtAddress.Text = rs.Fields(2)
            Me.txtContact.Text = rs.Fields(4)
            Me.txtCompany.Text = rs.Fields(5)
        
            msstream.Write rs.Fields(3).Value
            msstream.SaveToFile "C:\pic.jpg", adSaveCreateOverWrite
        
            Me.Image1.Picture = LoadPicture("C:\pic.jpg")
        
        End With
      
    Else
        msg = MsgBox("Walang nakita.")
    End If
End Sub

Private Sub Form_Load()
    Call InitDB
    Call RefGrid
End Sub

Private Sub RefGrid()
    Call InitRS
    rs.Open "SELECT * FROM customers", db
    Set Me.DataGrid1.DataSource = rs
End Sub
Private Sub ClearForm()
    With Me
        .txtID.Text = ""
        .txtAddress.Text = ""
        .txtName.Text = ""
        .Image1.Picture = Nothing
    End With
End Sub



Private Sub txtAddress_LostFocus()
    Me.txtAddress.Text = Trim(StrConv(Me.txtAddress.Text, vbProperCase))
End Sub

Private Sub txtCompany_LostFocus()
    Me.txtCompany.Text = Trim(StrConv(Me.txtCompany.Text, vbProperCase))
End Sub

Private Sub txtContact_LostFocus()
    If IsNumeric(Me.txtContact.Text) = False Then
        msg = MsgBox("Check contact.")
    End If
End Sub


Private Sub txtName_LostFocus()
    Me.txtName.Text = Trim(StrConv(Me.txtName.Text, vbProperCase))
End Sub
