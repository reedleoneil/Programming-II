VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Midterm Exam"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16815
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   16815
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Check Room Availability"
      Height          =   1455
      Left            =   120
      TabIndex        =   21
      Top             =   4680
      Width           =   4095
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   840
         TabIndex        =   27
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   840
         TabIndex        =   25
         Top             =   600
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Check"
         Height          =   375
         Left            =   2160
         TabIndex        =   23
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Time:"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Room #:"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Reserve Room"
      Height          =   2175
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   4095
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   840
         TabIndex        =   28
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   26
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   840
         TabIndex        =   20
         Top             =   960
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Reserve"
         Height          =   375
         Left            =   2160
         TabIndex        =   17
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label11 
         Caption         =   "Time Out:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Time In:"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "CID:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Room #:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid dgDefault 
      Height          =   6015
      Left            =   4320
      TabIndex        =   6
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   10610
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Customers"
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
   Begin VB.Frame Frame1 
      Caption         =   "Insert Customer"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   840
         TabIndex        =   4
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   3
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Insert"
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label7 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Contact:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "CID:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6015
      Left            =   10560
      TabIndex        =   13
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   10610
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Room Reservations"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   495
      Left            =   720
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim qry, feilds As String
Dim orderon As Boolean
Private Sub checkres()
Dim x, y, a, b As Date
Call ConnectTable
    rs.Open "Select TimeIn, TimeOut From Reservations WHERE Room = '" & Me.Text4.Text & "'", db
    Set Me.DataGrid2.DataSource = rs
    
    
   x = Me.DataGrid2.Columns(0).Text
    y = Me.DataGrid2.Columns(1).Text
  a = DateDiff("h", x, Me.Text6)
  b = DateDiff("h", y, Me.Text6)
    
    If a >= 0 And b <= 0 Then
        msg = MsgBox("Room " & Me.Text4.Text & " already reserved from " & x & " to " & y & "!", vbInformation, "Info")
    Else
        msg = MsgBox("Room " & Me.Text4.Text & " is available!", vbInformation, "Info")
    End If
End Sub
Private Sub Command1_Click()
    Dim x As String
     x = Now
    qry = "INSERT INTO Reservations (Room, CID, TimeIn, TimeOut) Values ('" & Me.Text2.Text & "','" & Me.Text3.Text & "','" & Me.Text5.Text & "','" & Me.Text7.Text & "')"
    
    db.Execute qry
    Call ConnectTable
    rs.Open "SELECT * FROM Reservations", db
    Set Me.DataGrid1.DataSource = rs
End Sub

Private Sub Command3_Click()
    Call checkres
End Sub




Private Sub Form_Load()
    Call ConnectDB
    
    
    Call SetDg
    
    
End Sub

    


'insert
Private Sub Command2_Click()

    'Insert
    
     
    
    feilds = "(CID, Name, Address, Contact) "
   
        qry = "INSERT INTO " + "Customers" + feilds + " VALUES ('" & Me.Text1(0).Text & "', '" & Me.Text1(1).Text & "', '" & Me.Text1(2).Text & "', '" & Me.Text1(3).Text & "')"
    
    
    
 
    db.Execute qry
   
    Call SetDg
    
    
    
   
     
  
    
End Sub

'functions
Public Sub SetDg()

    'set datagrid
 
    Call ConnectTable
    rs.Open "SELECT * FROM Customers", db
    Set dgDefault.DataSource = rs
    
    Call ConnectTable
    rs.Open "SELECT * FROM Reservations", db
    Set Me.DataGrid1.DataSource = rs
End Sub




