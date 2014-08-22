VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Midterm Exam"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   12600
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   975
      Left            =   960
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1720
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
      Caption         =   "Dummy"
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   1815
      Left            =   120
      TabIndex        =   34
      Top             =   2400
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3201
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
      Caption         =   "Rooms "
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
   Begin VB.Frame Frame3 
      Caption         =   "Check Room Availability"
      Height          =   1815
      Left            =   8400
      TabIndex        =   24
      Top             =   2400
      Width           =   4095
      Begin VB.OptionButton Option2 
         Caption         =   "PM"
         Height          =   255
         Left            =   3360
         TabIndex        =   36
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "AM"
         Height          =   255
         Left            =   2640
         TabIndex        =   35
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   840
         TabIndex        =   33
         Text            =   "MM/DD/YY"
         Top             =   960
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Check"
         Height          =   375
         Left            =   2280
         TabIndex        =   27
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   840
         TabIndex        =   26
         Text            =   "HH:MM"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   840
         TabIndex        =   25
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label13 
         Caption         =   "Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Room #:"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Time:"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Reserve Room"
      Height          =   2535
      Left            =   8400
      TabIndex        =   12
      Top             =   4320
      Width           =   4095
      Begin VB.OptionButton Option6 
         Caption         =   "PM"
         Height          =   255
         Left            =   3360
         TabIndex        =   40
         Top             =   1320
         Width           =   615
      End
      Begin VB.OptionButton Option5 
         Caption         =   "AM"
         Height          =   255
         Left            =   2640
         TabIndex        =   39
         Top             =   1320
         Width           =   615
      End
      Begin VB.OptionButton Option4 
         Caption         =   "PM"
         Height          =   255
         Left            =   3360
         TabIndex        =   38
         Top             =   960
         Width           =   615
      End
      Begin VB.OptionButton Option3 
         Caption         =   "AM"
         Height          =   255
         Left            =   2640
         TabIndex        =   37
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   840
         TabIndex        =   31
         Text            =   "MM/DD/YY"
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   840
         TabIndex        =   22
         Text            =   "HH:MM"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   21
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   840
         TabIndex        =   20
         Text            =   "HH:MM"
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Reserve"
         Height          =   375
         Left            =   2160
         TabIndex        =   17
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label12 
         Caption         =   "Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Time Out:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
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
   Begin VB.Frame Frame1 
      Caption         =   "Insert Customer"
      Height          =   2175
      Left            =   8400
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
      Height          =   2535
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4471
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
   Begin MSDataGridLib.DataGrid dgDefault 
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3836
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
    rs.Open "Select TimeIn, TimeOut, ResDate From Reservations WHERE Room = '" & Me.Text4.Text & "'", db
    Set Me.DataGrid2.DataSource = rs
    
    
   x = Me.DataGrid2.Columns(0).Text
    y = Me.DataGrid2.Columns(1).Text
  a = DateDiff("h", x, Me.Text6)
  b = DateDiff("h", y, Me.Text6)
    
    If (a >= 0 And b <= 0) And Me.Text9.Text = Me.DataGrid2.Columns(2).Text Then
        msg = MsgBox("Room " & Me.Text4.Text & " already reserved on " + Me.DataGrid2.Columns(2).Text + " from " & x & " to " & y & "!", vbInformation, "Info")
    Else
        msg = MsgBox("Room " & Me.Text4.Text & " is available on " & Me.Text9.Text & " " & Me.Text6.Text & "!", vbInformation, "Info")
    End If
End Sub
Private Sub Command1_Click()
    Dim x As String
     x = Now
    qry = "INSERT INTO Reservations (Room, CID, TimeIn, TimeOut, ResDate) Values ('" & Me.Text2.Text & "','" & Me.Text3.Text & "','" & Me.Text5.Text & "','" & Me.Text7.Text & "', '" & Me.Text8.Text & "')"
    
    db.Execute qry
    Call ConnectTable
    rs.Open "SELECT * FROM Reservations", db
    Set Me.DataGrid1.DataSource = rs
End Sub

Private Sub Command3_Click()
    Call checkres
End Sub




Private Sub DataGrid3_Click()
    Me.Text2.Text = Me.DataGrid3.Columns(0).Text
    Me.Text4.Text = Me.DataGrid3.Columns(0).Text
End Sub

Private Sub dgDefault_Click()
    Me.Text3.Text = Me.dgDefault.Columns(0).Text
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
    
    Call ConnectTable
    rs.Open "SELECT * FROM Rooms", db
    Set Me.DataGrid3.DataSource = rs
End Sub



Private Sub Option1_Click()
    Me.Text6.Text = Left(Me.Text6.Text, 5) & " AM"
End Sub

Private Sub Option2_Click()
Me.Text6.Text = Left(Me.Text6.Text, 5) & " PM"
End Sub

Private Sub Option3_Click()
    Me.Text5.Text = Left(Me.Text5.Text, 5) & " AM"
End Sub

Private Sub Option4_Click()
Me.Text5.Text = Left(Me.Text5.Text, 5) & " PM"
End Sub

Private Sub Option5_Click()
Me.Text7.Text = Left(Me.Text7.Text, 5) & " AM"
End Sub

Private Sub Option6_Click()
Me.Text7.Text = Left(Me.Text7.Text, 5) & " PM"
End Sub
