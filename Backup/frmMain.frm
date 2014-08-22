VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormMain 
   Caption         =   "Reed Database Managment System Version 3.0"
   ClientHeight    =   10350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   20250
   Begin VB.TextBox TextQuery 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Width           =   20250
   End
   Begin VB.CommandButton CommandVoidWarranty 
      Caption         =   "Void Warranty"
      Height          =   495
      Left            =   17880
      TabIndex        =   56
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton CommandActivateWarranty 
      Caption         =   "Activate Warranty"
      Height          =   495
      Left            =   15480
      TabIndex        =   55
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton CommandNewTransaction 
      Caption         =   "New Transaction"
      Height          =   495
      Left            =   17880
      TabIndex        =   54
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton CommandAbout 
      Caption         =   "About"
      Height          =   495
      Left            =   17880
      TabIndex        =   46
      Top             =   9720
      Width           =   2295
   End
   Begin VB.CommandButton CommandHelp 
      Caption         =   "Help"
      Height          =   495
      Left            =   15480
      TabIndex        =   45
      Top             =   9720
      Width           =   2295
   End
   Begin VB.Frame FrameStocks 
      Caption         =   "Stocks"
      Height          =   2175
      Left            =   15480
      TabIndex        =   43
      Top             =   6000
      Width           =   4695
      Begin VB.CommandButton CommandRefreshStocks 
         Caption         =   "Refresh Stocks"
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton CommandAddStocks 
         Caption         =   "Add Stocks"
         Height          =   375
         Left            =   2400
         TabIndex        =   44
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label LabelHS 
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   60
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label LabelHS 
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   59
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label LabelLS 
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   58
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label LabelLS 
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   57
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label LabelHS 
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   53
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label LabelLS 
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   52
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label LabelHighStocks 
         Caption         =   "High Stocks:"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   960
         Width           =   975
      End
      Begin VB.Label LabelLowStocks 
         Caption         =   "Low Stocks:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DataGridDummy 
      Height          =   375
      Left            =   10920
      TabIndex        =   38
      Top             =   4560
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin VB.Frame FrameDefault 
      Height          =   4410
      Left            =   4920
      TabIndex        =   40
      Top             =   5160
      Width           =   10410
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9840
         Top             =   3840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSDataGridLib.DataGrid DataGridDefault 
      Height          =   4410
      Left            =   4920
      TabIndex        =   39
      Top             =   720
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   7779
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   14
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
         Name            =   "Arial"
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
   Begin VB.Frame FramePreviewPane 
      Caption         =   "Preview Pane"
      Height          =   10155
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   4695
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   42
         Top             =   9720
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   36
         Top             =   9000
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   33
         Top             =   8280
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   32
         Top             =   7560
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   31
         Top             =   6840
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   6120
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   5400
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   4680
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   3960
         Width           =   4455
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   3375
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label09 
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   9480
         Width           =   4455
      End
      Begin VB.Label Label07 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   8040
         Width           =   4455
      End
      Begin VB.Label Label08 
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   8760
         Width           =   4455
      End
      Begin VB.Label Label04 
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   5880
         Width           =   4455
      End
      Begin VB.Label Label05 
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   6600
         Width           =   4455
      End
      Begin VB.Label Label06 
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   7320
         Width           =   4455
      End
      Begin VB.Label Label03 
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   5160
         Width           =   4455
      End
      Begin VB.Label Label02 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   4440
         Width           =   4455
      End
      Begin VB.Label Label01 
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3720
         Width           =   4455
      End
   End
   Begin VB.CommandButton CommandUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   15480
      TabIndex        =   19
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton CommandDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   17880
      TabIndex        =   18
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Frame FrameSort 
      Caption         =   "Sort"
      Height          =   1215
      Left            =   15480
      TabIndex        =   13
      Top             =   1200
      Width           =   4695
      Begin VB.CommandButton CommandSortDesc 
         Caption         =   "Sort Descending"
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton CommandSortAsc 
         Caption         =   "Sort Ascending"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo DataComboSrtBy 
         Height          =   300
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LabelSortBy 
         Caption         =   "Sort By:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame FrameBackup 
      Caption         =   "Backup"
      Height          =   1455
      Left            =   15480
      TabIndex        =   5
      Top             =   8160
      Width           =   4695
      Begin VB.CommandButton CommandBackupDatabase 
         Caption         =   "Backup Database"
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton CommandRestoreBackup 
         Caption         =   "Restore Backup"
         Height          =   375
         Left            =   2400
         TabIndex        =   47
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label LabelBackupDate 
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label LabelSystemDate 
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label LabelLastBackupDate 
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label1LastBackupDate 
         Caption         =   "Last Backup Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1BackupDate 
         Caption         =   "Backup Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1SystemDate 
         Caption         =   "System Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FrameSearch 
      Caption         =   "Search"
      Height          =   975
      Left            =   15480
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.TextBox TextSearch 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   4455
      End
      Begin MSDataListLib.DataCombo DataComboSrchBy 
         Height          =   300
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LabelSrchBy 
         Caption         =   "Search By:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComctlLib.TabStrip TabStripRecordSets 
      Height          =   1335
      Left            =   4920
      TabIndex        =   4
      Top             =   240
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   2355
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Products"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Customers"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Items"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Warranties"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Transactions"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Orders"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FrameQuery 
      Caption         =   "Query"
      Height          =   615
      Left            =   4920
      TabIndex        =   12
      Top             =   9600
      Width           =   10410
      Begin VB.CommandButton CommandQuery 
         Caption         =   "Query"
         Height          =   255
         Left            =   9360
         TabIndex        =   61
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton CommandInsert 
      Caption         =   "Insert"
      Height          =   495
      Left            =   15480
      TabIndex        =   37
      Top             =   2520
      Width           =   2295
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variable for query
Dim Qry As String
'variable for backup
Dim CurrentDate As Date
'variables for add stock
Dim stckstoadd As String
Dim nxtid As String
Dim inpt As String





Private Sub CommandOpenDatabase_Click()
    Dim appAccess As Access.Application
    Dim strDB As String

    ' Initialize string to database path.
    Const strConPathToSamples = "D:\Projects\"

    strDB = strConPathToSamples & "Database.mdb"
    ' Create new instance of Microsoft Access.
    Set appAccess = CreateObject("Access.Application")
    ' Open database in Microsoft Access window.
    appAccess.OpenCurrentDatabase strDB
    ' Open Orders form.
    appAccess.DoCmd.OpenTable "Products", acViewNormal, acReadOnly
End Sub
Private Sub CommandQuery_Click()
    msg = MsgBox("Warning: Using manual command line requires skills and knowledge about SQL statements and functions. Manual command line is only for debugging or special commands. Syntax errors can cause application crash or run-time-errors. Do you want to continue?", vbExclamation + vbYesNo, "Query")
    
    If msg = vbYes Then
        If Me.TextQuery.Text <> Empty Then
        Qry = Me.TextQuery.Text
        Call CnnctRcrdSt
        RcrdSt.Open Qry, Dtbs
        
        'set datagrid and preveiw pane
        x = Left(Me.TextQuery.Text, 6)
        If x = "SELECT" Or x = "Select" Or x = "select" Then
            Set Me.DataGridDefault.DataSource = RcrdSt
        ElseIf x = "DELETE" Or x = "Delete" Or x = "delete" Then
            For n = 0 To 8
                Me.Text1(n).Text = Empty
            Next n
            Call SetDataGridDefault
        ElseIf x = "INSERT" Or x = "Insert" Or x = "insert" Then
            If Me.TabStripRecordSets.SelectedItem = "Products" Then
                id = Mid(Me.TextQuery.Text, 123, 6)
            ElseIf Me.TabStripRecordSets.SelectedItem = "Customers" Then
                id = Mid(Me.TextQuery.Text, 86, 6)
            End If
            
            Call CnnctRcrdSt
            RcrdSt.Open "SELECT * FROM " + Me.TabStripRecordSets.SelectedItem + " WHERE " + Left(Me.TabStripRecordSets.SelectedItem.Caption, Len(Me.TabStripRecordSets.SelectedItem.Caption) - 1) + "ID = '" + id + "'", Dtbs
            Set Me.DataGridDefault.DataSource = RcrdSt
            
            Call DataGridDefault_Click
            Call SetDataGridDefault
        ElseIf x = "UPDATE" Or x = "Update" Or x = "update" Then
            id = Left(Right(Me.TextQuery.Text, 7), 6)
        
            Call CnnctRcrdSt
            RcrdSt.Open "SELECT * FROM " + Me.TabStripRecordSets.SelectedItem + " WHERE " + Left(Me.TabStripRecordSets.SelectedItem.Caption, Len(Me.TabStripRecordSets.SelectedItem.Caption) - 1) + "ID = '" + id + "'", Dtbs
            Set Me.DataGridDefault.DataSource = RcrdSt
            
            Call DataGridDefault_Click
            Call SetDataGridDefault
        End If
        Else
            msg = MsgBox("Error: Enter a query!", vbCritical, "Query")
        End If
    ElseIf msg = vbNo Then
    
    End If
End Sub





'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Help and About<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Private Sub CommandAbout_Click()
    FormAbout.Show
End Sub
Private Sub CommandHelp_click()
    msg = MsgBox("For more help and information, email me at reedleoneilpascual@yahoo.com", , "Help")
End Sub





'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>reed-define functions<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Private Sub SetDataGridDefault()
    'set datagrid
    Call CnnctRcrdSt
    Qry = "Select * From " + Me.TabStripRecordSets.SelectedItem
    RcrdSt.Open Qry, Dtbs
    Set Me.DataGridDefault.DataSource = RcrdSt
End Sub
Public Sub BackupDatabase()
    'backup database
    Dtbs.Execute "DELETE FROM BackupProducts"
    Dtbs.Execute "DELETE FROM BackupItems"
    Dtbs.Execute "DELETE FROM BackupCustomers"
    Dtbs.Execute "DELETE FROM BackupTransactions"
    Dtbs.Execute "DELETE FROM BackupWarranties"
    Dtbs.Execute "DELETE FROM BackupOrders"
    Dtbs.Execute "DELETE FROM BackupSecurity"
    Dtbs.Execute "DELETE FROM BackupSecurityLog"
    Dtbs.Execute "DELETE FROM BackupAppData"
    Dtbs.Execute "INSERT INTO BackupProducts SELECT * FROM Products"
    Dtbs.Execute "INSERT INTO BackupItems SELECT * FROM Items"
    Dtbs.Execute "INSERT INTO BackupCustomers SELECT * FROM Customers"
    Dtbs.Execute "INSERT INTO BackupTransactions SELECT * FROM Transactions"
    Dtbs.Execute "INSERT INTO BackupWarranties SELECT * FROM Warranties"
    Dtbs.Execute "INSERT INTO BackupOrders SELECT * FROM Orders"
    Dtbs.Execute "INSERT INTO BackupSecurity SELECT * FROM Security"
    Dtbs.Execute "INSERT INTO BackupSecurityLog SELECT * FROM SecurityLog"
    Dtbs.Execute "INSERT INTO BackupAppData SELECT * FROM AppData"
    msg = MsgBox("Backup database updated.", vbInformation + vbOKOnly, "Backup Database")
End Sub
Public Sub SetFrameBackup()
    'set frame backup
    Call CnnctRcrdSt
    x = "BckupDt"
    Qry = "SELECT " + x + " FROM AppData"
    RcrdSt.Open Qry, Dtbs
    Set Me.LabelBackupDate.DataSource = RcrdSt
    Me.LabelBackupDate.DataField = x
   
    Call CnnctRcrdSt
    y = "LstBckupDt"
    Qry = "SELECT " + y + " FROM AppData"
    RcrdSt.Open Qry, Dtbs
    Set Me.LabelLastBackupDate.DataSource = RcrdSt
    Me.LabelLastBackupDate.DataField = y
End Sub
Public Sub SetFrameStocks()
    'set frame stocks
    
    'set low stocks
    Call CnnctRcrdSt
    RcrdSt.Open "SELECT TOP 1 ProductName, Stocks, ProductID FROM Products ORDER BY Stocks ASC", Dtbs
    Set Me.DataGridDummy.DataSource = RcrdSt
    If Me.DataGridDummy.ApproxCount = 0 Then
    
    Else
        Me.LabelLS(0).Caption = Me.DataGridDummy.Columns(0).Text + " - " + Me.DataGridDummy.Columns(1).Text + " stocks remaining"
        x = Me.DataGridDummy.Columns(2).Text
    End If
    
    Call CnnctRcrdSt
    RcrdSt.Open "SELECT TOP 1 ProductName, Stocks, ProductID FROM Products WHERE ProductID <> '" & x & "' ORDER BY Stocks ASC", Dtbs
    Set Me.DataGridDummy.DataSource = RcrdSt
    If Me.DataGridDummy.ApproxCount = 0 Then
        
    Else
        Me.LabelLS(1).Caption = Me.DataGridDummy.Columns(0).Text + " - " + Me.DataGridDummy.Columns(1).Text + " stocks remaining"
        y = Me.DataGridDummy.Columns(2).Text
    End If
    
    Call CnnctRcrdSt
    RcrdSt.Open "SELECT TOP 1 ProductName, Stocks FROM Products WHERE ProductID <> '" & x & "' AND ProductID <> '" & y & "' ORDER BY Stocks ASC", Dtbs
    Set Me.DataGridDummy.DataSource = RcrdSt
    If Me.DataGridDummy.ApproxCount = 0 Then
        
    Else
        Me.LabelLS(2).Caption = Me.DataGridDummy.Columns(0).Text + " - " + Me.DataGridDummy.Columns(1).Text + " stocks remaining"
    End If
    
    'set high stocks
    Call CnnctRcrdSt
    RcrdSt.Open "SELECT TOP 1 ProductName, Stocks, ProductID FROM Products ORDER BY Stocks DESC", Dtbs
    Set Me.DataGridDummy.DataSource = RcrdSt
    If Me.DataGridDummy.ApproxCount = 0 Then
    
    Else
        Me.LabelHS(0).Caption = Me.DataGridDummy.Columns(0).Text + " - " + Me.DataGridDummy.Columns(1).Text + " stocks remaining"
        x = Me.DataGridDummy.Columns(2).Text
    End If
    
    Call CnnctRcrdSt
    RcrdSt.Open "SELECT TOP 1 ProductName, Stocks, ProductID FROM Products WHERE ProductID <> '" & x & "' ORDER BY Stocks DESC", Dtbs
    Set Me.DataGridDummy.DataSource = RcrdSt
    If Me.DataGridDummy.ApproxCount = 0 Then
    
    Else
        Me.LabelHS(1).Caption = Me.DataGridDummy.Columns(0).Text + " - " + Me.DataGridDummy.Columns(1).Text + " stocks remaining"
        y = Me.DataGridDummy.Columns(2).Text
    End If
    
    Call CnnctRcrdSt
    RcrdSt.Open "SELECT TOP 1 ProductName, Stocks FROM Products WHERE ProductID <> '" & x & "' AND ProductID <> '" & y & "' ORDER BY Stocks DESC", Dtbs
    Set Me.DataGridDummy.DataSource = RcrdSt
    If Me.DataGridDummy.ApproxCount = 0 Then
    
    Else
        Me.LabelHS(2).Caption = Me.DataGridDummy.Columns(0).Text + " - " + Me.DataGridDummy.Columns(1).Text + " stocks remaining"
    End If
End Sub





'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>initialize<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Private Sub Form_Load()
    Call CnnctDtbs
    
    'defaults at load
    Call SetDataGridDefault
    Call SetFrameStocks
    Call SetFrameBackup
   
    Call CnnctRcrdSt
    Me.DataComboSrchBy.Text = "ProductID"
    Qry = "SELECT Products FROM AppData"
    RcrdSt.Open Qry, Dtbs
    Set Me.DataComboSrchBy.RowSource = RcrdSt
    Me.DataComboSrchBy.ListField = "Products"
    
    Call CnnctRcrdSt
    Me.DataComboSrtBy.Text = "ProductID"
    Qry = "SELECT Products FROM AppData"
    RcrdSt.Open Qry, Dtbs
    Set Me.DataComboSrtBy.RowSource = RcrdSt
    Me.DataComboSrtBy.ListField = "Products"
    
    For n = 0 To 8
        Me.Text1(n).Locked = True
    Next n
    
    Me.Label01.Caption = "Product ID:"
    Me.Label02.Caption = "Name:"
    Me.Label03.Caption = "Model:"
    Me.Label04.Caption = "Manufacturer:"
    Me.Label05.Caption = "Specifiations:"
    Me.Label06.Caption = "Type:"
    Me.Label07.Caption = "Price:"
    Me.Label08.Caption = "Stocks:"
    Me.Label09.Caption = "Warranty:"
    
    Me.CommandInsert.Caption = "Insert Product"
    Me.CommandDelete.Caption = "Delete Product"
    Me.CommandUpdate.Caption = "Update Product"
    
    'backup database
    CurrentDate = Date
    Me.LabelSystemDate.Caption = CurrentDate
    
    If CurrentDate >= Me.LabelBackupDate Then
        Call BackupDatabase
        Dtbs.Execute "UPDATE AppData SET BckupDt = '" & DateAdd("d", 7, CurrentDate) & "', LstBckupDt = '" & Me.LabelSystemDate & "' WHERE Products = 'ProductID'"
        Call SetFrameBackup
    End If
End Sub





'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>clicking events<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'recordset clicked
Private Sub TabStripRecordSets_Click()
    Call SetDataGridDefault
    
    'set command buttons
    If Me.TabStripRecordSets.SelectedItem = "Products" Or Me.TabStripRecordSets.SelectedItem = "Customers" Then
        Me.CommandInsert.Caption = "Insert " + Left(Me.TabStripRecordSets.SelectedItem.Caption, Len(Me.TabStripRecordSets.SelectedItem.Caption) - 1)
        Me.CommandDelete.Caption = "Delete " + Left(Me.TabStripRecordSets.SelectedItem.Caption, Len(Me.TabStripRecordSets.SelectedItem.Caption) - 1)
        Me.CommandUpdate.Caption = "Update " + Left(Me.TabStripRecordSets.SelectedItem.Caption, Len(Me.TabStripRecordSets.SelectedItem.Caption) - 1)
        Me.CommandInsert.Enabled = True
        Me.CommandDelete.Enabled = True
        Me.CommandUpdate.Enabled = True
    Else
        Me.CommandInsert.Caption = "Insert "
        Me.CommandDelete.Caption = "Delete "
        Me.CommandUpdate.Caption = "Update "
        Me.CommandInsert.Enabled = False
        Me.CommandDelete.Enabled = False
        Me.CommandUpdate.Enabled = False
    End If
    
    'set datacombo search
    Call CnnctRcrdSt
    SrchWhr = Me.TabStripRecordSets.SelectedItem
    Qry = "SELECT " + SrchWhr + " FROM AppData"
    RcrdSt.Open Qry, Dtbs
    Set Me.DataComboSrchBy.RowSource = RcrdSt
    Me.DataComboSrchBy.ListField = SrchWhr
    Me.DataComboSrchBy.DataField = Empty
    Set Me.DataComboSrchBy.DataSource = RcrdSt
    Me.DataComboSrchBy.DataField = SrchWhr
    
    'set datacombo sort
    Call CnnctRcrdSt
    SrtWhr = Me.TabStripRecordSets.SelectedItem
    Qry = "SELECT " + SrtWhr + " FROM AppData"
    RcrdSt.Open Qry, Dtbs
    Set Me.DataComboSrtBy.RowSource = RcrdSt
    Me.DataComboSrtBy.ListField = SrtWhr
    Me.DataComboSrtBy.DataField = Empty
    Set Me.DataComboSrtBy.DataSource = RcrdSt
    Me.DataComboSrtBy.DataField = SrtWhr
    
    'set preview pane captions
    For x = 0 To 8
    Me.Text1(x).Text = Empty
    Next x
    
    Me.Label01.Caption = Empty
    Me.Label02.Caption = Empty
    Me.Label03.Caption = Empty
    Me.Label04.Caption = Empty
    Me.Label05.Caption = Empty
    Me.Label06.Caption = Empty
    Me.Label07.Caption = Empty
    Me.Label08.Caption = Empty
    Me.Label09.Caption = Empty
    
    Select Case Me.TabStripRecordSets.SelectedItem
        Case "Products"
            For x = 0 To 8
                Me.Text1(x).Visible = True
            Next x
            Me.Label01.Caption = "Product ID:"
            Me.Label02.Caption = "Name:"
            Me.Label03.Caption = "Model:"
            Me.Label04.Caption = "Manufacturer:"
            Me.Label05.Caption = "Specifiations:"
            Me.Label06.Caption = "Type:"
            Me.Label07.Caption = "Price:"
            Me.Label08.Caption = "Stocks:"
            Me.Label09.Caption = "Warranty:"
        Case "Customers"
            For x = 0 To 4
                Me.Text1(x).Visible = True
            Next x
            For x = 5 To 8
                Me.Text1(x).Visible = False
            Next x
            Me.Label01.Caption = "Customer ID:"
            Me.Label02.Caption = "Name:"
            Me.Label03.Caption = "Address:"
            Me.Label04.Caption = "Contact Number:"
            Me.Label05.Caption = "Discount:"
        Case "Items"
            For x = 0 To 4
                Me.Text1(x).Visible = True
            Next x
            For x = 5 To 8
                Me.Text1(x).Visible = False
            Next x
            Me.Label01.Caption = "Item ID:"
            Me.Label02.Caption = "Product ID:"
            Me.Label03.Caption = "Status:"
            Me.Label04.Caption = "Order ID:"
            Me.Label05.Caption = "Warranty ID:"
        Case "Warranties"
            For x = 0 To 2
                Me.Text1(x).Visible = True
            Next x
            For x = 3 To 8
                Me.Text1(x).Visible = False
            Next x
            Me.Label01.Caption = "Warranty ID:"
            Me.Label02.Caption = "Expiry Date:"
            Me.Label03.Caption = "Status:"
        Case "Transactions"
            For x = 0 To 2
                Me.Text1(x).Visible = True
            Next x
            For x = 3 To 8
                Me.Text1(x).Visible = False
            Next x
            Me.Label01.Caption = "Transaction ID:"
            Me.Label02.Caption = "Date:"
            Me.Label03.Caption = "Customer ID:"
        Case "Orders"
            For x = 0 To 3
                Me.Text1(x).Visible = True
            Next x
            For x = 4 To 8
                Me.Text1(x).Visible = False
            Next x
            Me.Label01.Caption = "Order ID:"
            Me.Label02.Caption = "Product ID:"
            Me.Label03.Caption = "Quantity:"
            Me.Label04.Caption = "Total Price:"
    End Select
End Sub
'datagrid clicked
Private Sub DataGridDefault_Click()
    'set preview pane text
    If Me.DataGridDefault.ApproxCount = 0 Then
    
    Else
        
        Select Case Me.TabStripRecordSets.SelectedItem
            Case "Products"
                x = 8
            Case "Items"
                x = 4
            Case "Customers"
                x = 4
            Case "Transactions"
                x = 2
            Case "Warranties"
                x = 2
            Case "Orders"
                x = 3
        End Select
    
    For y = 0 To 8
    Me.Text1(y).Text = Empty
    Next y
    
    For y = 0 To x
    Me.Text1(y).Text = Me.DataGridDefault.Columns(y).Text
    Next y
    
    End If
End Sub






'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>search and sort commands<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'search
Private Sub TextSearch_Change()
    Call CnnctRcrdSt
    Qry = "SELECT * FROM " + Me.TabStripRecordSets.SelectedItem + " WHERE " + Me.DataComboSrchBy.Text + " LIKE " + "'%" + Me.TextSearch.Text + "%'"
    Me.TextQuery.Text = Qry
    RcrdSt.Open Qry, Dtbs
    Set Me.DataGridDefault.DataSource = RcrdSt
End Sub
'sort
Private Sub CommandSortAsc_Click()
    'sort ascending
    Call CnnctRcrdSt
    Qry = "SELECT * FROM " + Me.TabStripRecordSets.SelectedItem + " WHERE " + Me.DataComboSrchBy.Text + " LIKE " + "'%" + Me.TextSearch.Text + "%'" + " ORDER BY " + Me.DataComboSrtBy.Text + " ASC"
    Me.TextQuery.Text = Qry
    RcrdSt.Open Qry, Dtbs
    Set Me.DataGridDefault.DataSource = RcrdSt
End Sub
Private Sub CommandSortDesc_Click()
    'sort descending
    Call CnnctRcrdSt
    Qry = "SELECT * FROM " + Me.TabStripRecordSets.SelectedItem + " WHERE " + Me.DataComboSrchBy.Text + " LIKE " + "'%" + Me.TextSearch.Text + "%'" + " ORDER BY " + Me.DataComboSrtBy.Text + " DESC"
    Me.TextQuery.Text = Qry
    RcrdSt.Open Qry, Dtbs
    Set Me.DataGridDefault.DataSource = RcrdSt
End Sub





'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>insert, update and  delete<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'cmd insert
Private Sub CommandInsert_Click()
    'set buttons
    Me.CommandUpdate.Enabled = False
    Me.CommandDelete.Enabled = False
    Me.TabStripRecordSets.Enabled = False
    
    'set new id
    If Me.TabStripRecordSets.SelectedItem = "Products" Then
        x = "p"
    ElseIf Me.TabStripRecordSets.SelectedItem = "Customers" Then
        x = "c"
    End If
    
    Call CnnctRcrdSt
    Qry = "SELECT LAST(" + Left(Me.TabStripRecordSets.SelectedItem.Caption, Len(Me.TabStripRecordSets.SelectedItem.Caption) - 1) + "ID) FROM Backup" + Me.TabStripRecordSets.SelectedItem
    RcrdSt.Open Qry, Dtbs
    Set Me.DataGridDummy.DataSource = RcrdSt
    FormInsert.Text1(0).Text = Me.DataGridDummy.Columns(0).Text
    FormInsert.Text1(0).Text = Val(Right(FormInsert.Text1(0).Text, 5)) + 1
    
    If Len(FormInsert.Text1(0).Text) = 1 Then
        FormInsert.Text1(0).Text = "0000" + FormInsert.Text1(0).Text
    ElseIf Len(FormInsert.Text1(0).Text) = 2 Then
        FormInsert.Text1(0).Text = "000" + FormInsert.Text1(0).Text
    ElseIf Len(FormInsert.Text1(0).Text) = 3 Then
        FormInsert.Text1(0).Text = "00" + FormInsert.Text1(0).Text
    ElseIf Len(FormInsert.Text1(0).Text) = 4 Then
        FormInsert.Text1(0).Text = "0" + FormInsert.Text1(0).Text
    End If
    
    FormInsert.Text1(0).Text = x + FormInsert.Text1(0).Text

    'set form insert
    If Me.TabStripRecordSets.SelectedItem = "Products" Then
        FormInsert.Caption = "Insert Product"
        
        For n = 0 To 7
            FormInsert.Text1(n).Visible = True
        Next n
        
        With FormInsert
            .Label01.Caption = "Product ID:"
            .Label02.Caption = "Name:"
            .Label03.Caption = "Model:"
            .Label04.Caption = "Manufacturer:"
            .Label05.Caption = "Specifiations:"
            .Label06.Caption = "Type:"
            .Label07.Caption = "Price:"
            .Label08.Caption = "Warranty:"
        End With
        
        With FormInsert
            .Width = 14025
            .CommandInsert.Left = 9240
            .CommandCancel.Left = 11520
        End With
        
    ElseIf Me.TabStripRecordSets.SelectedItem = "Customers" Then
        FormInsert.Caption = "Insert Customer"
        
        For n = 0 To 3
            FormInsert.Text1(n).Visible = True
        Next n
        
        For n = 4 To 7
            FormInsert.Text1(n).Visible = False
        Next n
        
        With FormInsert
            .Label01.Caption = "Customer ID:"
            .Label02.Caption = "Name:"
            .Label03.Caption = "Address:"
            .Label04.Caption = "Contact Number:"
            .Label05.Caption = ""
            .Label06.Caption = ""
            .Label07.Caption = ""
            .Label08.Caption = ""
        End With
        
        With FormInsert
            .Width = 9480
            .CommandInsert.Left = 4680
            .CommandCancel.Left = 6960
        End With
        
    End If
    
    FormInsert.Show
    
End Sub
'cmd update
Private Sub CommandUpdate_Click()
    If Me.Text1(0).Text <> Empty Then
        'set cmd buttons
        Me.CommandInsert.Enabled = False
        Me.CommandDelete.Enabled = False
        Me.TabStripRecordSets.Enabled = False
    
        'set form update
        If Me.TabStripRecordSets.SelectedItem = "Products" Then
            FormUpdate.Caption = "Update Product"
        
            For n = 0 To 6
                FormUpdate.Text1(n).Visible = True
                FormUpdate.Text1(n).Text = Me.Text1(n).Text
            Next n
            
            FormUpdate.Text1(7).Visible = True
            FormUpdate.Text1(7).Text = Me.Text1(8).Text
        
            With FormUpdate
                .Label01.Caption = "Product ID:"
                .Label02.Caption = "Name:"
                .Label03.Caption = "Model:"
                .Label04.Caption = "Manufacturer:"
                .Label05.Caption = "Specifiations:"
                .Label06.Caption = "Type:"
                .Label07.Caption = "Price:"
                .Label08.Caption = "Warranty:"
            End With
        
            With FormUpdate
                .Width = 14025
                .CommandUpdate.Left = 9240
                .CommandCancel.Left = 11520
            End With
        
        ElseIf Me.TabStripRecordSets.SelectedItem = "Customers" Then
            FormUpdate.Caption = "Update Customer"
    
            For n = 0 To 3
                FormUpdate.Text1(n).Visible = True
                FormUpdate.Text1(n).Text = Me.Text1(n).Text
            Next n
        
            For n = 4 To 7
                FormUpdate.Text1(n).Visible = False
            Next n
        
            With FormUpdate
                .Label01.Caption = "Customer ID:"
                .Label02.Caption = "Name:"
                .Label03.Caption = "Address:"
                .Label04.Caption = "Contact Number:"
                .Label05.Caption = ""
                .Label06.Caption = ""
                .Label07.Caption = ""
                .Label08.Caption = ""
                .Label08.Caption = ""
            End With
        
            With FormUpdate
                .Width = 9480
                .CommandUpdate.Left = 4680
                .CommandCancel.Left = 6960
            End With
        
        End If
    
        FormUpdate.Show
    
    Else
        msg = MsgBox("Select a " + Left(Me.TabStripRecordSets.SelectedItem.Caption, Len(Me.TabStripRecordSets.SelectedItem.Caption) - 1) + " to update.", vbInformation, "Update")
    End If
End Sub
'cmd delete
Private Sub CommandDelete_Click()
    If Me.Text1(0).Text <> Empty Then
        Me.CommandInsert.Enabled = False
        Me.CommandUpdate.Enabled = False
        Me.TabStripRecordSets.Enabled = False
        
        'set form delete
        If Me.TabStripRecordSets.SelectedItem = "Products" Then
        
        For n = 0 To 8
            FormDelete.Text1(n).Visible = True
            FormDelete.Text1(n).Text = Me.Text1(n).Text
        Next n
        
        With FormDelete
            .Label01.Caption = "Product ID:"
            .Label02.Caption = "Name:"
            .Label03.Caption = "Model:"
            .Label04.Caption = "Manufacturer:"
            .Label05.Caption = "Specifiations:"
            .Label06.Caption = "Type:"
            .Label07.Caption = "Price:"
            .Label08.Caption = "Stocks:"
            .Label09.Caption = "Warranty:"
        End With
        
        With FormDelete
                .Width = 14025
        End With
        
    ElseIf Me.TabStripRecordSets.SelectedItem = "Customers" Then
    
        For n = 0 To 4
            FormDelete.Text1(n).Visible = True
            FormDelete.Text1(n).Text = Me.Text1(n).Text
        Next n
        
        For n = 5 To 8
            FormDelete.Text1(n).Visible = False
        Next n
        
        With FormDelete
            .Label01.Caption = "Customer ID:"
            .Label02.Caption = "Name:"
            .Label03.Caption = "Address:"
            .Label04.Caption = "Contact Number:"
            .Label05.Caption = "Discount:"
            .Label06.Caption = ""
            .Label07.Caption = ""
            .Label08.Caption = ""
            .Label09.Caption = ""
        End With
        
        With FormDelete
                .Width = 9480
        End With
        
    End If
    
        FormDelete.Show
        
    Else
        msg = MsgBox("Select a " + Left(Me.TabStripRecordSets.SelectedItem.Caption, Len(Me.TabStripRecordSets.SelectedItem.Caption) - 1) + " to delete.", vbInformation, "Delete")
    End If
End Sub





'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Backup and Restore Database<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'cmd backup database
Private Sub CommandBackupDatabase_Click()
    Call BackupDatabase
End Sub
'cmd restore database
Private Sub CommandRestoreBackup_Click()
    Dtbs.Execute "DELETE FROM Products"
    Dtbs.Execute "DELETE FROM Customers"
    Dtbs.Execute "INSERT INTO Products SELECT * FROM BackupProducts"
    Dtbs.Execute "INSERT INTO Customers SELECT * FROM BackupCustomers"
    Call SetDataGridDefault
    msg = MsgBox("Backup database restored.", vbInformation + vbOKOnly, "Restore Database")
End Sub





'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>refresh stocks andd add stocks<<<<<<<<<<<<<<<<<<<<
'cmd add stocks
Private Sub CommandAddStocks_Click()
    If Me.TabStripRecordSets.SelectedItem = "Products" And Me.Text1(0).Text <> Empty Then
            
        inpt = InputBox("How many " + Me.Text1(1).Text + " dou you want to add?", "Add Stocks")
            
        If inpt = Empty Then
            
        ElseIf Val(inpt) > 0 Then
            
            'set item
            For x = 1 To Val(inpt)
                
                Call CnnctRcrdSt
                Qry = "SELECT LAST(ItemID) FROM Items"
                RcrdSt.Open Qry, Dtbs
                Set FormMain.DataGridDummy.DataSource = RcrdSt
                
                nxtid = FormMain.DataGridDummy.Columns(0).Text
                nxtid = Val(Right(nxtid, 5)) + 1
                    
                If Len(nxtid) = 1 Then
                    nxtid = "0000" + nxtid
                ElseIf Len(nxtid) = 2 Then
                    nxtid = "000" + nxtid
                ElseIf Len(nxtid) = 3 Then
                    nxtid = "00" + nxtid
                ElseIf Len(nxtid) = 4 Then
                    nxtid = "0" + nxtid
                End If
                    
                nxtid = "i" + nxtid
                Qry = "INSERT INTO Items (ItemID, ProductID, Status) VALUES ('" + nxtid + "', '" & FormMain.Text1(0).Text & "', 'Unsold')"
                Dtbs.Execute Qry
                
            Next x
            
            stckstoadd = Val(inpt) + Val(FormMain.Text1(7).Text)
            
            Qry = "UPDATE Products SET Stocks = '" + stckstoadd + "' WHERE ProductID = '" + FormMain.Text1(0).Text + "'"
            Dtbs.Execute Qry
            
            Qry = "UPDATE BackupProducts SET Stocks = '" + stckstoadd + "' WHERE ProductID = '" + FormMain.Text1(0).Text + "'"
            Dtbs.Execute Qry
            
            Call CnnctRcrdSt
            Qry = "SELECT * FROM " + Me.TabStripRecordSets.SelectedItem + " WHERE ProductID = '" & FormMain.Text1(0).Text & "'"
            RcrdSt.Open Qry, Dtbs
            Set FormMain.DataGridDefault.DataSource = RcrdSt
            
            Call DataGridDefault_Click
            Call SetDataGridDefault
        Else
            msg = MsgBox("Enter a number greater than 0.", vbInformation, "Add Stocks")
            Call CommandAddStocks_Click
        End If
    Else
        msg = MsgBox("Select a product first!", vbInformation, "Add Stocks")
    End If
End Sub
'cmd refresh stocks
Private Sub CommandRefreshStocks_Click()
    Call SetFrameStocks
End Sub




'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>New Transaction<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Public Sub CommandNewTransaction_Click()
    FormNewTransaction.Show
End Sub








