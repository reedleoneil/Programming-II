VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormNewTransaction 
   Caption         =   "New Transaction"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   17520
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   17520
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   13200
      TabIndex        =   1
      Top             =   7440
      Width           =   3375
   End
   Begin MSDataGridLib.DataGrid DataGridTransaction 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   7646
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
      Caption         =   "Transaction"
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
   Begin VB.Label Label04 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label05 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Label Label03 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label02 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label01 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "FormNewTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
