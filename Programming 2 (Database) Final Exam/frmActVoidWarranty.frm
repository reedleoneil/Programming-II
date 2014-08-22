VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormActivateVoidWarranty 
   Caption         =   "Warranty"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameSearch 
      Caption         =   "Search"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      Begin VB.TextBox TextSearch 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
   End
   Begin MSDataGridLib.DataGrid DataGridDefault 
      Height          =   2490
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   4392
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
End
Attribute VB_Name = "FormActivateVoidWarranty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
