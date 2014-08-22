VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FormNewTransaction 
   Caption         =   "New Transaction"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15495
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   15495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameSelectProduct 
      Caption         =   "Select Product"
      Height          =   3855
      Index           =   1
      Left            =   120
      TabIndex        =   79
      Top             =   4080
      Visible         =   0   'False
      Width           =   15255
      Begin VB.Frame FrameSearch 
         Caption         =   "Search"
         Height          =   1095
         Index           =   1
         Left            =   120
         TabIndex        =   85
         Top             =   240
         Width           =   4695
         Begin VB.TextBox TextSearch 
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   86
            Top             =   720
            Width           =   4455
         End
         Begin MSDataListLib.DataCombo DataComboSrchBy 
            Height          =   330
            Index           =   1
            Left            =   1320
            TabIndex        =   87
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   582
            _Version        =   393216
            Locked          =   -1  'True
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
            Index           =   1
            Left            =   120
            TabIndex        =   88
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame FrameSort 
         Caption         =   "Sort"
         Height          =   1095
         Index           =   1
         Left            =   120
         TabIndex        =   80
         Top             =   1440
         Width           =   4695
         Begin VB.CommandButton CommandSortDesc 
            Caption         =   "Sort Descending"
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   82
            Top             =   600
            Width           =   2175
         End
         Begin VB.CommandButton CommandSortAsc 
            Caption         =   "Sort Ascending"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   81
            Top             =   600
            Width           =   2175
         End
         Begin MSDataListLib.DataCombo DataComboSrtBy 
            Height          =   330
            Index           =   1
            Left            =   1320
            TabIndex        =   83
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   582
            _Version        =   393216
            Locked          =   -1  'True
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
            Index           =   1
            Left            =   120
            TabIndex        =   84
            Top             =   360
            Width           =   1095
         End
      End
      Begin MSDataGridLib.DataGrid DataGridDefault 
         Height          =   3495
         Index           =   1
         Left            =   4920
         TabIndex        =   89
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   6165
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
   Begin VB.Frame FrameProcessTransaction 
      Caption         =   "Process Transaction"
      Height          =   3855
      Left            =   10080
      TabIndex        =   71
      Top             =   120
      Width           =   5295
      Begin VB.TextBox TextTotalPrice 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton CommandProcessTransaction 
         Caption         =   "Process Transaction"
         Height          =   495
         Left            =   2880
         TabIndex        =   78
         Top             =   3240
         Width           =   2295
      End
      Begin VB.TextBox TextDiscountedPrice 
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox TextChange 
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox TextCash 
         Height          =   285
         Left            =   120
         TabIndex        =   72
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label LabelTotalPrice 
         Caption         =   "Total Price:"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   90
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label LabelDiscountedPrice 
         Caption         =   "Discounted Price:"
         Height          =   255
         Left            =   2760
         TabIndex        =   76
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label LabelChange 
         Caption         =   "Change:"
         Height          =   255
         Left            =   2760
         TabIndex        =   75
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label LabelCash 
         Caption         =   "Cash:"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   1080
         Width           =   2775
      End
   End
   Begin VB.Frame FrameOrders 
      Caption         =   "Orders"
      Height          =   3855
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   15255
      Begin VB.TextBox TextProdcuct 
         Height          =   285
         Index           =   4
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   1680
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox TextQuantity 
         Height          =   285
         Index           =   4
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TextPrice 
         Height          =   285
         Index           =   4
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TextPrice 
         Height          =   285
         Index           =   8
         Left            =   13320
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TextPrice 
         Height          =   285
         Index           =   7
         Left            =   13320
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TextPrice 
         Height          =   285
         Index           =   6
         Left            =   13320
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TextPrice 
         Height          =   285
         Index           =   5
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TextPrice 
         Height          =   285
         Index           =   3
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TextPrice 
         Height          =   285
         Index           =   2
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TextPrice 
         Height          =   285
         Index           =   1
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TextProdcuct 
         Height          =   285
         Index           =   0
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox TextQuantity 
         Height          =   285
         Index           =   0
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox TextQuantity 
         Height          =   285
         Index           =   1
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TextProdcuct 
         Height          =   285
         Index           =   1
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1680
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox TextProdcuct 
         Height          =   285
         Index           =   2
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2880
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox TextQuantity 
         Height          =   285
         Index           =   2
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TextProdcuct 
         Height          =   285
         Index           =   3
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox TextQuantity 
         Height          =   285
         Index           =   3
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TextQuantity 
         Height          =   285
         Index           =   5
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TextProdcuct 
         Height          =   285
         Index           =   5
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2880
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox TextProdcuct 
         Height          =   285
         Index           =   6
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox TextQuantity 
         Height          =   285
         Index           =   6
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TextQuantity 
         Height          =   285
         Index           =   7
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TextProdcuct 
         Height          =   285
         Index           =   7
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1680
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox TextPrice 
         Height          =   285
         Index           =   0
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox TextProdcuct 
         Height          =   285
         Index           =   8
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2880
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox TextQuantity 
         Height          =   285
         Index           =   8
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Index           =   4
         Left            =   5160
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LabelProduct 
         Caption         =   "Poduct Name:"
         Height          =   255
         Index           =   4
         Left            =   6360
         TabIndex        =   97
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LabelQuantity 
         Caption         =   "Quantity:"
         Height          =   255
         Index           =   4
         Left            =   6360
         TabIndex        =   96
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LabelTotalPrice 
         Caption         =   "Total Price:"
         Height          =   255
         Index           =   4
         Left            =   8280
         TabIndex        =   95
         Top             =   2040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LabelTotalPrice 
         Caption         =   "Total Price:"
         Height          =   255
         Index           =   8
         Left            =   13320
         TabIndex        =   70
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LabelTotalPrice 
         Caption         =   "Total Price:"
         Height          =   255
         Index           =   7
         Left            =   13320
         TabIndex        =   68
         Top             =   2040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LabelTotalPrice 
         Caption         =   "Total Price:"
         Height          =   255
         Index           =   6
         Left            =   13320
         TabIndex        =   66
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LabelTotalPrice 
         Caption         =   "Total Price:"
         Height          =   255
         Index           =   5
         Left            =   8280
         TabIndex        =   64
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LabelTotalPrice 
         Caption         =   "Total Price:"
         Height          =   255
         Index           =   3
         Left            =   8280
         TabIndex        =   62
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LabelTotalPrice 
         Caption         =   "Total Price:"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   60
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LabelTotalPrice 
         Caption         =   "Total Price:"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   58
         Top             =   2040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Index           =   0
         Left            =   120
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LabelProduct 
         Caption         =   "Poduct Name:"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   56
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label LabelQuantity 
         Caption         =   "Quantity:"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   55
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label LabelQuantity 
         Caption         =   "Quantity:"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   54
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LabelProduct 
         Caption         =   "Poduct Name:"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   53
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Index           =   1
         Left            =   120
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Index           =   2
         Left            =   120
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LabelProduct 
         Caption         =   "Poduct Name:"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   52
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LabelQuantity 
         Caption         =   "Quantity:"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   51
         Top             =   3240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Index           =   3
         Left            =   5160
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LabelProduct 
         Caption         =   "Poduct Name:"
         Height          =   255
         Index           =   3
         Left            =   6360
         TabIndex        =   50
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LabelQuantity 
         Caption         =   "Quantity:"
         Height          =   255
         Index           =   3
         Left            =   6360
         TabIndex        =   49
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LabelQuantity 
         Caption         =   "Quantity:"
         Height          =   255
         Index           =   5
         Left            =   6360
         TabIndex        =   48
         Top             =   3240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LabelProduct 
         Caption         =   "Poduct Name:"
         Height          =   255
         Index           =   5
         Left            =   6360
         TabIndex        =   47
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Index           =   5
         Left            =   5160
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Index           =   6
         Left            =   10200
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LabelProduct 
         Caption         =   "Poduct Name:"
         Height          =   255
         Index           =   6
         Left            =   11400
         TabIndex        =   46
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LabelQuantity 
         Caption         =   "Quantity:"
         Height          =   255
         Index           =   6
         Left            =   11400
         TabIndex        =   45
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LabelQuantity 
         Caption         =   "Quantity:"
         Height          =   255
         Index           =   7
         Left            =   11400
         TabIndex        =   44
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LabelProduct 
         Caption         =   "Poduct Name:"
         Height          =   255
         Index           =   7
         Left            =   11400
         TabIndex        =   43
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Index           =   7
         Left            =   10200
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LabelTotalPrice 
         Caption         =   "Total Price:"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   42
         Top             =   840
         Width           =   975
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Index           =   8
         Left            =   10200
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LabelProduct 
         Caption         =   "Poduct Name:"
         Height          =   255
         Index           =   8
         Left            =   11400
         TabIndex        =   41
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LabelQuantity 
         Caption         =   "Quantity:"
         Height          =   255
         Index           =   8
         Left            =   11400
         TabIndex        =   40
         Top             =   3240
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame FrameSelectCustomer 
      Caption         =   "Select Customer"
      Height          =   3855
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   9855
      Begin VB.Frame FrameSort 
         Caption         =   "Sort"
         Height          =   1095
         Index           =   0
         Left            =   5040
         TabIndex        =   16
         Top             =   240
         Width           =   4695
         Begin VB.CommandButton CommandSortAsc 
            Caption         =   "Sort Ascending"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   2175
         End
         Begin VB.CommandButton CommandSortDesc 
            Caption         =   "Sort Descending"
            Height          =   375
            Index           =   0
            Left            =   2400
            TabIndex        =   17
            Top             =   600
            Width           =   2175
         End
         Begin MSDataListLib.DataCombo DataComboSrtBy 
            Height          =   330
            Index           =   0
            Left            =   1320
            TabIndex        =   19
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   582
            _Version        =   393216
            Locked          =   -1  'True
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
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame FrameSearch 
         Caption         =   "Search"
         Height          =   1095
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4695
         Begin VB.TextBox TextSearch 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   4455
         End
         Begin MSDataListLib.DataCombo DataComboSrchBy 
            Height          =   330
            Index           =   0
            Left            =   1320
            TabIndex        =   14
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   582
            _Version        =   393216
            Locked          =   -1  'True
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
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1095
         End
      End
      Begin MSDataGridLib.DataGrid DataGridDefault 
         Height          =   2295
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4048
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
   Begin VB.Frame FrameCustomer 
      Caption         =   "Customer"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1920
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2640
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   3360
         Width           =   4935
      End
      Begin VB.Image ImageCustomer 
         BorderStyle     =   1  'Fixed Single
         Height          =   3495
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label01 
         Caption         =   "Customer ID:"
         Height          =   255
         Left            =   4800
         TabIndex        =   10
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label02 
         Caption         =   "Name:"
         Height          =   255
         Left            =   4800
         TabIndex        =   9
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label03 
         Caption         =   "Address:"
         Height          =   255
         Left            =   4800
         TabIndex        =   8
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Label Label05 
         Caption         =   "Discount:"
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   3120
         Width           =   3375
      End
      Begin VB.Label Label04 
         Caption         =   "Contact:"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   2400
         Width           =   2775
      End
   End
End
Attribute VB_Name = "FormNewTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim orderindex As Integer
Dim totalorders As Integer
Dim orderid As String
Dim transactionid As String
Dim productid As String
Dim warrantyid As String
Dim itemid As String
Dim stocks As Integer
Dim discount As Double

Private Sub Form_Load()
    Call CnnctRcrdSt
    Qry = "Select * From Customers"
    RcrdSt.Open Qry, Dtbs
    Set Me.DataGridDefault(0).DataSource = RcrdSt
    
    Call CnnctRcrdSt
    Me.DataComboSrchBy(0).Text = "CustomerID"
    Qry = "SELECT Customers FROM AppData"
    RcrdSt.Open Qry, Dtbs
    Set Me.DataComboSrchBy(0).RowSource = RcrdSt
    Me.DataComboSrchBy(0).ListField = "Customers"
    
    Call CnnctRcrdSt
    Me.DataComboSrtBy(0).Text = "CustomerID"
    Qry = "SELECT Customers FROM AppData"
    RcrdSt.Open Qry, Dtbs
    Set Me.DataComboSrtBy(0).RowSource = RcrdSt
    Me.DataComboSrtBy(0).ListField = "Customers"
    
    Call CnnctRcrdSt
    Qry = "Select * From Products"
    RcrdSt.Open Qry, Dtbs
    Set Me.DataGridDefault(1).DataSource = RcrdSt
    
    Call CnnctRcrdSt
    Me.DataComboSrchBy(1).Text = "ProductID"
    Qry = "SELECT Products FROM AppData"
    RcrdSt.Open Qry, Dtbs
    Set Me.DataComboSrchBy(1).RowSource = RcrdSt
    Me.DataComboSrchBy(1).ListField = "Products"
    
    Call CnnctRcrdSt
    Me.DataComboSrtBy(1).Text = "ProductID"
    Qry = "SELECT Products FROM AppData"
    RcrdSt.Open Qry, Dtbs
    Set Me.DataComboSrtBy(1).RowSource = RcrdSt
    Me.DataComboSrtBy(1).ListField = "Products"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormMain.Enabled = True
End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Process Transaction<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Private Sub CommandProcessTransaction_Click()
If Me.Text1(0).Text = Empty Then
    msg = MsgBox("Select a customer.", vbInformation, "New Transaction")
ElseIf Me.TextProdcuct(0).Text = Empty Then
    msg = MsgBox("Select an order.", vbInformation, "New Transaction")
Else

    'reset ordertotal
    totalorders = 0
    
    'set new transaction id
    Call CnnctRcrdSt
    Qry = "SELECT LAST(TransactionID) FROM Transactions"
    RcrdSt.Open Qry, Dtbs
    Set FormMain.DataGridDummy.DataSource = RcrdSt
                
    transactionid = Val(Right(FormMain.DataGridDummy.Columns(0).Text, 5)) + 1
                    
    If Len(transactionid) = 1 Then
        transactionid = "0000" + transactionid
    ElseIf Len(transactionid) = 2 Then
        transactionid = "000" + transactionid
    ElseIf Len(transactionid) = 3 Then
        transactionid = "00" + transactionid
    ElseIf Len(transactionid) = 4 Then
        transactionid = "0" + transactionid
    End If
        
    transactionid = "t" + transactionid
    
    'insert transaction
    
    Qry = "INSERT INTO Transactions (TransactionID, TransactionDate, CustomerID, TotalPrice) VALUES ('" & transactionid & "', '" & Now & "', '" & Me.Text1(0).Text & "', '" & Me.TextTotalPrice.Text & "')"
    Dtbs.Execute Qry
    
    'check total num of orders
    For n = 0 To 8
    
        If Me.TextProdcuct(n).Text <> Empty Then
            totalorders = totalorders + 1
        End If
        
    Next n
    
    'set new order id , update transaction, insert orders then insert warranty, update items
    For n = 1 To totalorders
                
        'set new order id
        Call CnnctRcrdSt
        Qry = "SELECT LAST(OrderID) FROM Orders"
        RcrdSt.Open Qry, Dtbs
        Set FormMain.DataGridDummy.DataSource = RcrdSt
                
        orderid = Val(Right(FormMain.DataGridDummy.Columns(0).Text, 5)) + 1
                    
        If Len(orderid) = 1 Then
            orderid = "0000" + orderid
        ElseIf Len(orderid) = 2 Then
            orderid = "000" + orderid
        ElseIf Len(orderid) = 3 Then
            orderid = "00" + orderid
        ElseIf Len(orderid) = 4 Then
            orderid = "0" + orderid
        End If
        
        orderid = "o" + orderid
        
        'update transaction
        Qry = "UPDATE Transactions SET OrderID" & n & " = '" & orderid & "' WHERE TransactionID = '" & transactionid & "'"
        Dtbs.Execute Qry
        
        'get product id
        Call CnnctRcrdSt
        Qry = "SELECT ProductID FROM Products WHERE ProductName = '" & Me.TextProdcuct(n - 1).Text & "'"
        RcrdSt.Open Qry, Dtbs
        Set FormMain.DataGridDummy.DataSource = RcrdSt
        productid = FormMain.DataGridDummy.Columns(0).Text
        
        'insert orders
        Qry = "INSERT INTO Orders (OrderID, ProductID, Quantity, TotalPrice) VALUES ('" & orderid & "', '" & productid & "', '" & Me.TextQuantity(n - 1).Text & "', '" & Me.TextPrice(n - 1) & "')"
        Dtbs.Execute Qry
        
        'insert warranty , update items
        For m = 1 To Me.TextQuantity(n - 1).Text
        
            'set new warranty id
            Call CnnctRcrdSt
            Qry = "SELECT LAST(WarrantyID) FROM Warranties"
            RcrdSt.Open Qry, Dtbs
            Set FormMain.DataGridDummy.DataSource = RcrdSt
                
            warrantyid = Val(Right(FormMain.DataGridDummy.Columns(0).Text, 5)) + 1
                    
            If Len(warrantyid) = 1 Then
                warrantyid = "0000" + warrantyid
            ElseIf Len(warrantyid) = 2 Then
                warrantyid = "000" + warrantyid
            ElseIf Len(warrantyid) = 3 Then
                warrantyid = "00" + warrantyid
            ElseIf Len(warrantyid) = 4 Then
                warrantyid = "0" + warrantyid
            End If
        
            warrantyid = "w" + warrantyid
        
            'insert warranty
            Qry = "INSERT INTO Warranties (WarrantyID, Status) VALUES ('" & warrantyid & "', 'Active')"
            Dtbs.Execute Qry
        
            'set item id
            Call CnnctRcrdSt
            Qry = "SELECT FIRST(ItemID) FROM Items WHERE ProductID = '" & productid & "' AND Status = 'Unsold'"
            RcrdSt.Open Qry, Dtbs
            Set FormMain.DataGridDummy.DataSource = RcrdSt
            itemid = FormMain.DataGridDummy.Columns(0).Text
        
            'update item
            Qry = "UPDATE Items SET OrderID = '" & orderid & "' WHERE ItemID = '" & itemid & "'"
            'msg = MsgBox(itemid)
            'Qry = "UPDATE Items SET OrderID = '" & orderid & "', WarrantyID = '" & warrantyid & "', Status = 'Sold' WHERE ItemID = '" & itemid & "'"
            Dtbs.Execute Qry
        
        Next m
        
        'set stocks
        Call CnnctRcrdSt
        Qry = "SELECT Stocks FROM Products WHERE ProductID = '" & productid & "'"
        RcrdSt.Open Qry, Dtbs
        Set FormMain.DataGridDummy.DataSource = RcrdSt
        stocks = Val(FormMain.DataGridDummy.Columns(0).Text) - Val(Me.TextQuantity(n - 1).Text)
        
        'update products
        Qry = "UPDATE Products SET Stocks = '" & stocks & "' WHERE ProductID = '" & productid & "'"
        Dtbs.Execute Qry
        
        'set discount
        discount = Val(Me.Text1(4).Text) + 0.01
        
        'update customer
        Qry = "UPDATE Customers SET Discount = '" & discount & "' WHERE CustomerID = '" & Me.Text1(0).Text & "'"
        Dtbs.Execute Qry
    Next n
    
    Unload Me
End If
End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Search and Sort<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'search
Private Sub TextSearch_Change(Index As Integer)
If Index = 0 Then
   x = "Customers"
Else
   x = "Products"
End If

    Call CnnctRcrdSt
    Qry = "SELECT * FROM " & x & " WHERE " + Me.DataComboSrchBy(Index).Text + " LIKE " + "'%" + Me.TextSearch(Index).Text + "%'"
    FormMain.TextQuery.Text = Qry
    RcrdSt.Open Qry, Dtbs
    Set Me.DataGridDefault(Index).DataSource = RcrdSt
End Sub
'sort ascending
Private Sub CommandSortAsc_Click(Index As Integer)
If Index = 0 Then
   x = "Customers"
Else
   x = "Products"
End If
    
    Call CnnctRcrdSt
    Qry = "SELECT * FROM " & x & " WHERE " + Me.DataComboSrchBy(Index).Text + " LIKE " + "'%" + Me.TextSearch(Index).Text + "%'" + " ORDER BY " + Me.DataComboSrtBy(Index).Text + " ASC"
    FormMain.TextQuery.Text = Qry
    RcrdSt.Open Qry, Dtbs
    Set Me.DataGridDefault(Index).DataSource = RcrdSt
End Sub
'sort descending
Private Sub CommandSortDesc_Click(Index As Integer)
If Index = 0 Then
   x = "Customers"
Else
   x = "Products"
End If

    Call CnnctRcrdSt
    Qry = "SELECT * FROM " & x & " WHERE " + Me.DataComboSrchBy(Index).Text + " LIKE " + "'%" + Me.TextSearch(Index).Text + "%'" + " ORDER BY " + Me.DataComboSrtBy(Index).Text + " DESC"
    FormMain.TextQuery.Text = Qry
    RcrdSt.Open Qry, Dtbs
    Set Me.DataGridDefault(Index).DataSource = RcrdSt
End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Clicking Events<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'datagrid clicked
Private Sub DataGridDefault_Click(Index As Integer)
If Index = 0 Then

    If Me.DataGridDefault(0).ApproxCount <> 0 Then
    
            For n = 0 To 4
                Me.Text1(n).Text = Me.DataGridDefault(0).Columns(n).Text
            Next n
            
    End If
    
    Me.FrameSelectCustomer(0).Visible = False
    
Else
    
    Me.TextProdcuct(orderindex).Text = Me.DataGridDefault(1).Columns(1).Text
    Me.FrameSelectProduct(1).Visible = False
    Me.TextQuantity(orderindex).Locked = False
    Me.TextQuantity(orderindex).Text = 1
    Me.TextQuantity(orderindex).SetFocus
    
End If
End Sub
'order image clicked
Private Sub Image1_Click(Index As Integer)
    Me.FrameSelectProduct(1).Visible = True
    orderindex = Index
    
'    If Index <> 8 Then
        Me.Image1(Index + 1).Visible = True
        Me.TextProdcuct(Index + 1).Visible = True
        Me.TextQuantity(Index + 1).Visible = True
        Me.TextPrice(Index + 1).Visible = True
'    End If
    
    Me.TextProdcuct(Index).Text = Empty
    Me.TextQuantity(Index).Text = Empty
    Me.TextPrice(Index).Text = Empty
End Sub
'customer image clicked
Private Sub ImageCustomer_Click()
    Me.FrameSelectCustomer(0).Visible = True
End Sub
'search clicked
Private Sub DataComboSrchBy_Click(Index As Integer, Area As Integer)
    If Area = 0 Then
        Me.DataComboSrchBy(Index).Locked = False
    Else
        Me.DataComboSrchBy(Index).Locked = True
    End If
End Sub
'sort clicked
Private Sub DataComboSrtBy_Click(Index As Integer, Area As Integer)
    If Area = 0 Then
        Me.DataComboSrtBy(Index).Locked = False
    Else
        Me.DataComboSrtBy(Index).Locked = True
    End If
End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Text Change Events<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'cash changed
Private Sub TextCash_Change()
    Me.TextChange.Text = Val(Me.TextCash.Text) - Val(Me.TextDiscountedPrice.Text)
End Sub
'order quantity changed
Private Sub TextQuantity_Change(Index As Integer)
    Call CnnctRcrdSt
    Qry = "SELECT * FROM Products WHERE ProductName = '" & Me.TextProdcuct(Index).Text & "'"
    RcrdSt.Open Qry, Dtbs
    Set FormMain.DataGridDummy.DataSource = RcrdSt
    
    'check stock and set quantity
'If Me.TextProdcuct(Index) <> Empty Then
'    If FormMain.DataGridDummy.Columns(7).Value < Val(Me.TextQuantity(Index).Text) Then

'        msg = MsgBox(Me.TextProdcuct(Index).Text & " has " & FormMain.DataGridDummy.Columns(7).Text & " stock(s) remaining.", vbInformation, "Low Stocks")
'        Me.TextQuantity(Index).Text = FormMain.DataGridDummy.Columns(7).Text

'    Else

        'set price, totalprice, discountedprice
'        If Me.TextQuantity(Index).Text <> Empty Then
            Me.TextPrice(Index).Text = Val(Me.TextQuantity(Index).Text) * FormMain.DataGridDummy.Columns(6).Value
'        End If

        Me.TextTotalPrice.Text = Val(Me.TextPrice(0).Text) + Val(Me.TextPrice(1).Text) + Val(Me.TextPrice(2).Text) + Val(Me.TextPrice(3).Text) + Val(Me.TextPrice(4).Text) + Val(Me.TextPrice(5).Text) + Val(Me.TextPrice(6).Text) + Val(Me.TextPrice(7).Text) + Val(Me.TextPrice(8).Text)
        Me.TextDiscountedPrice.Text = Val(Me.TextTotalPrice.Text) - (Val(Me.TextTotalPrice.Text) * (Val(Me.Text1(4).Text) * 0.01))

'    End If
'End If
End Sub
