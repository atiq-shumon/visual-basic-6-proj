VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Pay scale / Wage scale Setup"
   ClientHeight    =   6345
   ClientLeft      =   750
   ClientTop       =   1665
   ClientWidth     =   9765
   Icon            =   "frmPayscale.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form37"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9765
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   6705
      Picture         =   "frmPayscale.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5670
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   2760
      Picture         =   "frmPayscale.frx":234C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5670
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   1530
      Picture         =   "frmPayscale.frx":3CDE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5625
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   5355
      Picture         =   "frmPayscale.frx":5670
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5670
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   4095
      Picture         =   "frmPayscale.frx":725A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5670
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5325
      Left            =   135
      TabIndex        =   18
      Top             =   135
      Width           =   9510
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3885
         Left            =   180
         TabIndex        =   27
         Top             =   1170
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   6853
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BorderStyle     =   0
         ColumnHeaders   =   0   'False
         ForeColor       =   10485760
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   6
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
         ColumnCount     =   13
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
         BeginProperty Column04 
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
         BeginProperty Column05 
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
         BeginProperty Column06 
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
         BeginProperty Column07 
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
         BeginProperty Column08 
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
         BeginProperty Column09 
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
         BeginProperty Column10 
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
         BeginProperty Column11 
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
         BeginProperty Column12 
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
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   599.811
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtTiffin 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   7470
         TabIndex        =   9
         Top             =   765
         Width           =   510
      End
      Begin VB.TextBox txtConv 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   6795
         TabIndex        =   8
         Top             =   765
         Width           =   375
      End
      Begin VB.TextBox txtMed 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   6075
         TabIndex        =   7
         Top             =   765
         Width           =   510
      End
      Begin VB.TextBox txtMinHR 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   8775
         TabIndex        =   11
         Top             =   765
         Width           =   510
      End
      Begin VB.TextBox txtHR 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   8145
         TabIndex        =   10
         Top             =   765
         Width           =   465
      End
      Begin VB.TextBox txtYear 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1305
         TabIndex        =   1
         Top             =   765
         Width           =   870
      End
      Begin VB.TextBox txtScaleCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   270
         MaxLength       =   5
         TabIndex        =   0
         Top             =   765
         Width           =   780
      End
      Begin VB.TextBox txtStart 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2565
         MaxLength       =   5
         TabIndex        =   2
         Text            =   "0"
         Top             =   765
         Width           =   600
      End
      Begin VB.TextBox txtIncrement 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "0"
         Top             =   765
         Width           =   555
      End
      Begin VB.TextBox txtEnd_Limit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3900
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "0"
         Top             =   765
         Width           =   600
      End
      Begin VB.TextBox txtEBEnd_Limit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   5265
         MaxLength       =   6
         TabIndex        =   6
         Text            =   "0"
         Top             =   765
         Width           =   690
      End
      Begin VB.TextBox txtEBIncrement 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   4695
         MaxLength       =   4
         TabIndex        =   5
         Text            =   "0"
         Top             =   765
         Width           =   510
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Scale/Grade"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   190
         TabIndex        =   28
         Top             =   405
         Width           =   915
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Min (HR)"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   8685
         TabIndex        =   26
         Top             =   405
         Width           =   630
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year Ref."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   1395
         TabIndex        =   25
         Top             =   405
         Width           =   675
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   14
         Left            =   5985
         Top             =   360
         Width           =   690
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   13
         Left            =   8640
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tiffin"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   7560
         TabIndex        =   24
         Top             =   405
         Width           =   345
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   12
         Left            =   8100
         Top             =   360
         Width           =   555
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   10
         Left            =   8640
         Top             =   675
         Width           =   735
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Conv."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   6795
         TabIndex        =   23
         Top             =   405
         Width           =   420
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   9
         Left            =   7425
         Top             =   360
         Width           =   690
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   7
         Left            =   7425
         Top             =   675
         Width           =   690
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Medical"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   6030
         TabIndex        =   22
         Top             =   405
         Width           =   555
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   6
         Left            =   6660
         Top             =   360
         Width           =   780
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   0
         Left            =   5985
         Top             =   675
         Width           =   690
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   5
         Left            =   2430
         Top             =   360
         Width           =   3570
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "HR %"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   8190
         TabIndex        =   21
         Top             =   405
         Width           =   405
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   8
         Left            =   1170
         Top             =   360
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   3
         Left            =   135
         Top             =   360
         Width           =   1050
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   4
         Left            =   135
         Top             =   675
         Width           =   1050
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Scale Definition"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3600
         TabIndex        =   20
         Top             =   405
         Width           =   1110
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   1
         Left            =   1170
         Top             =   675
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   11
         Left            =   8640
         Top             =   675
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         X1              =   3150
         X2              =   3330
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000C0&
         X1              =   3825
         X2              =   4050
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   4020
         Index           =   2
         Left            =   135
         Top             =   1080
         Width           =   9240
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         X1              =   5175
         X2              =   5400
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EB"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4500
         TabIndex        =   19
         Top             =   780
         Width           =   210
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Index           =   0
      Left            =   1440
      Top             =   3600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Index           =   2
      Left            =   2655
      Top             =   3600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Index           =   3
      Left            =   2655
      Top             =   3870
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Index           =   4
      Left            =   225
      Top             =   3870
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Index           =   1
      Left            =   1440
      Top             =   3870
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Index           =   5
      Left            =   225
      Top             =   3600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   855
      TabIndex        =   17
      Top             =   630
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PayScale As New St_Payscale
Private PayScale_Rs As New Recordset
Dim Track_Id As Long
Private Sub cmdClear_Click()
    Clear_Screen
    txtScaleCode.SetFocus
End Sub

Private Sub cmdClose_Click()
    Close_Msg Me
End Sub
Private Sub cmdDelete_Click()
On Error GoTo Errdes
    With PayScale
        .Connstring = strCN.Connection_String
        .Scale_code = txtScaleCode
        .Delete
        .Show_Message
    End With
    
    Flash_Into_Grid
    Clear_Screen
    txtScaleCode.SetFocus
Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub
Private Sub cmdSave_Click()
On Error GoTo Errdes
If Len(Trim(txtScaleCode)) = 0 Then
    MsgBox "Scale Code is not Available", vbInformation, "IT Division, DNMIH"
    txtScaleCode.SetFocus
    Exit Sub
ElseIf Len(Trim(txtYear)) = 0 Then
    MsgBox "Year is not Available", vbInformation, "IT Division, DNMIH"
    txtYear.SetFocus
    Exit Sub
ElseIf Len(Trim(txtStart)) = 0 Then
    MsgBox "Data is not Available", vbInformation, "IT Division, DNMIH"
    txtStart.SetFocus
    Exit Sub
    
ElseIf Len(Trim(txtIncrement)) = 0 Then
    MsgBox "Data not Available", vbInformation, "IT Division, DNMIH"
    txtIncrement.SetFocus
    Exit Sub
    
ElseIf Len(Trim(txtEnd_Limit)) = 0 Then
    MsgBox "Data not Available", vbInformation, "IT Division, DNMIH"
    txtEnd_Limit.SetFocus
    Exit Sub
ElseIf Len(Trim(txtEBIncrement)) = 0 Then
    MsgBox "Data not Available", vbInformation, "IT Division, DNMIH"
    txtEBIncrement.SetFocus
    Exit Sub
ElseIf Len(Trim(txtEBEnd_Limit)) = 0 Then
    MsgBox "Data not Available", vbInformation, "IT Division, DNMIH"
    txtEBEnd_Limit.SetFocus
    Exit Sub
ElseIf Len(Trim(txtHR)) = 0 Then
    MsgBox "Data not Available", vbInformation, "IT Division, DNMIH"
    txtHR.SetFocus
    Exit Sub
ElseIf Len(Trim(txtMinHR)) = 0 Then
    MsgBox "Data not Available", vbInformation, "IT Division, DNMIH"
    txtMinHR.SetFocus
    Exit Sub
ElseIf Len(Trim(txtMed)) = 0 Then
    MsgBox "Data not Available", vbInformation, "IT Division, DNMIH"
    txtMed.SetFocus
    Exit Sub
ElseIf Len(Trim(txtConv)) = 0 Then
    MsgBox "Convence is not Available", vbInformation, "IT Division, DNMIH"
    txtConv.SetFocus
    Exit Sub
ElseIf Len(Trim(txtTiffin)) = 0 Then
    MsgBox "Year is not Available", vbInformation, "IT Division, DNMIH"
    txtTiffin.SetFocus
    Exit Sub
Else

    With PayScale
        .Connstring = strCN.Connection_String
        .Scale_code = txtScaleCode
        .Yr_Ref = txtYear
        .Str_Basic = txtStart
        .Incr = txtIncrement
        .End_basic = txtEnd_Limit
        .Eb_incr = txtEBIncrement
        .Eb_end = txtEBEnd_Limit
        .HR = txtHR
        .MinHR = txtMinHR
        .MED = txtMed
        .CONV = txtConv
        .TFN = txtTiffin
        .Save
        
    End With
    
    Flash_Into_Grid
    
    Track_Id = 0
    Clear_Screen
    
    txtScaleCode.SetFocus
End If
Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
        
    With PayScale
    
        txtScaleCode = DataGrid1.Columns(0)
        txtYear = DataGrid1.Columns(1)
        txtStart = DataGrid1.Columns(2)
        txtIncrement = DataGrid1.Columns(3)
        txtEnd_Limit = DataGrid1.Columns(4)
        txtEBIncrement = DataGrid1.Columns(5)
        txtEBEnd_Limit = DataGrid1.Columns(6)
        txtMed = DataGrid1.Columns(7)
        txtConv = DataGrid1.Columns(8)
        txtTiffin = DataGrid1.Columns(9)
                
       ' txtPercentage = DataGrid1.Columns(10)
        txtMinHR = DataGrid1.Columns(11)
        txtHR = DataGrid1.Columns(10)
        txtStart.SetFocus
     
     End With
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Flash_Into_Grid
End Sub

Public Sub Flash_Into_Grid()
On Error GoTo Errdes
    With PayScale
        .Connstring = strCN.Connection_String
     Set DataGrid1.DataSource = .GetAll
    End With
    
        With DataGrid1
            .Columns(0).Width = 660
            '.Columns(0).DataField = PayScale_Rs.Fields(0)
            .Columns(1).Width = 1260
            '.Columns(1).DataField = PayScale_Rs.Fields(1)
            .Columns(2).Width = 700
            '.Columns(2).DataField = PayScale_Rs.Fields(2)
            .Columns(3).Width = 700
            '.Columns(3).DataField = PayScale_Rs.Fields(3)
            .Columns(4).Width = 700
            '.Columns(4).DataField = PayScale_Rs.Fields(4)
            .Columns(5).Width = 750
            .Columns(6).Width = 700
            .Columns(7).Width = 700
            .Columns(8).Width = 750
            .Columns(9).Width = 675
            .Columns(10).Width = 590
            .Columns(11).Width = 650
             
        End With
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Destroy Me
End Sub


Private Sub txtEnd_KeyPress(KeyAscii As Integer)
   KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtConv_KeyPress(KeyAscii As Integer)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtEBEnd_Limit_KeyPress(KeyAscii As Integer)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtEBIncrement_KeyPress(KeyAscii As Integer)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only

End Sub

Private Sub txtEnd_Limit_KeyPress(KeyAscii As Integer)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtHR_KeyPress(KeyAscii As Integer)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only

End Sub

Private Sub txtIncrement_KeyPress(KeyAscii As Integer)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtMed_KeyPress(KeyAscii As Integer)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only

End Sub

Private Sub txtMinHR_KeyPress(KeyAscii As Integer)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only

End Sub

Private Sub txtScaleCode_KeyPress(KeyAscii As Integer)
On Error GoTo Errdes
    If txtScaleCode <> "" And KeyAscii = 13 Then
    
        With PayScale
            .Connstring = strCN.Connection_String
            .Scale_code = txtScaleCode
            .GetX
            
            txtScaleCode = .Scale_code
            txtYear = .Yr_Ref
            txtStart = .Str_Basic
            txtIncrement = .Incr
            txtEnd_Limit = .End_basic
            txtEBIncrement = .Eb_incr
            txtEBEnd_Limit = .Eb_end
            'txtPercentage = .HR
            txtHR = .HR
            txtMinHR = .MinHR
            txtMed = .MED
            txtConv = .CONV
            txtTiffin = .TFN
            
            txtStart.SetFocus
            
        End With
    
    End If
Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub txtStart_KeyPress(KeyAscii As Integer)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtTiffin_KeyPress(KeyAscii As Integer)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only

End Sub

Private Sub txtWash_KeyPress(KeyAscii As Integer)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only

End Sub
