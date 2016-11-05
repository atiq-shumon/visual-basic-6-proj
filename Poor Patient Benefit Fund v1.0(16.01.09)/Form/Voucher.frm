VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voucher Entry"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10575
   FillStyle       =   0  'Solid
   Icon            =   "Voucher.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "<<"
      Height          =   315
      Index           =   1
      Left            =   6900
      TabIndex        =   56
      ToolTipText     =   "Backward"
      Top             =   690
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   ">>"
      Height          =   315
      Index           =   0
      Left            =   7290
      TabIndex        =   55
      ToolTipText     =   "Forward"
      Top             =   690
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Height          =   1410
      Left            =   10020
      TabIndex        =   44
      Top             =   5250
      Visible         =   0   'False
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   2487
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14737632
      ForeColor       =   8388608
      BackColorFixed  =   14737632
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483635
      BackColorBkg    =   16777215
      FocusRect       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox nbrDebit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7620
      MaxLength       =   10
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   2880
      Width           =   1200
   End
   Begin VB.TextBox nbrCredit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8820
      MaxLength       =   12
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   2880
      Width           =   1320
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   1410
      Left            =   9810
      TabIndex        =   45
      Top             =   5100
      Visible         =   0   'False
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   2487
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14737632
      ForeColor       =   8388608
      BackColorFixed  =   14737632
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483635
      BackColorBkg    =   16777215
      FocusRect       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtparticular_name 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2010
      Locked          =   -1  'True
      MaxLength       =   45
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2220
      Width           =   5595
   End
   Begin VB.TextBox cboparticular 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   2220
      Width           =   1755
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   765
      Left            =   -90
      TabIndex        =   42
      Top             =   -120
      Width           =   10755
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Chart of Accounts"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9180
         TabIndex        =   54
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl_Vou_Cap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Entry Screen"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   690
         TabIndex        =   43
         Top             =   180
         Width           =   3300
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Caption         =   "Cheque Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   555
      Left            =   3630
      TabIndex        =   38
      Top             =   4980
      Width           =   2115
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "UnCash"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   40
         Top             =   210
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   39
         Top             =   210
         Width           =   765
      End
   End
   Begin VB.TextBox txtDiff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   8760
      MaxLength       =   10
      TabIndex        =   35
      Top             =   5280
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox txtTrack_ID 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   9360
      MaxLength       =   10
      TabIndex        =   34
      Text            =   "0"
      Top             =   2130
      Visible         =   0   'False
      Width           =   1380
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4620
      Top             =   60
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
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
      Caption         =   "Adodc1"
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
   Begin VB.TextBox txtCustOrdNo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6120
      MaxLength       =   25
      TabIndex        =   5
      Top             =   1620
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.TextBox nbrDollar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   -255
      MaxLength       =   12
      TabIndex        =   16
      Text            =   "0.00"
      Top             =   225
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox nbrRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   870
      MaxLength       =   12
      TabIndex        =   18
      Text            =   "0.00"
      Top             =   225
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox txtUnitCode 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7965
      MaxLength       =   10
      TabIndex        =   30
      Top             =   2115
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   180
      Picture         =   "Voucher.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Posting to Ledger"
      Top             =   5070
      Width           =   465
   End
   Begin VB.TextBox nbrDiff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5865
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5220
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.ComboBox cboUnit 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8370
      Style           =   1  'Simple Combo
      TabIndex        =   4
      Text            =   "000"
      Top             =   1950
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox cboVou_Type 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1905
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   705
      Width           =   795
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1125
      Picture         =   "Voucher.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Edit"
      Top             =   5070
      Width           =   465
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2520
      Picture         =   "Voucher.frx":0DB6
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Preview"
      Top             =   5070
      Width           =   465
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2055
      Picture         =   "Voucher.frx":1420
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Print"
      Top             =   5070
      Width           =   465
   End
   Begin VB.TextBox nbrtot_debit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7620
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   4905
      Width           =   1230
   End
   Begin VB.TextBox txtvou_Chq 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1905
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1620
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox txtVOU_NARR 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   1905
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1035
      Width           =   7755
   End
   Begin VB.TextBox txtVOU_NO 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   1
      Top             =   705
      Width           =   1845
   End
   Begin MSComCtl2.DTPicker dtVOU_DATE 
      Height          =   285
      Index           =   0
      Left            =   8265
      TabIndex        =   2
      Top             =   705
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   503
      _Version        =   393216
      Format          =   57081857
      CurrentDate     =   36955
   End
   Begin VB.CommandButton cmdDELETE 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1590
      Picture         =   "Voucher.frx":1A8A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Delete"
      Top             =   5070
      Width           =   465
   End
   Begin VB.CommandButton cmdADD 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   660
      Picture         =   "Voucher.frx":25C4
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "New"
      Top             =   5070
      Width           =   465
   End
   Begin VB.CommandButton cmdEXIT 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2985
      Picture         =   "Voucher.frx":2C2E
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Exit"
      Top             =   5070
      Width           =   465
   End
   Begin VB.TextBox nbrtot_credit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8850
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   4905
      Width           =   1320
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5910
      Top             =   60
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
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
      Caption         =   "Adodc1"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   7200
      Top             =   60
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
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
      Caption         =   "Adodc1"
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
   Begin MSComCtl2.DTPicker dtVOU_DATE 
      Height          =   285
      Index           =   1
      Left            =   3780
      TabIndex        =   7
      Top             =   1620
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   503
      _Version        =   393216
      Format          =   57081857
      CurrentDate     =   36955
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1725
      Left            =   225
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3210
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   3043
      _Version        =   393216
      Rows            =   0
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14737632
      ForeColor       =   64
      ForeColorFixed  =   12640511
      BackColorSel    =   -2147483624
      ForeColorSel    =   16711680
      BackColorBkg    =   14737632
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.TextBox cboUserAcc 
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   1755
   End
   Begin VB.TextBox txtacc_name 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2010
      Locked          =   -1  'True
      MaxLength       =   45
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2880
      Width           =   5625
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   2055
      TabIndex        =   53
      Top             =   2610
      Width           =   360
   End
   Begin VB.Label lblDR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   52
      Top             =   2610
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   8490
      TabIndex        =   51
      Top             =   2610
      Width           =   420
   End
   Begin VB.Label lblCR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   9705
      TabIndex        =   50
      Top             =   2610
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H008080FF&
      Height          =   615
      Index           =   1
      Left            =   -30
      Top             =   2580
      Width           =   10635
   End
   Begin VB.Label lblDR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Index           =   2
      Left            =   2100
      TabIndex        =   48
      Top             =   1980
      Width           =   360
   End
   Begin VB.Label lblDR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Particular Code"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Index           =   1
      Left            =   210
      TabIndex        =   47
      Top             =   1980
      Width           =   1185
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H008080FF&
      Height          =   615
      Index           =   0
      Left            =   -60
      Top             =   1950
      Width           =   10845
   End
   Begin VB.Shape Shape1 
      Height          =   525
      Left            =   120
      Top             =   5010
      Width           =   3375
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Index           =   1
      Left            =   3390
      TabIndex        =   37
      Top             =   1620
      Width           =   360
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Difference"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   7920
      TabIndex        =   36
      Top             =   5160
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   1095
      TabIndex        =   33
      Top             =   -45
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dollar"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   375
      TabIndex        =   32
      Top             =   -45
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order #"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   5400
      TabIndex        =   31
      Top             =   1635
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Centre"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   8340
      TabIndex        =   29
      Top             =   2040
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H80000001&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   2
      Left            =   -75
      Top             =   4935
      Width           =   10740
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Type"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   765
      TabIndex        =   27
      Top             =   750
      Width           =   1050
   End
   Begin VB.Label lbl_chq 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque#"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Index           =   0
      Left            =   1080
      TabIndex        =   26
      Top             =   1650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Narration "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   1020
      TabIndex        =   25
      Top             =   1050
      Width           =   810
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Index           =   0
      Left            =   7785
      TabIndex        =   24
      Top             =   705
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher #"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   4140
      TabIndex        =   23
      Top             =   705
      Width           =   810
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1485
      Index           =   0
      Left            =   -30
      Top             =   480
      Width           =   10650
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteOne 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strMode As String
Dim intTrack_id As Double
'Dim Report3 As New CrystalReport3
Dim strVar As String

Private Sub cboparticular_Click()
      Call GetAccName(Me, Trim(cboparticular))
End Sub

Private Sub cboparticular_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    cboparticular_LostFocus
  End If
End Sub

Private Sub cboparticular_LostFocus()
   If Len(Trim(cboparticular.Text)) = 0 Then Exit Sub

    Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "select acc_name from acct where user_acc='" & Trim(cboparticular.Text) & "' and acc_code not in(select acc_head from acct)"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
       txtparticular_name.Text = Adodc3.Recordset!acc_name
    Else
        MSFlexGrid3.Left = cboparticular.Left
        MSFlexGrid3.Top = cboparticular.Top
        MSFlexGrid3.TabIndex = cboparticular.TabIndex + 1
        Call getparticular_code(Trim(cboparticular.Text))
       Exit Sub
    End If
 
End Sub
Private Sub getparticular_code(strAcc_des As String)
  On Error GoTo err_loop
    MSFlexGrid3.Clear
    MSFlexGrid3.Rows = 0
    
    MSFlexGrid3.ColWidth(0) = "1200"
    MSFlexGrid3.ColAlignment(0) = 1
    
    MSFlexGrid3.ColWidth(1) = "5800"
    
    Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "select user_acc,acc_name from acct where acc_code not in(select acc_head from acct) and upper(acc_name) like '" & Trim(UCase(strAcc_des)) & "%'"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
        Do Until Adodc3.Recordset.EOF
            MSFlexGrid3.AddItem Adodc3.Recordset!user_acc & vbTab & Adodc3.Recordset!acc_name
            Adodc3.Recordset.MoveNext
       Loop
    End If
    
    MSFlexGrid3.Visible = True
    MSFlexGrid3.SetFocus
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next

End Sub
Private Sub cboUnit_Click()

    Call GetUnitCode
'    Call GetUnitCode(Me, Trim(Me.cboUnit))
    
End Sub

Private Sub cboUserAcc_Click()

    Call GetAccName(Me, Trim(Me.cboUserAcc.Text))
    
End Sub

Private Sub cboVou_Type_LostFocus()

    On Error GoTo err_sub
    Dim track_id As Integer
    Adodc1.ConnectionString = strcn.Connection_String
    ''Adodc1.RecordSource = "Select max(track_id)as track_id from vou"
    Adodc1.RecordSource = "Select max(to_number(vou_no))as track_id from vou "
'    where vou_type='" & Trim(Me.cboVou_Type.Text) & "'"
    Adodc1.Refresh
    If IsNull(Adodc1.Recordset!track_id) = True Then
         txtVOU_NO.Text = 1
    Else
'    If Adodc1.Recordset.RecordCount > 0 Then
        txtVOU_NO.Text = Adodc1.Recordset!track_id + 1
'        track_id = Adodc1.Recordset!track_id
'
'        Adodc1.ConnectionString = strcn.Connection_String
'        Adodc1.RecordSource = "Select vou_no from vou where  track_id =" & Val(track_id) & ""
'        Adodc1.Refresh
'        If Adodc1.Recordset.RecordCount > 0 Then
'           txtVOU_NO.Text = Adodc1.Recordset!vou_no
'        End If
 
       
    End If
    
    
    
'    Con.ConnectionTimeout = 0
'    Con.Open strcn.Connection_String
'    RS.Open "select max(track_id) from vou", Con.Connection_String
'    If RS.EOF = False Then
'       txtVOU_NO.Text = (RS!lst_vou_no)
'    Else
'       txtVOU_NO.Text = 1
'    End If
'    RS.Close
'    Con.Close
    Exit Sub
err_sub:
    MsgBox Err.Description, vbCritical
    Resume Next
    
End Sub

Private Sub cmdDELETE_Click()

    If Len(Trim(txtVOU_NO.Text)) = 0 Then Exit Sub
       On Error GoTo err_loop
       Dim reply As String
       reply = MsgBox("Do you want to delete?", vbCritical + vbYesNo, "Warning...")
       If reply = vbYes Then
            Dim Conn As New Connection
            Dim cmd As New Command
            Dim RS As New Recordset
            Dim unitcode As String
            unitcode = Trim(txtUnitCode.Text)
            
            Conn.Open strcn.Connection_String
            
            Set cmd.ActiveConnection = Conn
            
            cmd.CommandType = adCmdText
            cmd.CommandText = "Delete from vou where vou_no='" & Trim(Me.txtVOU_NO.Text) & "' and vou_type='" & Trim(cboVou_Type.Text) & "'"
            RS.CursorLocation = adUseClient
            RS.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
            Call delete_ledger
            Call flush_grd
            txtDiff = "0.00"
            nbrDebit = "0.00"
            nbrCredit = "0.00"
            
            
       End If
       Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub
Private Sub delete_ledger()
            Dim Conn As New Connection
            Dim cmd As New Command
            Dim RS As New Recordset
            Dim unitcode As String
            unitcode = Trim(txtUnitCode.Text)
            
            Conn.Open strcn.Connection_String
            
            Set cmd.ActiveConnection = Conn
            
            cmd.CommandType = adCmdText
            cmd.CommandText = "Delete from ledger where vou_no='" & Trim(Me.txtVOU_NO.Text) & "' and vou_type='" & Trim(cboVou_Type.Text) & "'"
            RS.CursorLocation = adUseClient
            RS.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
            
End Sub
Private Sub UPDATE_POST_STATE()
            Dim Conn As New Connection
            Dim cmd As New Command
            Dim RS As New Recordset
            Dim unitcode As String
            unitcode = Trim(txtUnitCode.Text)
            
            Conn.Open strcn.Connection_String
            
            Set cmd.ActiveConnection = Conn
            
            cmd.CommandType = adCmdText
            cmd.CommandText = "UPDATE VOU SET POST_STATE=0 where vou_no='" & Trim(Me.txtVOU_NO.Text) & "' and vou_type='" & Trim(cboVou_Type.Text) & "'"
            RS.CursorLocation = adUseClient
            RS.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
            
End Sub

Private Sub cmdSAVE_Click()

    If Len(Trim(txtVOU_NO.Text)) = 0 Then
       MsgBox "Voucher no. required", vbCritical
       txtVOU_NO.SetFocus
       Exit Sub
    End If
    
    If Val(nbrtot_debit) <> Val(nbrtot_credit) Then
       MsgBox "Debit and Credit not equal", vbCritical
       txtVOU_NO.SetFocus
       Exit Sub
    End If
    
         On Error GoTo err_loop
    '''''---------------------------------------------------------------------------
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter



    
    Dim userid As String
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 40, txtVOU_NO.Text)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 500, cboVou_Type.Text)
    cmd.Parameters.Append Param2
    
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL PostVou(?, ?)}"
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

    '''''-----------------------------------------------------------------------------
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Posted..."
    Call flush_grd
    txtVOU_NO.SetFocus
    txtDiff = "0.00"
    nbrDebit = "0.00"
    nbrCredit = "0.00"
    Call cboVou_Type_LostFocus
            
  Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub





Private Sub Command1_Click(Index As Integer)
  Select Case Index
         Case 0
            txtVOU_NO = txtVOU_NO + 1
            txtVOU_NO_LostFocus
         Case 1
            txtVOU_NO = txtVOU_NO - 1
            txtVOU_NO_LostFocus
  End Select
            
End Sub

Private Sub dtVOU_DATE_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case Index
          Case 0
            If KeyCode = 13 Then
                SendKeys Chr(9)
            End If
  End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF5 Then
       Call cmdADD_Click
    End If
    
End Sub

Private Sub Label11_Click()
    Form4.Show vbModal
End Sub





Private Sub Label11_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Label11.ForeColor = vbWhite
End Sub

Private Sub mnuDeleteOne_Click()

    intTrack_id = Val(MSFlexGrid1.Text)
    txtTrack_ID.Text = intTrack_id
    If Val(intTrack_id) <= 0 Then Exit Sub

    On Error GoTo err_sub
            Dim reply As String
       reply = MsgBox("Do you want to delete?", vbCritical + vbYesNo, "Warning...")
       If reply = vbYes Then
            Dim Conn As New Connection
            Dim cmd As New Command
            Dim RS As New Recordset
            Dim unitcode As String
            unitcode = Trim(txtUnitCode.Text)
            
            Conn.Open strcn.Connection_String
            
            Set cmd.ActiveConnection = Conn
            
            cmd.CommandType = adCmdText
            cmd.CommandText = "Delete from vou where track_id=" & Val(txtTrack_ID.Text) & ""
            RS.CursorLocation = adUseClient
            RS.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
            
            Call flush_grd
            delete_ledger
            UPDATE_POST_STATE
       End If
     
        Exit Sub
err_sub:
        MsgBox "Error occurs", vbCritical
        Resume Next
        
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then
       PopupMenu mnuDelete, 2
       Call flush_grd
    End If
    
End Sub

Private Sub MSFlexGrid2_DblClick()

    If Len(Trim(MSFlexGrid2.Text)) <> 0 Then
       cboUserAcc.Text = MSFlexGrid2.Text
       Call cboUserAcc_LostFocus
       nbrDebit.SetFocus
              
    Else
       cboUserAcc.SetFocus
    End If
    MSFlexGrid2.Visible = False
    
End Sub

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
    
End Sub

Private Sub MSFlexGrid2_LostFocus()

    Call MSFlexGrid2_DblClick
    
End Sub

Private Sub cboUserAcc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       If Len(Trim(cboUserAcc.Text)) = 0 Then
          cmdSAVE.SetFocus
       Else
          SendKeys Chr(9)
       End If
    End If
    
End Sub

Private Sub cboVou_Type_Click()

    Select Case cboVou_Type
        Case "JV"
            lbl_chq(0).Visible = False
            txtvou_Chq.Visible = False
            lbl_Vou_Cap.Caption = "Journal Voucher"
             dtVOU_DATE(1).Visible = False
            Label19(1).Visible = False
        Case "CP"
            lbl_chq(0).Visible = False
            txtvou_Chq.Visible = False
            lbl_Vou_Cap.Caption = "Cash Payment Voucher"
             dtVOU_DATE(1).Visible = False
            Label19(1).Visible = False
        Case "CR"
            lbl_chq(0).Visible = False
            txtvou_Chq.Visible = False
            lbl_Vou_Cap.Caption = "Cash Receipt Voucher"
             dtVOU_DATE(1).Visible = False
            Label19(1).Visible = False
        Case "BP"
            lbl_chq(0).Visible = True
            txtvou_Chq.Visible = True
            lbl_Vou_Cap.Caption = "Bank Payment Voucher"
            dtVOU_DATE(1).Visible = True
            Label19(1).Visible = True
        Case "BR"
            lbl_chq(0).Visible = True
            txtvou_Chq.Visible = True
            lbl_Vou_Cap.Caption = "Bank Receipt Voucher"
             dtVOU_DATE(1).Visible = True
            Label19(1).Visible = True
    End Select
    
End Sub

Private Sub cmdADD_Click()

     Call ClearScreen
     cboVou_Type_LostFocus
     txtVOU_NO.SetFocus
     
End Sub


Private Sub cmdEdit_Click()
    Exit Sub
    Call edit
    
End Sub

Private Sub cmdEXIT_Click()
    Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdPREVIEW_Click()

    rptMode = 3
    CRViewer1.Show vbModal
    
End Sub

Private Sub cmdPrint_Click()
    
'    Report3.DiscardSavedData
'    RS.Open "exec MR '" & Trim(txtvou_no.Text) & "','" & Trim(cboVou_Type.Text) & "'", strcn
'    Report3.txtCompName.SetText objectCompSetup.comp_name
'    Report3.txtCompAddr.SetText objectCompSetup.comp_addr1
'    Report3.TakaInWord.SetText ConvertX(Val(nbrtot_debit.Text))
'    Report3.Database.SetDataSource RS
'    Report3.PrintOut
'    RS.Close
    
End Sub
Private Sub ClearScreen()

     txtVOU_NO.Text = ""
'     dtVOU_DATE.Value = Date
     txtvou_Chq.Text = ""
     txtVOU_NARR.Text = ""
     cboUnit.Text = ""
     Me.txtCustOrdNo.Text = ""
     cboUserAcc.Text = ""
     txtacc_name = ""
    '' Me.nbrDollar.Text = "0.00"
    '' Me.nbrRate.Text = "0.00"
     nbrDebit.Text = "0.00"
     nbrCredit.Text = "0.00"
     nbrtot_debit.Text = "0.00"
     nbrtot_credit.Text = "0.00"
     intTrack_id = 0
     
End Sub

Private Sub Form_Load()

    Call ClearScreen
    
'    objectCompSetup.Flush_Comp (strcn)
'
'    dtVOU_DATE.Value = Date
'
'    Call format_grd
'    Call flush_grd
'    Call GetUnitName(Me)
'    Call GetUserAcc(Me)
    Call GetUnitName
    cboVou_Type.AddItem "JV"
    cboVou_Type.AddItem "CP"
    cboVou_Type.AddItem "CR"
    cboVou_Type.AddItem "BP"
    cboVou_Type.AddItem "BR"
    
    cboVou_Type.Text = "JV"
    
End Sub
Private Sub GetUnitName()
    
    Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "Select prj_name from project"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
        Adodc3.Recordset.MoveFirst
        While Adodc3.Recordset.EOF = False
            cboUnit.AddItem Adodc3.Recordset!prj_name
            Adodc3.Recordset.MoveNext
        Wend
    End If

End Sub
Private Sub GetUnitCode()
    
    Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "Select prj_code from project where prj_name='" & Trim(cboUnit.Text) & "'"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
       txtUnitCode.Text = Adodc3.Recordset!prj_code
    End If

End Sub
Private Sub MSFlexGrid1_DblClick()

    If Len(Trim(MSFlexGrid1.Text)) = 0 Then Exit Sub
    
    Dim strUserAcc As String
    Dim strUnitName As String
    
    If Val(Me.MSFlexGrid1.Text) = 0 Then Exit Sub
    intTrack_id = Val(MSFlexGrid1.Text)
    txtTrack_ID.Text = intTrack_id
    On Error GoTo err_sub
    Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "select vou_no,vou_type,vou_date,vou_narr,acc_code,(select user_acc from acct where acct.acc_code=vou.acc_code) as user_acc,(select acc_name from acct where acct.acc_code=vou.acc_code) as acc_name,dr_amt,cr_amt,vou_chq from vou where vou_no='" & Trim(txtVOU_NO.Text) & "' and vou_type='" & Trim(cboVou_Type) & "' and track_id=" & Val(intTrack_id) & ""
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
        cboVou_Type.Text = Adodc3.Recordset!vou_type
            txtVOU_NO = Adodc3.Recordset!vou_no
            dtVOU_DATE(0).Value = Adodc3.Recordset!vou_date
            txtVOU_NARR.Text = "" & Adodc3.Recordset!vou_narr
            cboUserAcc.Text = Adodc3.Recordset!user_acc
            txtacc_name = Adodc3.Recordset!acc_name
'            Me.cboUnit = Adodc3.Recordset!prj_name
            nbrDebit.Text = Adodc3.Recordset!dr_amt
            nbrCredit.Text = Adodc3.Recordset!cr_amt
            If IsNull(Adodc3.Recordset!vou_chq) = False Then txtvou_Chq.Text = "" & Adodc3.Recordset!vou_chq
    End If
    Exit Sub
err_sub:
    MsgBox Err.Description
    Resume Next
    
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       Call MSFlexGrid1_DblClick
    End If
    
End Sub

Private Sub MSFlexGrid3_DblClick()
   If Len(Trim(MSFlexGrid3.Text)) <> 0 Then
       cboparticular.Text = MSFlexGrid3.Text
       Call cboparticular_LostFocus
       cboUserAcc.SetFocus
       'nbrDollar.SetFocus
    Else
       cboparticular.SetFocus
    End If
    MSFlexGrid3.Visible = False

End Sub

Private Sub MSFlexGrid3_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    MSFlexGrid3_DblClick
 End If
End Sub

Private Sub MSFlexGrid3_LostFocus()
        MSFlexGrid3_DblClick
End Sub

Private Sub nbrCredit_GotFocus()

    nbrCredit.SelLength = Len(Trim(nbrCredit.Text))
    
End Sub

Private Sub nbrCredit_KeyPress(KeyAscii As Integer)

    If KeyAscii > 26 Then
       If InStr("0123456789.", Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
    
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
    
End Sub

Private Sub nbrCredit_LostFocus()
  
    If IsNumeric(nbrCredit.Text) = False Then
       MsgBox "Accepts numeric values", vbCritical, "IT Division,DNMIH"
       nbrCredit.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txtVOU_NO.Text)) = 0 Then
       MsgBox "Voucher no. required", vbCritical, "IT Division,DNMIH"
       txtVOU_NO.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(cboUserAcc.Text)) = 0 Then
       MsgBox "Accounts code required ", vbCritical, "IT Division,DNMIH"
       cboUserAcc.SetFocus
       Exit Sub
    End If
    
    If Val(nbrDebit.Text) = 0 And Val(nbrCredit.Text) = 0 Then
       cboUserAcc.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(cboparticular.Text)) = 0 Then
       MsgBox "Contra Account code required ", vbCritical, "IT Division,DNMIH"
       cboparticular.SetFocus
       Exit Sub
    End If
    
    If UCase(cboVou_Type) = UCase("BP") Or UCase(cboVou_Type) = UCase("BR") Then
       If Len(Trim(txtvou_Chq)) = 0 Then
          MsgBox "Cheque No Mandatory...Please Put a Cheque No", vbCritical + vbDefaultButton1, "IT Division,DNMIH"
          txtvou_Chq.SetFocus
          Exit Sub
        End If
    End If
       
       
    Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "Select vou_no,vou_date from vou where  vou_no='" & Trim(txtVOU_NO.Text) & "'"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
       dtVOU_DATE(0).Value = Adodc3.Recordset!vou_date
    End If


    Call Save
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
    nbrDebit.Text = "0.00"
    nbrCredit.Text = "0.00"
    cboUserAcc.Text = ""
    txtacc_name.Text = ""
    intTrack_id = 0
    txtTrack_ID.Text = 0
    cboUserAcc.SetFocus
    
End Sub

Private Sub nbrDebit_GotFocus()

    nbrDebit.SelLength = Len(Trim(nbrDebit.Text))
    
End Sub

Private Sub nbrDebit_KeyPress(KeyAscii As Integer)

    If KeyAscii > 26 Then
       If InStr("0123456789.", Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
    
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If

End Sub
Private Sub cboUserAcc_LostFocus()

    If Len(Trim(cboUserAcc.Text)) = 0 Then Exit Sub

    Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "select acc_name from acct where user_acc='" & Trim(cboUserAcc.Text) & "' and acc_code not in(select acc_head from acct)"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
       txtacc_name.Text = Adodc3.Recordset!acc_name
    Else
        MSFlexGrid2.Left = cboUserAcc.Left
        MSFlexGrid2.Top = cboUserAcc.Top
        MSFlexGrid2.TabIndex = cboUserAcc.TabIndex + 1
        Call getAcc_Code(Trim(cboUserAcc.Text))
       Exit Sub
    End If
    
    
End Sub

Private Sub nbrDollar_Change()

    Call cal_dollar
    
End Sub

Private Sub nbrDollar_GotFocus()

    nbrDollar.SelLength = Len(Trim(nbrDollar.Text))
    
End Sub

Private Sub nbrDollar_KeyPress(KeyAscii As Integer)

    If KeyAscii > 26 Then
       If InStr("0123456789.", Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
    
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
    
End Sub

Private Sub nbrDollar_LostFocus()

    Call cal_dollar
    
End Sub

Private Sub nbrRate_Change()
    Call cal_dollar
End Sub

Private Sub nbrRate_GotFocus()
    nbrRate.SelLength = Len(Trim(nbrRate.Text))
End Sub

Private Sub nbrRate_KeyPress(KeyAscii As Integer)
    If KeyAscii > 26 Then
       If InStr("0123456789.", Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
    
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub nbrRate_LostFocus()
    Call cal_dollar
End Sub

Private Sub Option1_Click(Index As Integer)
   If Option1(0).Value = True Then
      chequeStatus = 1
   ElseIf Option1(1).Value = True Then
       chequeStatus = 2
   End If
End Sub

Private Sub txtvou_Chq_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub txtVOU_NARR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
       KeyAscii = Asc(Chr(96))
    End If
End Sub

Private Sub txtVOU_NO_GotFocus()
    txtVOU_NO.SelStart = Len(Trim(txtVOU_NO.Text))
End Sub

Private Sub txtvou_no_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

'    If KeyAscii > 26 Then
'       If InStr("0123456789", Chr(KeyAscii)) = 0 Then
'          KeyAscii = 0
'       End If
'    End If

    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub txtVOU_NO_LostFocus()
    If Len(Trim(txtVOU_NO.Text)) = 0 Then Exit Sub
    Dim tot_dr, tot_cr, tot_diff As Double
    Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "Select vou_no,vou_type,vou_date,vou_narr,vou_chq,prj_code,(select prj_name from project where vou.prj_code=Project.prj_code) as prj_name,check_date from vou where  vou_no='" & Trim(txtVOU_NO.Text) & "'"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
        cboVou_Type.Text = Adodc3.Recordset!vou_type
        dtVOU_DATE(0).Value = Adodc3.Recordset!vou_date
       If dtVOU_DATE(1).Visible = True Then
        If Len((Adodc3.Recordset!check_date)) > 0 Then
           dtVOU_DATE(1).Value = Adodc3.Recordset!check_date
           End If
        End If
        txtVOU_NARR.Text = "" & Adodc3.Recordset!vou_narr
        txtUnitCode.Text = Adodc3.Recordset!prj_code
        txtvou_Chq = "" & Adodc3.Recordset!vou_chq
        cboVou_Type = Adodc3.Recordset!vou_type
        If IsNull(Adodc3.Recordset!prj_name) = False Then cboUnit.Text = Adodc3.Recordset!prj_name
'        If IsNull(Adodc3.Recordset!vou_chq) = False Then txtvou_Chq.Text = Adodc3.Recordset!vou_chq
               
    Else
        dtVOU_DATE(0).Value = Date
        dtVOU_DATE(1).Value = Date
        txtVOU_NARR.Text = ""
        cboparticular.Text = ""
        txtvou_Chq = ""
    End If
    
       Call flush_grd
    
End Sub

Private Sub Save()
    On Error GoTo err_loop
    '''''---------------------------------------------------------------------------
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    Dim Param0 As New Parameter
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    Dim Param10 As New Parameter
    Dim Param11 As New Parameter
    Dim Param12 As New Parameter
    Dim Param13 As New Parameter
    Dim Param14 As New Parameter


    
    Dim userid As String
    userid = Form1.Label2(2)
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
     Set Param0 = cmd.CreateParameter("param0", adInteger, adParamInput, 40, 1)
     cmd.Parameters.Append Param0

    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 40, txtVOU_NO.Text)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, dtVOU_DATE(0).Value)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 500, txtVOU_NARR.Text)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 40, cboUserAcc.Text)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 9, Val(nbrDebit.Text))
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 9, Val(nbrCredit.Text))
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 4, cboVou_Type.Text)
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 40, txtvou_Chq.Text)
    cmd.Parameters.Append Param8
    
    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 20, "000")
    cmd.Parameters.Append Param9
    
    Set Param10 = cmd.CreateParameter("param10", adInteger, adParamInput, 4, Val(txtTrack_ID.Text))
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 50, userid)
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adVarChar, adParamInput, 40, cboparticular.Text)
    cmd.Parameters.Append Param12
    
    Set Param13 = cmd.CreateParameter("param12", adDate, adParamInput, 10, dtVOU_DATE(1).Value)
    cmd.Parameters.Append Param13
    
    Set Param14 = cmd.CreateParameter("param14", adInteger, adParamInput, 10, chequeStatus)
    cmd.Parameters.Append Param14
   
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL SAVEVOU(?,?, ?, ?, ?, ?, ?, ?, ?,?, ?, ?,?,?,?)}"
    Set RS = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False

    '''''-----------------------------------------------------------------------------
   

    Call flush_grd

  Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub
Private Sub edit()
On Error GoTo err_loop
    '''''---------------------------------------------------------------------------
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    Dim Param0 As New Parameter
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    Dim Param10 As New Parameter
    Dim Param11 As New Parameter
    Dim Param12 As New Parameter
    Dim Param13 As New Parameter
    Dim Param14 As New Parameter


    
    Dim userid As String
    userid = "Emdad"
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
     Set Param0 = cmd.CreateParameter("param0", adInteger, adParamInput, 40, 1)
     cmd.Parameters.Append Param0

    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 40, txtVOU_NO.Text)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, dtVOU_DATE(0).Value)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 500, txtVOU_NARR.Text)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 40, cboUserAcc.Text)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 9, Val(nbrDebit.Text))
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 9, Val(nbrCredit.Text))
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 4, cboVou_Type.Text)
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 40, txtvou_Chq.Text)
    cmd.Parameters.Append Param8
    
    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 20, "000")
    cmd.Parameters.Append Param9
    
    Set Param10 = cmd.CreateParameter("param10", adInteger, adParamInput, 4, Val(txtTrack_ID.Text))
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 50, userid)
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adVarChar, adParamInput, 40, cboparticular.Text)
    cmd.Parameters.Append Param12
    
    Set Param13 = cmd.CreateParameter("param12", adDate, adParamInput, 10, dtVOU_DATE(1).Value)
    cmd.Parameters.Append Param13
    
    Set Param14 = cmd.CreateParameter("param14", adInteger, adParamInput, 10, chequeStatus)
    cmd.Parameters.Append Param14
   
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL SAVEVOU(?,?, ?, ?, ?, ?, ?, ?, ?,?, ?, ?,?,?,?)}"
    Set RS = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False

    '''''-----------------------------------------------------------------------------
   

    Call flush_grd

  Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

Private Sub flush_grd()
    If Len(Trim(txtVOU_NO.Text)) = 0 Then Exit Sub
    Dim tot_dr, tot_cr, tot_diff As Double
    Dim cnt As Integer
    cnt = 0
    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 0
    
    Adodc2.ConnectionString = strcn.Connection_String
    Adodc2.RecordSource = "select vou_no,vou_date,vou_narr,acc_code,(select user_acc from acct where acct.acc_code=vou.acc_code) as user_acc,(select acc_name from acct where acct.acc_code=vou.acc_code) as acc_name,dr_amt,0 as dollar,0 as rate,cr_amt,vou_type,vou_chq,track_id from vou where vou_no='" & Trim(txtVOU_NO.Text) & "' and vou_type='" & Trim(cboVou_Type.Text) & "'"
    Adodc2.Refresh


    If Adodc2.Recordset.RecordCount >= 0 Then
       Do Until Adodc2.Recordset.EOF
          cnt = cnt + 1
          MSFlexGrid1.AddItem Adodc2.Recordset!track_id & vbTab & cnt & vbTab & Adodc2.Recordset!user_acc _
          & vbTab & Adodc2.Recordset!acc_name & vbTab & Adodc2.Recordset!dr_amt & vbTab & Adodc2.Recordset!cr_amt  ''& vbTab & Adodc2.Recordset!dollar & vbTab & Adodc2.Recordset!Rate _

          Adodc2.Recordset.MoveNext
       Loop
            tot_dr = 0
            tot_cr = 0
            tot_diff = 0
            If Adodc2.Recordset.RecordCount > 0 Then
                Adodc2.Recordset.MoveFirst
            End If
            While Adodc2.Recordset.EOF = False
            tot_dr = tot_dr + Adodc2.Recordset!dr_amt
            tot_cr = tot_cr + Adodc2.Recordset!cr_amt
            Adodc2.Recordset.MoveNext
            Wend
            nbrtot_debit.Text = Val(tot_dr)
            nbrtot_credit.Text = Val(tot_cr)
            tot_diff = Val(tot_dr) - Val(tot_cr)
            txtDiff.Text = Val(tot_diff)
    Else
            nbrtot_debit.Text = 0
            nbrtot_credit.Text = 0
            tot_diff = 0
            txtDiff.Text = 0
       
    End If
       
    Call format_grd
End Sub

Private Sub format_grd()
    MSFlexGrid1.Cols = 8
    MSFlexGrid1.Rows = 100
    
    MSFlexGrid1.ColWidth(0) = 0
    
    MSFlexGrid1.ColWidth(1) = 500
    MSFlexGrid1.ColAlignment(1) = 1
    
    MSFlexGrid1.ColWidth(2) = 1290
    MSFlexGrid1.ColAlignment(2) = 1
    
    MSFlexGrid1.ColWidth(3) = 5600 '''3855
    
    MSFlexGrid1.ColWidth(4) = 1210
    MSFlexGrid1.ColWidth(5) = 1310 ''''+ 3855 + 1110
    
    MSFlexGrid1.ColWidth(6) = 1290
    MSFlexGrid1.ColWidth(7) = 1290
    
    MSFlexGrid1.GridLines = flexGridRaised
End Sub

Private Sub getAcc_Code(strAcc_des As String)
    On Error GoTo err_loop
    MSFlexGrid2.Clear
    MSFlexGrid2.Rows = 0
    
    MSFlexGrid2.ColWidth(0) = "1200"
    MSFlexGrid2.ColAlignment(0) = 1
    
    MSFlexGrid2.ColWidth(1) = "5800"
    
    Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "select user_acc,acc_name from acct where acc_code not in(select acc_head from acct) and upper(acc_name) like '" & Trim(UCase(strAcc_des)) & "%'"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
        Do Until Adodc3.Recordset.EOF
            MSFlexGrid2.AddItem Adodc3.Recordset!user_acc & vbTab & Adodc3.Recordset!acc_name
            Adodc3.Recordset.MoveNext
       Loop
    End If
    
    MSFlexGrid2.Visible = True
    MSFlexGrid2.SetFocus
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

Private Sub cal_dollar()
    If IsNumeric(nbrDollar.Text) = True And IsNumeric(nbrRate.Text) = True Then
        nbrDebit.Text = Round((Val(nbrDollar.Text) * Val(nbrRate.Text)), 2)
        nbrCredit.Text = Round((Val(nbrDollar.Text) * Val(nbrRate.Text)), 2)
    End If
End Sub

