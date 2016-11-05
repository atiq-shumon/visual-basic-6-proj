VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form22 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheque Register"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   FillStyle       =   0  'Solid
   Icon            =   "frmChq_reg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   1410
      Left            =   6870
      TabIndex        =   27
      Top             =   -900
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
   Begin VB.TextBox txtBill_Amt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6270
      MaxLength       =   12
      TabIndex        =   12
      Text            =   "0.00"
      Top             =   2430
      Width           =   1290
   End
   Begin VB.TextBox txtacc_name 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   1410
      Locked          =   -1  'True
      MaxLength       =   45
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2430
      Width           =   4905
   End
   Begin VB.ComboBox cboUserAcc 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   120
      Style           =   1  'Simple Combo
      TabIndex        =   11
      Top             =   2430
      Width           =   1275
   End
   Begin VB.ComboBox cboUserAcc 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   120
      Style           =   1  'Simple Combo
      TabIndex        =   10
      Top             =   1890
      Width           =   1275
   End
   Begin VB.TextBox txtacc_name 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1410
      Locked          =   -1  'True
      MaxLength       =   45
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1890
      Width           =   4755
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000009&
      Caption         =   "Cheque Cancel"
      Height          =   315
      Left            =   6360
      TabIndex        =   33
      Top             =   1530
      Width           =   1605
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   735
      Left            =   -60
      TabIndex        =   31
      Top             =   -120
      Width           =   9315
      Begin VB.Label lbl_Vou_Cap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque Register(Original Amount)"
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
         Left            =   210
         TabIndex        =   32
         Top             =   210
         Width           =   5280
      End
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000009&
      Caption         =   "Original Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   195
      Index           =   3
      Left            =   510
      TabIndex        =   0
      Top             =   720
      Value           =   -1  'True
      Width           =   1755
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000009&
      Caption         =   "Security"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   195
      Index           =   2
      Left            =   6090
      TabIndex        =   3
      Top             =   720
      Width           =   2085
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000009&
      Caption         =   "VAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   195
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   720
      Width           =   1155
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000009&
      Caption         =   "Income Tax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   195
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   6210
      TabIndex        =   30
      Top             =   1050
      Width           =   2985
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000009&
         Caption         =   "Received"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000009&
         Caption         =   "Paid"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   2
         Left            =   2160
         TabIndex        =   9
         Top             =   270
         Width           =   675
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000009&
         Caption         =   "Unpaid"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   1
         Left            =   1200
         TabIndex        =   8
         Top             =   270
         Width           =   1065
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5580
      Top             =   4650
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
      Left            =   1410
      MaxLength       =   25
      TabIndex        =   15
      Top             =   2970
      Width           =   4935
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
      Left            =   120
      Picture         =   "frmChq_reg.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Posting to Ledger"
      Top             =   3810
      Width           =   465
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
      Left            =   1065
      Picture         =   "frmChq_reg.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Edit"
      Top             =   3810
      Width           =   465
   End
   Begin VB.CommandButton cmdPreview 
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
      Left            =   2460
      Picture         =   "frmChq_reg.frx":0DB6
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Preview"
      Top             =   3810
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
      Left            =   1995
      Picture         =   "frmChq_reg.frx":1420
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Print"
      Top             =   3810
      Width           =   465
   End
   Begin VB.TextBox txtvou_Chq 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1410
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1350
      Width           =   3270
   End
   Begin VB.TextBox txt_nature 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   750
      Left            =   7590
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   2400
      Width           =   1515
   End
   Begin VB.TextBox txtSRL 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1350
      Width           =   1275
   End
   Begin MSComCtl2.DTPicker dtISS_DATE 
      Height          =   285
      Index           =   0
      Left            =   4710
      TabIndex        =   6
      Top             =   1350
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-mon-yyyy"
      Format          =   19857409
      CurrentDate     =   38416
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
      Left            =   1530
      Picture         =   "frmChq_reg.frx":1A8A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Delete"
      Top             =   3810
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
      Left            =   600
      Picture         =   "frmChq_reg.frx":25C4
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "New"
      Top             =   3810
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
      Left            =   2925
      Picture         =   "frmChq_reg.frx":2C2E
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Exit"
      Top             =   3810
      Width           =   465
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5730
      Top             =   4620
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
      Left            =   5880
      Top             =   4740
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
      Left            =   120
      TabIndex        =   14
      Top             =   2970
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-mon-yyyy"
      Format          =   19857409
      CurrentDate     =   36955
   End
   Begin VB.Label lblCR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   6930
      TabIndex        =   40
      Top             =   2160
      Width           =   570
   End
   Begin VB.Label lblDR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Acc.Code"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   39
      Top             =   2205
      Width           =   1065
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   0
      Left            =   1410
      TabIndex        =   38
      Top             =   2205
      Width           =   345
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   1
      Left            =   1410
      TabIndex        =   37
      Top             =   1680
      Width           =   345
   End
   Begin VB.Label lblDR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Acc.Code"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   36
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      Height          =   525
      Left            =   60
      Top             =   3750
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Left            =   -30
      Top             =   630
      Width           =   9255
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   1
      Left            =   90
      TabIndex        =   29
      Top             =   2730
      Width           =   915
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Challan No #"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   1410
      TabIndex        =   28
      Top             =   2790
      Width           =   915
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H00000000&
      FillColor       =   &H80000001&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   2
      Left            =   -30
      Top             =   3660
      Width           =   10140
   End
   Begin VB.Label lbl_chq 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque#"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   0
      Left            =   1410
      TabIndex        =   26
      Top             =   1080
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purpose :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   7650
      TabIndex        =   25
      Top             =   2130
      Width           =   645
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   0
      Left            =   4710
      TabIndex        =   24
      Top             =   1110
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No #"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   120
      TabIndex        =   23
      Top             =   1110
      Width           =   765
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   3045
      Index           =   0
      Left            =   -30
      Top             =   600
      Width           =   9270
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteOne 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strMode As String
Dim intTrack_id As Double
Dim strVar As String
'Private Sub cboUserAcc_Click(Index As Integer)
'  Select Case Index
'         Case 0
'
'                Call GetAccName(Me, Trim(Me.cboUserAcc(0).Text))
'
'         Case 1
'
'               Call GetAccName(Me, Trim(Me.cboUserAcc(1).Text))
'
'          End Select
'
'End Sub


Private Sub cboUserAcc_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
           Case 0
                  
                   If KeyAscii = 13 Then
                       If Len(Trim(cboUserAcc(0).Text)) = 0 Then
                            cboUserAcc(1).SetFocus
                        Else
                                  SendKeys Chr(9)
                        End If
                    End If
                     flex_grid_var = 0
             Case 1
                    
                   If KeyAscii = 13 Then
                       If Len(Trim(cboUserAcc(1).Text)) = 0 Then
                            txtBill_Amt.SetFocus
                        Else
                                  SendKeys Chr(9)
                        End If
                    End If
                     flex_grid_var = 1
                    
             End Select
        
End Sub

Private Sub cboUserAcc_LostFocus(Index As Integer)
    Select Case Index
           Case 0
                 
                If Len(Trim(cboUserAcc(0).Text)) = 0 Then Exit Sub

                Adodc3.ConnectionString = strcn.Connection_String
                Adodc3.RecordSource = "select acc_name from acct where user_acc='" & Trim(cboUserAcc(0).Text) & "'"
                Adodc3.Refresh
                    If Adodc3.Recordset.RecordCount > 0 Then
                        txtacc_name(0).Text = Adodc3.Recordset!acc_name
                    Else
                        MSFlexGrid2.Left = cboUserAcc(0).Left
                        MSFlexGrid2.Top = cboUserAcc(0).Top
                        MSFlexGrid2.TabIndex = cboUserAcc(0).TabIndex + 1
                        Call getAcc_Code(Trim(cboUserAcc(0).Text))
                       
                       Exit Sub
                    End If
                    
             Case 1
                 
                 If Len(Trim(cboUserAcc(1).Text)) = 0 Then Exit Sub

                Adodc3.ConnectionString = strcn.Connection_String
                Adodc3.RecordSource = "select acc_name from acct where user_acc='" & Trim(cboUserAcc(1).Text) & "'"
                Adodc3.Refresh
                    If Adodc3.Recordset.RecordCount > 0 Then
                        txtacc_name(1).Text = Adodc3.Recordset!acc_name
                    Else
                        MSFlexGrid2.Left = cboUserAcc(1).Left
                        MSFlexGrid2.Top = cboUserAcc(1).Top
                        MSFlexGrid2.TabIndex = cboUserAcc(1).TabIndex + 1
                        Call getAcc_Code(Trim(cboUserAcc(1).Text))
                       
                       Exit Sub
                    End If
                    
    End Select
End Sub

Private Sub cmdDELETE_Click()
  Dim validation As Variant
   If cboUserAcc(1).Text = "" Then
        MsgBox "please select a Bank Account ", vbCritical, "IT Division,DNMIH"
        cboUserAcc(0).SetFocus
        Exit Sub
    End If
    
    If txtvou_Chq.Text = "" Then
        MsgBox "please put Check No ", vbCritical, "IT Division,DNMIH"
        txtvou_Chq.SetFocus
        Exit Sub
    End If
    
    validation = MsgBox("Are you sure to Delete?", vbYesNo + vbInformation, "IT Division,DNMIH")
    If validation = vbYes Then
                Dim Conn As New ADODB.Connection
                Dim cmd As New ADODB.Command
                Dim RS As New ADODB.Recordset
            
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
                Dim Param15 As New Parameter
                Dim Param16 As New Parameter
            
            
            
                
                Dim userid As String
                Conn.Open strcn.Connection_String
                
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                
                '----------------------------------------------------------------------------------
                Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 4, 3) ''p_mode
                cmd.Parameters.Append Param1
                
                Set Param2 = cmd.CreateParameter("param2", adInteger, adParamInput, 10, txtSRL.Text)
                cmd.Parameters.Append Param2
                
                Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 15, dtISS_DATE(0).Value)
                cmd.Parameters.Append Param3
                
                Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 100, txt_nature.Text)
                cmd.Parameters.Append Param4
                
                Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 10, chequeReg_Val)
                cmd.Parameters.Append Param5
                
                Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 15, txtvou_Chq)
                cmd.Parameters.Append Param6
                
                Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 30, txtCustOrdNo.Text)
                cmd.Parameters.Append Param7
                
                Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 15, cboUserAcc(0).Text)
                cmd.Parameters.Append Param8
                    
                Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 30, txtBill_Amt.Text)
                cmd.Parameters.Append Param10
                
                Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 30, "Emdad")
                cmd.Parameters.Append Param11
                      
                Set Param13 = cmd.CreateParameter("param13", adInteger, adParamInput, 30, rec_pay)
                cmd.Parameters.Append Param13
                    
                 Set Param14 = cmd.CreateParameter("param14", adInteger, adParamInput, 30, Check1.Value)
                cmd.Parameters.Append Param14
                
                Set Param15 = cmd.CreateParameter("param15", adInteger, adParamInput, 15, cboUserAcc(1).Text)
                cmd.Parameters.Append Param15
                
                
                Set Param16 = cmd.CreateParameter("param16", adDate, adParamInput, 15, dtISS_DATE(0).Value)
                cmd.Parameters.Append Param16
                
                '----------------------------------------------------------------------------------
            
                cmd.Properties("PLSQLRSet") = True
                
                cmd.CommandText = "{CALL s_U_D_CHQ_Reg(?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
                Set RS = cmd.Execute
                
            
                cmd.Properties("PLSQLRSet") = False
            
                '''''-----------------------------------------------------------------------------
                MsgBox "Operation successfull", vbInformation + vbOKOnly, "Delete..."
                
                txtSRL.SetFocus
              
                
                txtBill_Amt = "0.00"
   End If
               
                        
  Exit Sub

End Sub

Private Sub cmdEdit_Click()


    If cboUserAcc(1).Text = "" Then
        MsgBox "please select a Bank Account ", vbCritical, "IT Division,DNMIH"
        cboUserAcc(0).SetFocus
        Exit Sub
    End If
    
    If txtvou_Chq.Text = "" Then
        MsgBox "please put Check No ", vbCritical, "IT Division,DNMIH"
        txtvou_Chq.SetFocus
        Exit Sub
    End If
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

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
    Dim Param15 As New Parameter
    Dim Param16 As New Parameter



    
    Dim userid As String
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 4, 2) ''p_mode
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adInteger, adParamInput, 10, txtSRL.Text)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 15, dtISS_DATE(0).Value)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 100, txt_nature.Text)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 10, chequeReg_Val)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 15, txtvou_Chq)
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 30, txtCustOrdNo.Text)
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 15, cboUserAcc(0).Text)
    cmd.Parameters.Append Param8
        
    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 30, txtBill_Amt.Text)
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 30, "Emdad")
    cmd.Parameters.Append Param11
          
    Set Param13 = cmd.CreateParameter("param13", adInteger, adParamInput, 30, rec_pay)
    cmd.Parameters.Append Param13
        
     Set Param14 = cmd.CreateParameter("param14", adInteger, adParamInput, 30, Check1.Value)
    cmd.Parameters.Append Param14
    
    Set Param15 = cmd.CreateParameter("param15", adInteger, adParamInput, 15, cboUserAcc(1).Text)
    cmd.Parameters.Append Param15
    
    
    Set Param16 = cmd.CreateParameter("param16", adDate, adParamInput, 15, dtVOU_DATE(1).Value)
    cmd.Parameters.Append Param16
    
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL s_U_D_CHQ_Reg(?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

    '''''-----------------------------------------------------------------------------
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Edit..."
    
    txtSRL.SetFocus
  
    
    txtBill_Amt = "0.00"
   
            
  Exit Sub

End Sub

Private Sub CLEAR_FIELDS()
         dtISS_DATE(0).Value = Date
         txt_nature = ""
         Option2(3).Value = True
         cboUserAcc(1).Text = ""
         cboUserAcc(0).Text = ""
         txtCustOrdNo = ""
         txtBill_Amt = ""
         Check1.Value = 0
         Option3(0) = True
         txtacc_name(0) = ""
        txtacc_name(1) = ""
        End Sub


Private Sub cmdSAVE_Click()

    If Len(Trim(txtSRL.Text)) = 0 Then
       MsgBox "Serial  no. required", vbCritical
       txtSRL.SetFocus
       Exit Sub
    End If
    
 
  On Error GoTo err_loop
    '''''---------------------------------------------------------------------------
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

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
    Dim Param15 As New Parameter
    Dim Param16 As New Parameter



    
    Dim userid As String
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 4, 1) ''p_mode
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adInteger, adParamInput, 10, txtSRL.Text)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 15, dtISS_DATE(0).Value)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 100, txt_nature.Text)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 10, chequeReg_Val)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 15, txtvou_Chq)
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 30, txtCustOrdNo.Text)
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 15, cboUserAcc(0).Text)
    cmd.Parameters.Append Param8
        
    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 30, txtBill_Amt.Text)
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 30, "Emdad")
    cmd.Parameters.Append Param11
          
    Set Param13 = cmd.CreateParameter("param13", adInteger, adParamInput, 30, rec_pay)
    cmd.Parameters.Append Param13
        
     Set Param14 = cmd.CreateParameter("param14", adInteger, adParamInput, 30, Check1.Value)
    cmd.Parameters.Append Param14
    
    Set Param15 = cmd.CreateParameter("param15", adInteger, adParamInput, 15, cboUserAcc(1).Text)
    cmd.Parameters.Append Param15
    
    
    Set Param16 = cmd.CreateParameter("param16", adDate, adParamInput, 15, dtVOU_DATE(1).Value)
    cmd.Parameters.Append Param16
    
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL s_U_D_CHQ_Reg(?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

    '''''-----------------------------------------------------------------------------
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
    
    txtSRL.SetFocus
  
    
    txtBill_Amt = "0.00"
   
            
  Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub



Private Sub Command1_Click()
 Form4.Show vbModal
End Sub

Private Sub dtVOU_DATE_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case Index
          Case 0
            If KeyCode = 13 Then
                SendKeys Chr(9)
            End If
  End Select
End Sub



Private Sub MSFlexGrid2_DblClick()
'       flex_grid_var = 1
          If flex_grid_var = 0 Then
                If Len(Trim(MSFlexGrid2.Text)) <> 0 Then
                    cboUserAcc(0).Text = MSFlexGrid2.Text
                     Call cboUserAcc_LostFocus(0)
                       cboUserAcc(1).SetFocus
                     
       
            Else
                
                  cboUserAcc(0).SetFocus
                  
            End If
                    MSFlexGrid2.Visible = False
     ElseIf flex_grid_var = 1 Then
           If Len(Trim(MSFlexGrid2.Text)) <> 0 Then
              cboUserAcc(1).Text = MSFlexGrid2.Text
              Call cboUserAcc_LostFocus(1)
              txtBill_Amt.SetFocus
      
            Else
                cboUserAcc(1).SetFocus
            End If
               MSFlexGrid2.Visible = False
               
    End If

End Sub

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
    
End Sub

Private Sub MSFlexGrid2_LostFocus()

    Call MSFlexGrid2_DblClick
    
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

     txtSRL.Text = ""
     txtvou_Chq.Text = ""
     txt_nature.Text = ""
     Me.txtCustOrdNo.Text = ""
     cboUserAcc(0).Text = ""
     cboUserAcc(1).Text = ""
     txtacc_name(0) = ""
     txtacc_name(1) = ""

    
    
     txtBill_Amt.Text = "0.00"
   
 

     
End Sub

Private Sub Form_Load()
    chequeReg_Val = 4
    Call ClearScreen
   
   
End Sub











Private Sub nbrCredit_GotFocus()

  
    
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



'    dtVOU_DATE.MaxDate = objectCompSetup.ed_dt
'    dtVOU_DATE.MinDate = objectCompSetup.st_dt
    






Private Sub nbrDollar_Change()

    
    
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

    
    
End Sub

Private Sub nbrRate_Change()
    
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
    
End Sub
Private Sub Option2_Click(Index As Integer)
  Select Case Index
         Case 0
                If Option2(0).Value = True Then
                     lbl_Vou_Cap = "Cheque Register(Income Tax)"
                     chequeReg_Val = 1
                End If
         Case 1
                If Option2(1).Value = True Then
                     lbl_Vou_Cap = "Cheque Register(VAT)"
                     chequeReg_Val = 2
                End If
        Case 2
                If Option2(2).Value = True Then
                     lbl_Vou_Cap = "Cheque Register(Security)"
                     chequeReg_Val = 3
                End If
       Case 3
                If Option2(3).Value = True Then
                     lbl_Vou_Cap = "Cheque Register(Original Amount)"
                     chequeReg_Val = 4
                End If
   End Select
                 
End Sub

Private Sub Option3_Click(Index As Integer)
  Select Case Index
         Case 0
               If Option3(0).Value = True Then
                  rec_pay = 0
               End If
        Case 1
             If Option3(1).Value = True Then
                  rec_pay = 1
               End If
         Case 2
             If Option3(2).Value = True Then
                  rec_pay = 2
               End If
       End Select
End Sub

Private Sub txtSRL_GotFocus()
   Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "Select nvl(max(serial_NO),0) as serial_NO from CHEQUE_REG"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
         txtSRL = (Adodc3.Recordset!serial_NO) + 1
         End If
End Sub

Private Sub txtSRL_LostFocus()
     Adodc3.ConnectionString = strcn.Connection_String
     Adodc3.RecordSource = "Select Issue_DATE,nature,ID_NO,cheque_no, Challan_no,party_code,BANK_CODE,Bill_amt,rec_pay_sts,CHK_CANCEL, D_O_Dep from CHEQUE_REG where serial_NO='" & Trim(txtSRL.Text) & "'"
     Adodc3.Refresh
    
    If Adodc3.Recordset.RecordCount > 0 Then
         dtISS_DATE(0).Value = Adodc3.Recordset!Issue_DATE
         txt_nature = "" & Adodc3.Recordset!nature
         If Adodc3.Recordset!ID_NO = 1 Then
            Option2(0).Value = True
          End If
          If Adodc3.Recordset!ID_NO = 2 Then
            Option2(1).Value = True
          End If
          If Adodc3.Recordset!ID_NO = 3 Then
            Option2(2).Value = True
          End If
          If Adodc3.Recordset!ID_NO = 4 Then
            Option2(3).Value = True
          End If
         cboUserAcc(1).Text = Adodc3.Recordset!BANK_CODE
         cboUserAcc(0).Text = Adodc3.Recordset!party_code
         txtCustOrdNo = "" & Adodc3.Recordset!challan_no
         txtBill_Amt = "" & Adodc3.Recordset!Bill_amt
         Check1.Value = "" & Adodc3.Recordset!CHK_CANCEL
         txtvou_Chq.Text = "" & Adodc3.Recordset!cheque_no
           If Adodc3.Recordset!rec_pay_sts = 0 Then
                 Option3(0) = True
            ElseIf Adodc3.Recordset!rec_pay_sts = 1 Then
                 Option3(1) = True
            ElseIf Adodc3.Recordset!rec_pay_sts = 2 Then
                     Option3(2) = True
           End If
    
     Adodc3.ConnectionString = strcn.Connection_String
     Adodc3.RecordSource = "Select acc_name from acct where acc_code='" & Trim(cboUserAcc(0)) & "'"
     Adodc3.Refresh
    
    If Adodc3.Recordset.RecordCount > 0 Then
        txtacc_name(0) = "" & Adodc3.Recordset!acc_name
     End If
     
     
     Adodc3.ConnectionString = strcn.Connection_String
     Adodc3.RecordSource = "Select acc_name from acct where acc_code='" & Trim(cboUserAcc(1)) & "'"
     Adodc3.Refresh
    
    If Adodc3.Recordset.RecordCount > 0 Then
        txtacc_name(1) = "" & Adodc3.Recordset!acc_name
     End If
    
    Else
       txtvou_Chq = ""
       CLEAR_FIELDS
           
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




Private Sub flush_grd()
    If Len(Trim(txtSRL.Text)) = 0 Then Exit Sub
    Dim tot_dr, tot_cr, tot_diff As Double
    Dim cnt As Integer
    cnt = 0
    
   
    
    Adodc2.ConnectionString = strcn.Connection_String
    Adodc2.RecordSource = "select serial_NO,Issue_DATE  ,nature ,party_code ,(select user_acc from acct where acct.acc_code=vou.acc_code) as user_acc,(select acc_name from acct where acct.acc_code=vou.acc_code) as acc_name,dr_amt,0 as dollar,0 as rate,cr_amt,vou_type,vou_chq,track_id from vou where serial_no='" & Trim(txtSRL.Text) & "'"
    Adodc2.Refresh


    If Adodc2.Recordset.RecordCount >= 0 Then
       Do Until Adodc2.Recordset.EOF
          cnt = cnt + 1
        
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
           
   
    End If
       
    
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


Private Sub txtvou_Chq_LostFocus()
    Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "Select serial_NO, Issue_DATE,nature,ID_NO,cheque_no, Challan_no,party_code,BANK_CODE,Bill_amt,rec_pay_sts,CHK_CANCEL, D_O_Dep from CHEQUE_REG where cheque_no='" & Trim(txtvou_Chq) & "'"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
         txtSRL = Adodc3.Recordset!serial_NO
         dtISS_DATE(0).Value = Adodc3.Recordset!Issue_DATE
         txt_nature = "" & Adodc3.Recordset!nature
         If Adodc3.Recordset!ID_NO = 1 Then
            Option2(0).Value = True
          End If
          If Adodc3.Recordset!ID_NO = 2 Then
            Option2(1).Value = True
          End If
          If Adodc3.Recordset!ID_NO = 3 Then
            Option2(2).Value = True
          End If
          If Adodc3.Recordset!ID_NO = 4 Then
            Option2(3).Value = True
          End If
         cboUserAcc(1).Text = Adodc3.Recordset!BANK_CODE
         cboUserAcc(0).Text = Adodc3.Recordset!party_code
         txtCustOrdNo = "" & Adodc3.Recordset!challan_no
         txtBill_Amt = "" & Adodc3.Recordset!Bill_amt
        Check1.Value = "" & Adodc3.Recordset!CHK_CANCEL
         
           If Adodc3.Recordset!rec_pay_sts = 0 Then
                 Option3(0) = True
            ElseIf Adodc3.Recordset!rec_pay_sts = 1 Then
                 Option3(1) = True
            ElseIf Adodc3.Recordset!rec_pay_sts = 2 Then
                     Option3(2) = True
           End If
           
           Adodc3.ConnectionString = strcn.Connection_String
     Adodc3.RecordSource = "Select acc_name from acct where acc_code='" & Trim(cboUserAcc(0)) & "'"
     Adodc3.Refresh
    
    If Adodc3.Recordset.RecordCount > 0 Then
        txtacc_name(0) = "" & Adodc3.Recordset!acc_name
     End If
     
     
     Adodc3.ConnectionString = strcn.Connection_String
     Adodc3.RecordSource = "Select acc_name from acct where acc_code='" & Trim(cboUserAcc(1)) & "'"
     Adodc3.Refresh
    
    If Adodc3.Recordset.RecordCount > 0 Then
        txtacc_name(1) = "" & Adodc3.Recordset!acc_name
     End If
Else
       CLEAR_FIELDS
                     
      End If
End Sub
