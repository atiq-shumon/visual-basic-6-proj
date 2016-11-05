VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDeptTransferPatientRelease 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   14010
   FillColor       =   &H00800000&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000001&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   14010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   -90
      TabIndex        =   158
      Top             =   10920
      Width           =   24135
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Developed and Maintenanced By :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   510
         TabIndex        =   160
         Top             =   240
         Width           =   3585
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Software Programmer, IT, DNMIH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   4260
         TabIndex        =   159
         Top             =   180
         Width           =   4650
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Caption         =   "Test Info"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      Left            =   0
      TabIndex        =   154
      Top             =   6210
      Width           =   6255
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   585
         Left            =   120
         TabIndex        =   155
         ToolTipText     =   "PRESS DOUBLE CLICK TO VIEW TEST"
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   1032
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   14737632
         BackColorBkg    =   -2147483647
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   2
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox txtDepartmentSerial 
      Height          =   285
      Left            =   10530
      TabIndex        =   112
      Top             =   8970
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtExtraBedFlag 
      Height          =   285
      Left            =   9060
      TabIndex        =   106
      Top             =   9030
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   525
      Left            =   7620
      TabIndex        =   14
      Top             =   9390
      Width           =   1395
   End
   Begin VB.CommandButton NewButton 
      Caption         =   "NEW"
      Height          =   525
      Left            =   9030
      TabIndex        =   103
      Top             =   9390
      Width           =   1395
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "REPORT"
      Height          =   525
      Left            =   10440
      TabIndex        =   102
      Top             =   9390
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   525
      Left            =   11850
      TabIndex        =   101
      Top             =   9390
      Width           =   1395
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Caption         =   "Bed Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   2070
      Left            =   0
      TabIndex        =   76
      Top             =   780
      Width           =   6240
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1605
         Left            =   60
         TabIndex        =   109
         Top             =   210
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   2831
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   14737632
         BackColorBkg    =   -2147483647
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   2
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000001&
      Caption         =   "Operation Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   1860
      Left            =   0
      TabIndex        =   75
      Top             =   4230
      Width           =   6240
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
         Height          =   1515
         Left            =   120
         TabIndex        =   156
         Top             =   210
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   2672
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   14737632
         BackColorBkg    =   -2147483647
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   2
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000001&
      Caption         =   "Extra Bed Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   795
      Left            =   0
      TabIndex        =   74
      Top             =   3435
      Width           =   6240
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   435
         Left            =   60
         TabIndex        =   110
         Top             =   240
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   767
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   14737632
         BackColorBkg    =   -2147483647
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   2
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2310
      Top             =   8790
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
      Caption         =   "Adodc2"
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
      Left            =   3645
      Top             =   8760
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
   Begin VB.Frame Frame4 
      BackColor       =   &H80000001&
      Height          =   855
      Left            =   0
      TabIndex        =   34
      Top             =   -60
      Width           =   12915
      Begin VB.TextBox comDepartmentRelease 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4560
         TabIndex        =   111
         Top             =   495
         Width           =   1965
      End
      Begin VB.TextBox txtNameRelease 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   180
         Width           =   3675
      End
      Begin VB.TextBox txtBedType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9810
         TabIndex        =   104
         Top             =   480
         Width           =   1905
      End
      Begin VB.TextBox txtRegNoShow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1740
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   495
         Width           =   1215
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   9780
         TabIndex        =   108
         Top             =   150
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483647
         ForeColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "DD-MMM-YYYY"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   9510
         TabIndex        =   115
         Top             =   510
         Width           =   105
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   9510
         TabIndex        =   114
         Top             =   180
         Width           =   105
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   4380
         TabIndex        =   113
         Top             =   510
         Width           =   105
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Bed "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7800
         TabIndex        =   105
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1500
         TabIndex        =   63
         Top             =   180
         Width           =   105
      End
      Begin VB.Label lblRegNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration No :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   60
         Top             =   495
         Width           =   1395
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Dept."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2970
         TabIndex        =   37
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   150
         TabIndex        =   36
         Top             =   180
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Admission Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   7800
         TabIndex        =   35
         Top             =   180
         Width           =   1500
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   7890
      Left            =   6870
      TabIndex        =   38
      Top             =   810
      Width           =   6225
      Begin VB.TextBox txtDisc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3750
         TabIndex        =   15
         Text            =   "0"
         Top             =   6960
         Width           =   675
      End
      Begin VB.TextBox txtDueAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   390
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   153
         Top             =   7500
         Width           =   2220
      End
      Begin VB.TextBox TXTaDVANCE2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   6180
         Width           =   960
      End
      Begin VB.TextBox TXTaDVANCE1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   6180
         Width           =   960
      End
      Begin VB.TextBox txtDisc1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   116
         Text            =   "0"
         Top             =   6960
         Width           =   930
      End
      Begin VB.TextBox txtServiceCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   99
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   720
         Width           =   960
      End
      Begin VB.TextBox txtOperationCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   98
         Text            =   "0"
         Top             =   1380
         Width           =   960
      End
      Begin VB.TextBox txtExtraBedCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   97
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1050
         Width           =   960
      End
      Begin VB.TextBox txtMiscelleneousCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   96
         Text            =   "0"
         Top             =   5145
         Width           =   960
      End
      Begin VB.TextBox txtTotal1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   5820
         Width           =   960
      End
      Begin VB.TextBox txtAdmissionCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   94
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   375
         Width           =   960
      End
      Begin VB.TextBox txtBedCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   93
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   45
         Width           =   960
      End
      Begin VB.TextBox txtBabyCareCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   92
         Text            =   "0"
         Top             =   2400
         Width           =   960
      End
      Begin VB.TextBox txtNeunetalCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   91
         Text            =   "0"
         Top             =   2730
         Width           =   960
      End
      Begin VB.TextBox txtExTrafusionCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   90
         Text            =   "0"
         Top             =   3060
         Width           =   960
      End
      Begin VB.TextBox txtBloodSugarCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   89
         Text            =   "0"
         Top             =   3750
         Width           =   960
      End
      Begin VB.TextBox txtPhotoTherapyCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   88
         Text            =   "0"
         Top             =   3405
         Width           =   960
      End
      Begin VB.TextBox txtMedicineCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "0"
         Top             =   5490
         Width           =   960
      End
      Begin VB.TextBox txtDeliveryCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   86
         Text            =   "0"
         Top             =   2055
         Width           =   960
      End
      Begin VB.TextBox txtCCuCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   85
         Text            =   "0"
         Top             =   4455
         Width           =   960
      End
      Begin VB.TextBox txtAnesthesiaCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   84
         Text            =   "0"
         Top             =   1725
         Width           =   960
      End
      Begin VB.TextBox txtNebuliserCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   83
         Text            =   "0"
         Top             =   4815
         Width           =   960
      End
      Begin VB.TextBox txtIncubatorCharge1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "0"
         Top             =   4095
         Width           =   960
      End
      Begin VB.TextBox txtServiceCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   81
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   720
         Width           =   960
      End
      Begin VB.TextBox txtOperationCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         TabIndex        =   0
         Text            =   "0"
         Top             =   1380
         Width           =   960
      End
      Begin VB.TextBox txtExtraBedCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         TabIndex        =   80
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1050
         Width           =   960
      End
      Begin VB.TextBox txtMiscelleneousCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         TabIndex        =   11
         Text            =   "0"
         Top             =   5145
         Width           =   960
      End
      Begin VB.TextBox txtTotal2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   5820
         Width           =   960
      End
      Begin VB.TextBox txtAdmissionCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   78
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   375
         Width           =   960
      End
      Begin VB.TextBox txtBedCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   45
         Width           =   960
      End
      Begin VB.TextBox txtBabyCareCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         TabIndex        =   3
         Text            =   "0"
         Top             =   2400
         Width           =   960
      End
      Begin VB.TextBox txtNeunetalCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         TabIndex        =   4
         Text            =   "0"
         Top             =   2730
         Width           =   960
      End
      Begin VB.TextBox txtExTrafusionCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         TabIndex        =   5
         Text            =   "0"
         Top             =   3060
         Width           =   960
      End
      Begin VB.TextBox txtBloodSugarCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         TabIndex        =   7
         Text            =   "0"
         Top             =   3750
         Width           =   960
      End
      Begin VB.TextBox txtPhotoTherapyCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         TabIndex        =   6
         Text            =   "0"
         Top             =   3405
         Width           =   960
      End
      Begin VB.TextBox txtMedicineCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3750
         TabIndex        =   12
         Text            =   "0"
         Top             =   5490
         Width           =   960
      End
      Begin VB.TextBox txtDeliveryCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   3750
         TabIndex        =   2
         Text            =   "0"
         Top             =   2055
         Width           =   960
      End
      Begin VB.TextBox txtCCuCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         TabIndex        =   9
         Text            =   "0"
         Top             =   4455
         Width           =   960
      End
      Begin VB.TextBox txtAnesthesiaCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         TabIndex        =   1
         Text            =   "0"
         Top             =   1725
         Width           =   960
      End
      Begin VB.TextBox txtNebuliserCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         TabIndex        =   10
         Text            =   "0"
         Top             =   4815
         Width           =   960
      End
      Begin VB.TextBox txtIncubatorCharge2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3750
         TabIndex        =   8
         Text            =   "0"
         Top             =   4095
         Width           =   960
      End
      Begin VB.ComboBox txtstaff 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   270
         TabIndex        =   73
         Top             =   6570
         Width           =   1515
      End
      Begin VB.TextBox txtIncubator 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0"
         Top             =   4095
         Width           =   870
      End
      Begin VB.TextBox txtNebuliser 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "0"
         Top             =   4815
         Width           =   870
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "Check2"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1860
         TabIndex        =   68
         Top             =   6630
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   67
         Top             =   6630
         Width           =   225
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0"
         Top             =   1725
         Width           =   870
      End
      Begin VB.TextBox txtCCU_Charge 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0"
         Top             =   4455
         Width           =   870
      End
      Begin VB.TextBox txtDeliveryCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0"
         Top             =   2055
         Width           =   870
      End
      Begin VB.TextBox txtdisc_percent 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4620
         TabIndex        =   13
         Text            =   "0"
         Top             =   6960
         Width           =   390
      End
      Begin VB.TextBox txtMedicine_charge 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "0"
         Top             =   5490
         Width           =   870
      End
      Begin VB.TextBox txtP_therapyCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0"
         Top             =   3405
         Width           =   870
      End
      Begin VB.TextBox txtBloodTher_charge 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0"
         Top             =   3750
         Width           =   870
      End
      Begin VB.TextBox txtExtTransfusion 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0"
         Top             =   3060
         Width           =   870
      End
      Begin VB.TextBox txtNeunetalCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0"
         Top             =   2730
         Width           =   870
      End
      Begin VB.TextBox txtBabyCareCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0"
         Top             =   2400
         Width           =   870
      End
      Begin VB.TextBox txtTotalBedCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   45
         Width           =   870
      End
      Begin VB.TextBox txtWithAdmissionCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   375
         Width           =   870
      End
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   5820
         Width           =   870
      End
      Begin VB.TextBox txtmisce 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0"
         Top             =   5145
         Width           =   870
      End
      Begin VB.TextBox TxtTotalDisount 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5175
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   6960
         Width           =   870
      End
      Begin VB.TextBox txtNetTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5175
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   6570
         Width           =   870
      End
      Begin VB.TextBox txtExtraBedTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1050
         Width           =   870
      End
      Begin VB.TextBox txtTotalOpr 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         Top             =   1380
         Width           =   870
      End
      Begin VB.TextBox txtServiceCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   720
         Width           =   870
      End
      Begin VB.TextBox txtAdvanceRelease 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   6180
         Width           =   870
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         Height          =   435
         Left            =   30
         Top             =   6900
         Width           =   6045
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   33
         Left            =   3270
         TabIndex        =   150
         Top             =   5205
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   32
         Left            =   3270
         TabIndex        =   149
         Top             =   5520
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   30
         Left            =   3270
         TabIndex        =   146
         Top             =   45
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   29
         Left            =   3270
         TabIndex        =   145
         Top             =   4845
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   28
         Left            =   3270
         TabIndex        =   144
         Top             =   4515
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   27
         Left            =   3270
         TabIndex        =   143
         Top             =   4125
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   26
         Left            =   3270
         TabIndex        =   142
         Top             =   3780
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   25
         Left            =   3270
         TabIndex        =   141
         Top             =   3435
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   24
         Left            =   3270
         TabIndex        =   140
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   23
         Left            =   3270
         TabIndex        =   139
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   22
         Left            =   3270
         TabIndex        =   138
         Top             =   2430
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   21
         Left            =   3270
         TabIndex        =   137
         Top             =   2085
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   20
         Left            =   3270
         TabIndex        =   136
         Top             =   1785
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   3270
         TabIndex        =   135
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   18
         Left            =   3270
         TabIndex        =   134
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   4830
         TabIndex        =   133
         Top             =   6240
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   4830
         TabIndex        =   132
         Top             =   5850
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   4830
         TabIndex        =   131
         Top             =   5520
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   4830
         TabIndex        =   130
         Top             =   5205
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   4830
         TabIndex        =   129
         Top             =   4845
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   4830
         TabIndex        =   128
         Top             =   4515
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   4830
         TabIndex        =   127
         Top             =   4125
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   4830
         TabIndex        =   126
         Top             =   3780
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   4830
         TabIndex        =   125
         Top             =   3435
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   4830
         TabIndex        =   124
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   4830
         TabIndex        =   123
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   4830
         TabIndex        =   122
         Top             =   2430
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   4830
         TabIndex        =   121
         Top             =   2085
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   4830
         TabIndex        =   120
         Top             =   1785
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   4830
         TabIndex        =   119
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   4830
         TabIndex        =   118
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   4800
         TabIndex        =   117
         Top             =   45
         Width           =   240
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   100
         Top             =   5820
         Width           =   480
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Incubator  Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   72
         Top             =   4095
         Width           =   1530
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nebulizer Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   71
         Top             =   4815
         Width           =   1425
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Poor Patient"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   225
         Left            =   2100
         TabIndex        =   70
         Top             =   6630
         Width           =   1095
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Staff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   225
         Left            =   330
         TabIndex        =   69
         Top             =   6630
         Width           =   405
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Anaesthesia Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   66
         Top             =   1725
         Width           =   1740
      End
      Begin VB.Label lab10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CCU Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   65
         Top             =   4455
         Width           =   975
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   64
         Top             =   2055
         Width           =   1350
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   4440
         TabIndex        =   62
         Top             =   6990
         Width           =   210
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   59
         Top             =   5490
         Width           =   1395
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Sugar Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   58
         Top             =   3750
         Width           =   1650
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Photo Therapy Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   57
         Top             =   3405
         Width           =   1890
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ex.Transfusion Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   56
         Top             =   3060
         Width           =   1875
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Neunetal Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   55
         Top             =   2730
         Width           =   1395
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Baby-care Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   54
         Top             =   2400
         Width           =   1515
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DUE AMOUNT "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   2145
         TabIndex        =   53
         Top             =   7590
         Width           =   1140
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   105
         TabIndex        =   52
         Top             =   6990
         Width           =   735
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   51
         Top             =   45
         Width           =   960
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Admission Fee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   50
         Top             =   375
         Width           =   1215
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Miscelleneous Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   49
         Top             =   5145
         Width           =   1830
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NET TOTAL AMT."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   3690
         TabIndex        =   48
         Top             =   6600
         Width           =   1350
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Ext. BedCharge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   47
         Top             =   1050
         Width           =   1320
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Operation Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   46
         Top             =   1380
         Width           =   1485
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   45
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label92 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Advance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   44
         Top             =   6180
         Width           =   735
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(--)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   3270
         TabIndex        =   43
         Top             =   6240
         Width           =   255
      End
   End
   Begin VB.Label OccuranceDateLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OCCURANCE DATE :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   90
      TabIndex        =   162
      Top             =   3030
      Width           =   3045
   End
   Begin VB.Label DischargeTypeLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label34"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1050
      TabIndex        =   161
      Top             =   9540
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "PATIENT RELEASE INFO. ENTRY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   270
      TabIndex        =   151
      Top             =   8790
      Width           =   7545
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   12570
      TabIndex        =   157
      Top             =   870
      Width           =   45
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "staff name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   270
      TabIndex        =   152
      Top             =   8400
      Width           =   1305
   End
   Begin VB.Shape Shape4 
      Height          =   645
      Left            =   7560
      Top             =   9330
      Width           =   5745
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Total"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7080
      TabIndex        =   33
      Top             =   9480
      Visible         =   0   'False
      Width           =   660
   End
End
Attribute VB_Name = "frmDeptTransferPatientRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim discount_check_val As Integer
Public strUid As String
Dim PATIENT_BED_SERIAL_NO As Integer
Dim total_bed_charge As Double
Dim Extra_bed_Flag_Indicator As Integer
Dim UTILITY As New clsUtility
Dim SERVICE_CHARGE_DEPT_SERIAL As Integer

Private Sub Check1_Click()
     
 If Check1.Value = 1 Then
    Label27.Visible = False
    txtstaff.Visible = True
 If Not IsNull(Val(txtstaff)) And IsBedDiscountIrregularityOccured = False Then
           txtdisc = txtTotal.Text
           discount_check_val = 1
           Check2.Value = 0
 End If
 Else
    Check2.Value = 1
    Label27.Visible = True
    txtstaff.Visible = False
End If

End Sub

Private Sub Check2_Click()
  
  If Check2.Value = 1 Then
     txtdisc = 0
     txtdisc_percent = 0
     Check1.Value = 0
     discount_check_val = 0
     Label16.Caption = ""
 Else
    Check1.Value = 1
 End If
   
End Sub
Private Sub cmdExit_Click()
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
       Unload Me
        If LockingFlag = True Then
          frmIrregularPatientEntry.Show 1
        ElseIf LockingFlag = False Then
          frmRelease.Show 1
        End If
       
    End If
End Sub
Private Sub cmdPrint_Click()
'              Dim validation As Variant
'              Adodc1.ConnectionString = strcn.Connection_String
'              Adodc1.RecordSource = "Select user_id, user_name, user_password,user_type,shift_name From Security Where (User_Id = '" & frmMAIN.lbluser_id & "')"
'              Adodc1.Refresh
'              validation = Adodc1.Recordset!user_id
'
'                Dim conn As New ADODB.Connection
'                Dim cmd As New ADODB.Command
'                Dim rs As New ADODB.Recordset
'
'                        Dim Param1 As New ADODB.Parameter
'                    If conn.State = 0 Then
'                        conn.Open strcn.Connection_String
'                     End If
'                    Set cmd.ActiveConnection = conn
'                    cmd.CommandType = adCmdText
'
'                   Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, validation)
'                    cmd.Parameters.Append Param1 'validation
'                    cmd.Properties("PLSQLRSet") = True
'
'                     cmd.CommandText = "{CALL shift_validation(?)}"
'
'
'
'                    Set rs = cmd.Execute
'
'
'                cmd.Properties("PLSQLRSet") = False
'                Set cmd = Nothing
'                 Set rs = Nothing
'                  Set conn = Nothing
'          Adodc2.ConnectionString = strcn.Connection_String
'          Adodc2.RecordSource = "Select * From user_validation"
'          Adodc2.Refresh
'
'
'
'             If Adodc2.Recordset!validation = 0 Then
'                    MsgBox "Your Working Time has been Expired", vbInformation, " IT, DNMIH."
'                    Exit Sub
'             End If
'
'Dim reply As String
'    reply = MsgBox("Are you sure to print?", vbQuestion + vbYesNo, "Print...")
'    If reply = vbYes Then
'        ''Get_Voucher_Number
'        Call save_release_info_rough
'          rptMode = 31
'         '''' Viewer.Show vbModal
'              '''print_release_rough
'
''        Unload Me
''        frmRelease.Show
'        Else
'
'    Unload Me
'     frmRelease.Show
'End If
'Set cmd = Nothing
'Set rs = Nothing
End Sub
Private Sub cmdSave_Click()
If UTILITY.User_Shift_validation(frmMAIN.lbluser_id, frmMAIN.lblUserType.Caption) = False Then
    MsgBox "Mr. " & frmMAIN.lblName & "  Your Shift has been Expired.. " & vbCrLf & " " & vbCrLf & "Please Contact With Administrator", vbInformation, " IT, DNMIH."
    Exit Sub
End If
             
             
 If Val(txtCCuCharge2.Text) > 0 Then
     Select Case UCase(Mid(txtbedType, 1, 3))
            Case "FRE"
                 If Val(txtCCuCharge2.Text) < 500 Then
                     MsgBox "Pls Verify CCU Charge", vbInformation, " IT, DNMIH"
                     txtCCuCharge2.SetFocus
                     Exit Sub
                 End If
                 
            Case "CAB", "PAY"
                If Val(txtCCuCharge2.Text) < 2000 Then
                     MsgBox "Pls Verify CCU Charge", vbInformation, " IT, DNMIH"
                     txtCCuCharge2.SetFocus
                     Exit Sub
                 End If
       End Select
 
 End If
  
            
   '''''''''DEPARTMENT WISE PAYMENT VALIDATION''''''''''''''''''
              
  '   NOTE :ALREADY VALIDATED ON FORM LOAD
      '' OPERATION AND DELIVERY CHARGE FOR DEPARTMENT ARE VALIDATED ON FORM LOAD
  
         


  
  '''''''''End of DEPARTMENT WISE PAYMENT VALIDATION''''''''''''''''''
             
 Dim MSG As String
 Dim reply As String

 MSG = UTILITY.GetPatientCurrentStatusInStringValue(Cur_reg_no, cur_yr_code)
 
 If PatientStatus <> 1 And PatientStatus <> -1 And LockingFlag = True Then '''ABSCONDED/hold /BACKDATED
            PatientStatus = PatientStatusToBe
            reply = MsgBox("Are you sure to lock this Account?", vbQuestion + vbYesNo, "LOCKING...ACCOUNT")
            If reply = 6 Then
                  save_fled_info
                  MsgBox "Account Locked Successfully", vbInformation, "IT Division,DNMIH"
                  Unload Me
                  frmIrregularPatientEntry.Show 1
             End If
       
 ElseIf PatientStatus <> 1 And PatientStatus <> -1 And LockingFlag = False Then  ''''IF PatientStatus=0 ,REGULAR RELEASE
        reply = MsgBox("Are you sure to Release?", vbQuestion + vbYesNo, "Release...")
                      If reply = 6 Then
                            If Check1.Value = 1 And (txtstaff) = "" Then
                                    MsgBox "Please Select a  Staff Id ", vbCritical + vbOKOnly, " IT, DNMIH."
                                    txtstaff.SetFocus
                                    Exit Sub
                            ElseIf Check1.Value = 1 And UTILITY.LOAD_STAFF(txtstaff.Text) = "0" Then
                                    Label16.Caption = "INVALID STAFF ID,... PLEASE VERIFY"
                                    Label16.ForeColor = vbRed
                                    txtstaff.SetFocus
                                    Exit Sub
                            End If  '''END OF CHECK VALUE
               
                            If Check1.Value = 0 And Check2.Value = 0 Then
                               MsgBox "Please select check box", vbInformation, " IT, DNMIH"
                               Exit Sub
                            End If
                            
                            Call save_release_info
                            print_release
                            Unload Me
                            frmRelease.Show 1
                    Else '''IF NOT REPLY=6
                        Unload Me
                        frmRelease.Show 1
                   End If
       
      
      
      Else
      
                   MsgBox MSG, vbInformation, "IT DEPARTMENT,DNMIH"
                   frmRelease.txtRegNoRelease.Text = ""
                   Exit Sub
                   
           
 End If '''' END OF PatientStatus=0
          
 End Sub
Private Sub print_release()
            Dim Connrel As New Connection
            Dim cmdrel As New Command
            Dim RSrel As New Recordset
            Dim Param1 As New Parameter
            Dim Param2 As New Parameter
          If Connrel.State = 0 Then
            Connrel.Open strcn.Connection_String
          End If
            Set cmdrel.ActiveConnection = Connrel
            cmdrel.CommandType = adCmdText
            
            
            Dim Report5   As New CrystalReport5
            
            Set Param1 = cmdrel.CreateParameter("param1", adInteger, adParamInput, 20, Cur_reg_no)
             cmdrel.Parameters.Append Param1 'combo
             
             Set Param2 = cmdrel.CreateParameter("param2", adVarChar, adParamInput, 10, cur_yr_code)
             cmdrel.Parameters.Append Param2 'combo
       
       
            cmdrel.Properties("PLSQLRSet") = True
            cmdrel.CommandText = "{CALL RptPatientRelease(?,?)}"
            Set RSrel = cmdrel.Execute
            cmdrel.Properties("PLSQLRSet") = False
            
            Report5.Database.SetDataSource RSrel

            Report5.PrintOut
            Set RSrel = Nothing
           If Connrel.State = 1 Then
               Connrel.Close
               Set Connrel = Nothing
               Set cmdrel = Nothing
           End If
End Sub
Private Sub save_fled_info()
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
    Dim param16 As New Parameter
    Dim Param17 As New Parameter
    Dim Param18 As New Parameter
    Dim Param19 As New Parameter
    Dim Param20 As New Parameter
    Dim Param21 As New Parameter
    Dim Param22 As New Parameter
    Dim Param23 As New Parameter
    Dim Param24 As New Parameter
    
    
   If Conn.State = 0 Then
     Conn.Open strcn.Connection_String
   End If
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, Cur_reg_no)
    cmd.Parameters.Append Param1 'in_reg_no
    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, cur_yr_code)
    cmd.Parameters.Append Param2 ' FISCAL YR
 
     
    Set Param3 = cmd.CreateParameter("param3", adSingle, adParamInput, 10, txtOperationCharge2)
    cmd.Parameters.Append Param3 'total Operation sum
   
    Set Param4 = cmd.CreateParameter("param4", adSingle, adParamInput, 10, txtBedCharge2)
    cmd.Parameters.Append Param4 'bed_sum
    
    
    Set Param5 = cmd.CreateParameter("param5", adSingle, adParamInput, 10, txtExtraBedCharge2)
    cmd.Parameters.Append Param5 'EXTRA BED CHARGE
   
    Set Param6 = cmd.CreateParameter("param6", adSingle, adParamInput, 10, txtBabyCareCharge2)
    cmd.Parameters.Append Param6 'BABY CARE
       
    Set Param7 = cmd.CreateParameter("param7", adSingle, adParamInput, 10, txtNeunetalCharge2)
    cmd.Parameters.Append Param7 'txtNeunetalCharge2
     
    Set Param8 = cmd.CreateParameter("param8", adSingle, adParamInput, 10, txtExTrafusionCharge2)
    cmd.Parameters.Append Param8 'txtExTrafusionCharge2
    
     Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 10, txtPhotoTherapyCharge2)
    cmd.Parameters.Append Param9 'txtPhotoTherapyCharge2
    
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, txtBloodSugarCharge2)
    cmd.Parameters.Append Param10 'BLOOD SUGAR charge
    
    Set Param11 = cmd.CreateParameter("param11", adDouble, adParamInput, 10, txtMedicineCharge2)
    cmd.Parameters.Append Param11 'txtMedicineCharge2
    
    Set Param12 = cmd.CreateParameter("param12", adDouble, adParamInput, 10, txtMiscelleneousCharge2)
    cmd.Parameters.Append Param12 'txtMiscelleneousCharge2
    
    Set Param13 = cmd.CreateParameter("param13", adDouble, adParamInput, 10, txtDeliveryCharge2)
    cmd.Parameters.Append Param13 'txtDeliveryCharge2
    
    Set Param14 = cmd.CreateParameter("param14", adDouble, adParamInput, 10, txtCCuCharge2)
    cmd.Parameters.Append Param14 'txtCCuCharge
    
    Set Param15 = cmd.CreateParameter("param15", adDouble, adParamInput, 10, txtAnesthesiaCharge2)
    cmd.Parameters.Append Param15 'txtAnesthesiaCharge2
  
     Set param16 = cmd.CreateParameter("param16", adDouble, adParamInput, 10, txtNebuliserCharge2)
    cmd.Parameters.Append param16 ' txtNebuliserCharge2
    
     Set Param17 = cmd.CreateParameter("param17", adDouble, adParamInput, 10, txtIncubatorCharge2)
     cmd.Parameters.Append Param17 ' txtIncubatorCharge2
   
    Set Param18 = cmd.CreateParameter("param18", adDouble, adParamInput, 10, txtNetTotal)
    cmd.Parameters.Append Param18 ' txtNetTotal
    
    Set Param19 = cmd.CreateParameter("param19", adDouble, adParamInput, 10, txtdisc)
    cmd.Parameters.Append Param19 ' txtDisc

   Set Param20 = cmd.CreateParameter("param20", adDate, adParamInput, 15, FLED_DATE)
   cmd.Parameters.Append Param20 ' FLED DATE
   
   Set Param21 = cmd.CreateParameter("param21", adVarChar, adParamInput, 10, frmMAIN.lbluser_id)
   cmd.Parameters.Append Param21 'u_id
    
   Set Param22 = cmd.CreateParameter("param22", adVarChar, adParamInput, 10, frmMAIN.lblBooth)
   cmd.Parameters.Append Param22 'booth_no
   
   Set Param23 = cmd.CreateParameter("param23", adVarChar, adParamInput, 10, txtDepartmentSerial)
   cmd.Parameters.Append Param23 'txtDepartmentSerial
  
  Set Param24 = cmd.CreateParameter("param24", adInteger, adParamInput, 2, PatientStatusToBe)
  cmd.Parameters.Append Param24
        
   cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL save_fled_indoor(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
        
    Set RS = cmd.Execute
    cmd.Properties("PLSQLRSet") = False
  
  If Conn.State = 1 Then
       Set Conn = Nothing
       Set cmd = Nothing
       Set RS = Nothing
  End If
End Sub

Private Sub print_release_rough()
 
''            Dim Connrelr As New ADODB.Connection
''            Dim cmdrelr As New ADODB.Command
''            ''Dim RSrelr As New ADODB.Recordset
''            Dim Param1r As New ADODB.Parameter
''            Dim Report14   As New CrystalReport14
''          If Connrelr.State = 0 Then
''                Connrelr.Open strcn.Connection_String
''          End If
''           Set cmdrelr.ActiveConnection = Connrelr
''           cmdrelr.CommandType = adCmdText
''
''
''
''
''            Set Param1r = cmdrelr.CreateParameter("param1r", adInteger, adParamInput, 20, frmRelease.txtRegNoRelease.Text)
''             cmdrelr.Parameters.Append Param1r 'combo
''
''            cmdrelr.Properties("PLSQLRSet") = True
''            cmdrelr.CommandText = "{CALL RptPatientRelease_rough(?)}"
''            Set rs = cmdrelr.Execute
''            cmdrelr.Properties("PLSQLRSet") = False
''
''            Report14.Database.SetDataSource rs
''            Report14.DiscardSavedData
''            Viewer.CRViewer91.ReportSource = Report14
''            Viewer.CRViewer91.ViewReport
''            Screen.MousePointer = vbDefault
''
''
''
''
''           'Report14.PrintOut
''           ' Set RSrel = Nothing
''           If Connrelr.State = 1 Then
''               Connrelr.Close
''               Set Connrelr = Nothing
''           End If
''           ''Set RSrelr = Nothing
''           Set cmdrelr = Nothing
 End Sub
Private Sub save_release_info()
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
    Dim Param15 As New Parameter
    Dim param16 As New Parameter
    Dim Param17 As New Parameter
    Dim Param18 As New Parameter
    Dim Param19 As New Parameter
    Dim Param20 As New Parameter
    Dim Param21 As New Parameter
    Dim Param22 As New Parameter
    Dim Param23 As New Parameter
    Dim Param24 As New Parameter
    Dim Param25 As New Parameter
    Dim Param26 As New Parameter
    Dim param27 As New Parameter
    Dim param28 As New Parameter
    
    Dim param29 As New Parameter
    Dim Param30 As New Parameter
    Dim Param31 As New Parameter
    Dim Param32 As New Parameter
    Dim Param33 As New Parameter
    Dim Param34 As New Parameter
    Dim Param35 As New Parameter
    Dim Param36 As New Parameter
    Dim Param37 As New Parameter
    Dim Param38 As New Parameter
    Dim Param39 As New Parameter
    Dim param40 As New Parameter
    Dim param41 As New Parameter
    Dim param42 As New Parameter
    Dim param43 As New Parameter
    Dim param44 As New Parameter
    Dim param45 As New Parameter
    Dim param46 As New Parameter
    Dim param47 As New Parameter
    Dim param48 As New Parameter
    Dim param49 As New Parameter
    Dim MODE As Integer
    
   If Conn.State = 0 Then
     Conn.Open strcn.Connection_String
   End If
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    '----------------------------------------------------------------------------------
    If txtDepartmentSerial.Text = "1" Then
       MODE = 0
    Else
      MODE = 1
    End If
    Set Param0 = cmd.CreateParameter("param0", adInteger, adParamInput, 2, MODE)
    cmd.Parameters.Append Param0 'MODE
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, Cur_reg_no)
    cmd.Parameters.Append Param1 'in_reg_no
    
    Set Param2 = cmd.CreateParameter("param2", adDouble, adParamInput, 10, Val(txtTotalBedCharge))
    cmd.Parameters.Append Param2 'BED CHARGE
    
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, Val(txtWithAdmissionCharge))
    cmd.Parameters.Append Param3 'total ADMISSION CHARGE
   
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, Val(txtServiceCharge))
    cmd.Parameters.Append Param4 'SERVICE CHARGE
    

    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 10, Val(txtExtraBedTotal))
    cmd.Parameters.Append Param5 'EXTRA BED CHARGE
    
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 10, Val(txtTotalOpr))
    cmd.Parameters.Append Param6 'TOTAL OPERATION CHARGE
    
    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 10, Val(Text1.Text))
    cmd.Parameters.Append Param7 'ANESTHESIA CHARGE
    

    Set Param8 = cmd.CreateParameter("param8", adDouble, adParamInput, 10, Val(txtDeliveryCharge))
    cmd.Parameters.Append Param8 'DELIVERY CHARGE
    
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 10, Val(txtBabyCareCharge))
    cmd.Parameters.Append Param9 'BABY CARE CHARGE
    
     
    
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, Val(txtNeunetalCharge))
    cmd.Parameters.Append Param10 'NEUNETAL BED CHARGE
    
    
    Set Param11 = cmd.CreateParameter("param11", adDouble, adParamInput, 10, Val(txtExtTransfusion))
    cmd.Parameters.Append Param11 'EXCHANGE TRANSFUSION CHARGE
    
     Set Param12 = cmd.CreateParameter("param12", adDouble, adParamInput, 10, Val(txtP_therapyCharge))
      cmd.Parameters.Append Param12 'PHOTO THERAPY CHARGE
    
     Set Param13 = cmd.CreateParameter("param13", adDouble, adParamInput, 10, Val(txtBloodTher_charge))
      cmd.Parameters.Append Param13 'BLOOD SUGAR CHARGE
    
    Set Param14 = cmd.CreateParameter("param14", adDouble, adParamInput, 10, Val(txtIncubator))
      cmd.Parameters.Append Param14 'INCUBATOR CHARGE
    
    Set Param15 = cmd.CreateParameter("param15", adDouble, adParamInput, 10, Val(txtCCU_Charge))
      cmd.Parameters.Append Param15 'CCU CHARGE
    
    Set param16 = cmd.CreateParameter("param16", adDouble, adParamInput, 10, Val(txtNebuliser))
      cmd.Parameters.Append param16 'NEBULISER CHARGE
    
    Set Param17 = cmd.CreateParameter("param17", adDouble, adParamInput, 10, Val(txtmisce))
    cmd.Parameters.Append Param17 'MISCELLENEOUS CHARGE
        
    Set Param18 = cmd.CreateParameter("param18", adDouble, adParamInput, 10, Val(txtMedicine_charge))
    cmd.Parameters.Append Param18 'MEDICINE CHARGE
    
    Set Param19 = cmd.CreateParameter("param19", adDouble, adParamInput, 10, Val(txtAdvanceRelease))
    cmd.Parameters.Append Param19 ' ADVANCE
   
    Set Param20 = cmd.CreateParameter("param20", adDouble, adParamInput, 10, Val(TxtTotalDisount))
    cmd.Parameters.Append Param20 ' DISCOUNT
   
   Set Param21 = cmd.CreateParameter("param21", adInteger, adParamInput, 5, discount_check_val)
   cmd.Parameters.Append Param21 'STAFF OR POOR PATIENT FLAG
    
   If discount_check_val = 0 Then
     Set Param22 = cmd.CreateParameter("param22", adVarChar, adParamInput, 15, 0)
     cmd.Parameters.Append Param22 ' STAFF ID
   Else
     Set Param22 = cmd.CreateParameter("param22", adVarChar, adParamInput, 15, txtstaff)
      cmd.Parameters.Append Param22 ' STAFF ID

   End If
    
    Set Param23 = cmd.CreateParameter("param23", adDouble, adParamInput, 10, Val(txtTotal))
    cmd.Parameters.Append Param23 ' TOTAL
   
   Set Param24 = cmd.CreateParameter("param24", adDouble, adParamInput, 10, Val(txtNetTotal))
   cmd.Parameters.Append Param24 ' NET TOTAL
   
   Set Param25 = cmd.CreateParameter("param25", adDouble, adParamInput, 10, Val(txtDueAmount))
   cmd.Parameters.Append Param25 'TOTAL DUE
   
   Set Param26 = cmd.CreateParameter("param26", adVarChar, adParamInput, 10, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param26 'USER ID
    
   Set param27 = cmd.CreateParameter("param27", adVarChar, adParamInput, 10, frmMAIN.lblBooth)
   cmd.Parameters.Append param27 'BOOTH NO
    
   Set param28 = cmd.CreateParameter("param28", adVarChar, adParamInput, 10, cur_yr_code)
    cmd.Parameters.Append param28 'YEAR CODE
    
   Set param29 = cmd.CreateParameter("param29", adInteger, adParamInput, 5, Val(txtDepartmentSerial))
   cmd.Parameters.Append param29 ' DEPARTMENT SERIAL
   
   Set Param30 = cmd.CreateParameter("param30", adDouble, adParamInput, 10, Val(txtBedCharge2))
   cmd.Parameters.Append Param30 ' CUR DEPT BED CHARGE
  
   Set Param31 = cmd.CreateParameter("param31", adDouble, adParamInput, 10, Val(txtExtraBedCharge2))
   cmd.Parameters.Append Param31 ' CUR DEPT EXTRA BED CHARGE
 
    Set Param32 = cmd.CreateParameter("param32", adDouble, adParamInput, 10, Val(txtOperationCharge2))
    cmd.Parameters.Append Param32 ' CUR DEPT OPERATION CHARGE
    
    Set Param33 = cmd.CreateParameter("param33", adDouble, adParamInput, 10, Val(txtAnesthesiaCharge2))
    cmd.Parameters.Append Param33 ' CUR ANESTHESIA CHARGE

    Set Param34 = cmd.CreateParameter("param34", adDouble, adParamInput, 10, Val(txtDeliveryCharge2))
    cmd.Parameters.Append Param34 ' CUR DEPT DELIVERY CHARGE
    
    
    Set Param35 = cmd.CreateParameter("param35", adDouble, adParamInput, 10, Val(txtBabyCareCharge2))
    cmd.Parameters.Append Param35 ' CUR DEPT BABY CARE CHARGE
    
    Set Param36 = cmd.CreateParameter("param36", adDouble, adParamInput, 10, Val(txtNeunetalCharge2))
    cmd.Parameters.Append Param36 ' CUR DEPT NEUNETAL CHARGE
    
    Set Param37 = cmd.CreateParameter("param37", adDouble, adParamInput, 10, Val(txtExTrafusionCharge2))
    cmd.Parameters.Append Param37 ' CUR DEPT TRANSFUSION CHARGE
    
    Set Param38 = cmd.CreateParameter("param38", adDouble, adParamInput, 10, Val(txtPhotoTherapyCharge2))
    cmd.Parameters.Append Param38 ' CUR DEPT PHOTOTHERAPY CHARGE
    
    Set Param39 = cmd.CreateParameter("param39", adDouble, adParamInput, 10, Val(txtBloodSugarCharge2))
    cmd.Parameters.Append Param39 ' CUR DEPT BLOOD SUGAR CHARGE


    Set param40 = cmd.CreateParameter("param40", adDouble, adParamInput, 10, Val(txtIncubatorCharge2))
    cmd.Parameters.Append param40 ' CUR DEPT INCUBATOR CHARGE
    
    
    Set param41 = cmd.CreateParameter("param41", adDouble, adParamInput, 10, Val(txtCCuCharge2))
    cmd.Parameters.Append param41 ' CUR DEPT CCU CHARGE


    Set param42 = cmd.CreateParameter("param42", adDouble, adParamInput, 10, Val(txtNebuliserCharge2))
    cmd.Parameters.Append param42 ' CUR DEPT NEBULISER CHARGE

    Set param43 = cmd.CreateParameter("param43", adDouble, adParamInput, 10, Val(txtMiscelleneousCharge2))
    cmd.Parameters.Append param43 ' CUR DEPT MISCELLENEOUS CHARGE


    Set param44 = cmd.CreateParameter("param44", adDouble, adParamInput, 10, Val(txtMedicineCharge2))
    cmd.Parameters.Append param44 ' CUR DEPT MEDICINE CHARGE

    Set param45 = cmd.CreateParameter("param45", adDouble, adParamInput, 10, Val(TXTaDVANCE2))
    cmd.Parameters.Append param45 ' CUR DEPT ADVANCE

    Set param46 = cmd.CreateParameter("param46", adDouble, adParamInput, 10, Val(txtdisc))
    cmd.Parameters.Append param46 ' CUR DEPT DISCOUNT

    Set param47 = cmd.CreateParameter("param47", adDouble, adParamInput, 10, Val(txtTotal2))
    cmd.Parameters.Append param47 ' CUR DEPT TOTAL CHARGE
    
    Set param48 = cmd.CreateParameter("param478", adInteger, adParamInput, 10, IRREGULAR_CASE)
    cmd.Parameters.Append param48 ' FLED FLAG
    
    Set param49 = cmd.CreateParameter("param49", adVarChar, adParamInput, 1, dischargeType)
    cmd.Parameters.Append param49 ' discharge type
 
 
         
   cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL save_calculation_indoor(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    Set RS = cmd.Execute
    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
       Conn.Close
       Set Conn = Nothing
       Set RS = Nothing
       Set cmd = Nothing
  End If
End Sub
Private Sub LOAD_TEST_INFO()
  Dim Conn As New ADODB.Connection
  Dim RS As New ADODB.Recordset
  Dim cmd As New ADODB.Command
  Dim i As Integer
  Dim CHARGE_TOTAL As Integer
  Conn.Open strcn.Connection_String
  With cmd
       .ActiveConnection = Conn
       .CommandType = adCmdText
       .CommandText = "SELECT B.s_name,B.CHARGE  FROM pat_info_sub1_out_door A,Test_Info_SUB B WHERE TO_NUMBER(A.m_code)=TO_NUMBER(B.m_code) AND TO_NUMBER(A.S_code)=TO_NUMBER(B.S_code) AND A.in_reg_no ='" & Trim(Cur_reg_no) & "' AND A.YRCODE ='" & Trim(cur_yr_code) & "'"
  End With
       RS.CursorLocation = adUseClient
       RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
       
       If Not RS.EOF Then
          i = 1
          With MSFlexGrid3
              Do Until RS.EOF
                 .Rows = i + 1
                 .TextMatrix(i, 0) = RS!s_name
                 .ColAlignment(1) = 0
               .TextMatrix(i, 1) = RS!charge
                CHARGE_TOTAL = CHARGE_TOTAL + RS!charge
                i = i + 1
                RS.MoveNext
             Loop
             .Rows = i + 1
             .Row = i
             .Col = 0
             .CellForeColor = vbRed
             .TextMatrix(i, 0) = "TOTAL TEST    = " & i - 1 & "                  TOTAL CHARGE= "
             .Row = i
             .Col = 1
             .CellForeColor = vbRed
             .TextMatrix(i, 1) = CHARGE_TOTAL
             
             .Row = 0
             .Col = 0: .Text = "Test Name  " & "---TOTAL  TEST=" & i - 1 & " ---CHARGE= " & CHARGE_TOTAL
     
    End With
 Else
    MSFlexGrid3.Rows = 1
 End If
 
     Conn.Close
     
     Set Conn = Nothing
     Set RS = Nothing
     Set cmd = Nothing
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyInsert Then
            cmdSave_Click
End If
If KeyCode = vbKeyEscape Then
  cmdExit_Click
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub Form_Load()
   OccuranceDateLabel.Caption = ""
   Label16.Caption = ""
   txtstaff.Visible = False
   discount_check_val = 0
   
   txtRegNoShow = Cur_reg_no
   DischargeTypeLabel.Caption = ""
   
   If UserRole = "ADMIN" Then
      txtBedCharge2.Enabled = True
   Else
      txtBedCharge2.Enabled = False
   End If
   

   Call LOAD_CURRENT_BED
   
   Call Load_Patient_Info
   
   
     
     
   
   Call FORMAT_GRID(1)
   Call FORMAT_GRID(2)
    
   Call Load_Current_Department_Info
   
   
   Call LOAD_BED_HISTORY_FOR_DEPT
   txtBedCharge2 = total_bed_charge
   
'    If Extra_bed_Flag_Indicator = 1 Then
'        Call LOAD_EXTRA_BED_HISTORY
'        txtExtraBedTotal = Val(total_EXTRA_bed_charge)
'     End If

   Call LOAD_ADVANCE
   
   Call LOAD_MAX_ADMISSION_SERVICE_FEE
   
   If PatientStatus = 0 Then
     Call PREVIOUS_DEPT_VALUE(0)
      DischargeTypeLabel.Caption = "Discharge Type is :  " & UTILITY.GetDischargeTypeInString(dischargeType)
      DischargeTypeLabel.ForeColor = getForeColorForDischarge(dischargeType)

   ElseIf PatientStatus = 2 Or PatientStatus = 3 Then
     If Val(txtDepartmentSerial) > 1 Then
        Call PREVIOUS_DEPT_VALUE(1)
     End If
     If LockingFlag = False Then
        DischargeTypeLabel.Caption = "Discharge Type is :  " & UTILITY.GetDischargeTypeInString(dischargeType)
        DischargeTypeLabel.ForeColor = getForeColorForDischarge(dischargeType)
        Call CURRENT_DEPT_VALUE_FOR_FLED_PATIENT
     End If
   End If
   
   LOAD_TOTAL1
   LOAD_TOTAL2
   LOAD_ADDITION
   
   Call txtTotal2_Change
   TxtTotalDisount = Val(txtDisc1) + Val(txtdisc)
   txtDueAmount = Val(txtNetTotal) - Val(TxtTotalDisount)
   '''''''''''''''''''''''''''extra bed''''''''''''''''''''''
         ''''cmd.Properties("iRowsetChange") = False
'''''''''''''''''''''''''operation info'''''''''''''
'       Adodc1.ConnectionString = strcn.Connection_String
'       Adodc1.RecordSource = "select opr_name,opr_charge,annay_charge,service_charge from indoor_pat_operation_info where in_reg_no='" & Trim(Cur_reg_no) & "' AND YRCODE='" & Trim(cur_yr_code) & "'"
'       Adodc1.Refresh
'
   LOAD_ALL_COLLECTION
   
   
     If Len(txtstaff) = 0 Then
       txtstaff.Locked = False
       Call LOAD_STAFF
     Else ''IF STAFF
        Check1.Value = 1
        If IsBedDiscountIrregularityOccured = False Then
              txtdisc_percent = 100
              txtDueAmount = Val(txtNetTotal) - Val(TxtTotalDisount) - Val(txtAdvanceRelease)
              txtstaff.Locked = True
        Else
           MsgBox "THIS STAFF IS NOT ENTITLED TO CABIN", vbInformation, "IT, DNMIH"
        End If
        
     End If
   
   
   
   If comDepartmentRelease.Text = "Gynae-1" Or comDepartmentRelease.Text = "Gynae-2" Or comDepartmentRelease.Text = "Gynae-3" Then
      txtDeliveryCharge2.Locked = False
   Else
      txtDeliveryCharge2.Locked = True
   End If
   
   
   Select Case UCase(comDepartmentRelease)
          Case "SURGERY-1", "SURGERY-2", "SURGERY-3", "GYNAE-1", "GYNAE-2", "GYNAE-3", "ENT", "OPHTH.", "ORTHO."
                txtOperationCharge2.Locked = False
   Case Else
          txtOperationCharge2.Locked = True
   End Select
   
   
   
   
   
    End Sub
Private Function getForeColorForDischarge(dischargeType As String) As String
   Select Case dischargeType
          Case "N"
               getForeColorForDischarge = vbGreen
          Case "R"
               getForeColorForDischarge = vbYellow
          Case "D"
                getForeColorForDischarge = vbRed
          Case "T"
               getForeColorForDischarge = vbBlack
          Case Else
              getForeColorForDischarge = vbGreen
   End Select
          


End Function
Private Sub LOAD_STAFF()
   Adodc1.ConnectionString = strcn.Connection_String
      Adodc1.RecordSource = "select PAYROLL.EMP_INFO.EMP_ID  AS EMP_ID from PAYROLL.EMP_INFO ORDER BY EMP_ID"
      Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        While Adodc1.Recordset.EOF = False
          txtstaff.AddItem Adodc1.Recordset!EMP_ID
          Adodc1.Recordset.MoveNext
        Wend
   End If
End Sub
Private Sub LOAD_TOTAL1()
   txtTotal1 = Val(txtBedCharge1) + Val(txtAdmissionCharge1) + Val(txtServiceCharge1) + Val(txtExtraBedCharge1) + Val(txtOperationCharge1) + Val(txtAnesthesiaCharge1) + Val(txtDeliveryCharge1) + Val(txtBabyCareCharge1) + Val(txtNeunetalCharge1) + Val(txtExTrafusionCharge1) + Val(txtPhotoTherapyCharge1) + Val(txtBloodSugarCharge1) + Val(txtIncubatorCharge1) + Val(txtCCuCharge1) + Val(txtNebuliserCharge1) + Val(txtMiscelleneousCharge1) + Val(txtMedicineCharge1)
End Sub
Private Sub LOAD_TOTAL2()
   txtTotal2 = Val(txtBedCharge2) + Val(txtAdmissionCharge2) + Val(txtServiceCharge2) + Val(txtExtraBedCharge2) + Val(txtOperationCharge2) + Val(txtAnesthesiaCharge2) + Val(txtDeliveryCharge2) + Val(txtBabyCareCharge2) + Val(txtNeunetalCharge2) + Val(txtExTrafusionCharge2) + Val(txtPhotoTherapyCharge2) + Val(txtBloodSugarCharge2) + Val(txtIncubatorCharge2) + Val(txtCCuCharge2) + Val(txtNebuliserCharge2) + Val(txtMiscelleneousCharge2) + Val(txtMedicineCharge2)
End Sub
Private Sub Load_Current_Department_Info()
  Dim Conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim RS As New ADODB.Recordset
  
  If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
  End If
  cmd.ActiveConnection = Conn
  cmd.CommandType = adCmdText
  cmd.CommandText = "SELECT doc_dept,SERIAL_NO FROM INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(Cur_reg_no) & "' AND YRCODE='" & Trim(cur_yr_code) & "' AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(Cur_reg_no) & "' AND YRCODE='" & Trim(cur_yr_code) & "')"
  
  cmd.Properties("iRowsetChange") = True
  cmd.Properties("updatability") = 7
  RS.CursorLocation = adUseClient

  RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
  cmd.Properties("iRowsetChange") = False

  If RS.RecordCount > 0 Then
     comDepartmentRelease = "" & RS!doc_dept
     txtDepartmentSerial = RS!SERIAL_NO
  End If
  
  Conn.Close
  Set Conn = Nothing
  Set RS = Nothing
  Set cmd = Nothing

End Sub
Private Sub LOAD_ALL_COLLECTION()
    txtTotalBedCharge = Val(txtBedCharge1) + Val(txtBedCharge2)
    txtWithAdmissionCharge = Val(txtAdmissionCharge1) + Val(txtAdmissionCharge2)
    txtServiceCharge = Val(txtServiceCharge1) + Val(txtServiceCharge2)
    txtExtraBedTotal = Val(txtExtraBedCharge1) + Val(txtExtraBedCharge2)
    txtTotalOpr = Val(txtOperationCharge1) + Val(txtOperationCharge2)
    Text1 = Val(txtAnesthesiaCharge1) + Val(txtAnesthesiaCharge2)
    txtDeliveryCharge = Val(txtDeliveryCharge1) + Val(txtDeliveryCharge2)
    txtBabyCareCharge = Val(txtBabyCareCharge1) + Val(txtBabyCareCharge2)
    txtNeunetalCharge = Val(txtNeunetalCharge1) + Val(txtNeunetalCharge2)
    txtExtTransfusion = Val(txtExTrafusionCharge1) + Val(txtExTrafusionCharge2)
    txtP_therapyCharge = Val(txtPhotoTherapyCharge1) + Val(txtPhotoTherapyCharge2)
    txtBloodTher_charge = Val(txtBloodSugarCharge1) + Val(txtBloodSugarCharge2)
    txtIncubator = Val(txtIncubatorCharge1) + Val(txtIncubatorCharge2)
    txtCCU_Charge = Val(txtCCuCharge1) + Val(txtCCuCharge2)
    txtNebuliser = Val(txtNebuliserCharge1) + Val(txtNebuliserCharge2)
    txtmisce = Val(txtMiscelleneousCharge1) + Val(txtMiscelleneousCharge2)
    txtMedicine_charge = Val(txtMedicineCharge1) + Val(txtMedicineCharge2)
    txtTotal = Val(txtTotal1) + Val(txtTotal2)
End Sub
Private Sub LOAD_BED_HISTORY_FOR_DEPT()
  Dim Conn As New ADODB.Connection
  Dim RS As New ADODB.Recordset
  Dim cmd As New ADODB.Command
  Dim INNER_CHARGE As Single
  Dim firstRecordAdmissionDate As Date
  Dim endDate As Date
  total_bed_charge = 0
  
  
  Dim i As Integer
  
  If Conn.State = 0 Then
     Conn.ConnectionString = strcn.Connection_String
     Conn.Open
  End If
  cmd.ActiveConnection = Conn
  cmd.CommandType = adCmdText
  cmd.CommandText = "select  bed_no,bed_type,BED_TYPE_NO,bed_charge,admission_charge,admission_date,Extra_bed_flag,ed_dt,DOC_DEPT,DEPT_SERIAL,SERIAL_NO,SYSDATE  From Indoor_pat_bed_info Where " & _
  "in_reg_no ='" & Trim(Cur_reg_no) & "' AND YRCODE='" & Trim(cur_yr_code) & "'  ORDER BY SERIAL_NO ASC"
  cmd.Properties("iRowsetChange") = True
  cmd.Properties("updatability") = 7
  RS.CursorLocation = adUseClient

  RS.Open cmd.CommandText, Conn, adOpenForwardOnly, adLockReadOnly
  cmd.Properties("iRowsetChange") = False

  If Not RS.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until RS.EOF ''' LOOP
                INNER_CHARGE = 0
                If i = 1 Then
                   firstRecordAdmissionDate = Format(RS!admission_date, "DD/MM/YYYY")
                End If
               .Rows = i + 1
            If txtDepartmentSerial <> RS!DEPT_SERIAL Then ''' CREATE FLEXGRID ROW
                  .Row = i
                  .Col = 0
                  .CellBackColor = &HE0E0E0
                  .Col = 1
                  .CellBackColor = &HE0E0E0
                  .Col = 2
                  .CellBackColor = &HE0E0E0
                  .Col = 3
                  .CellBackColor = &HE0E0E0
             End If ''' END OF CREATE FLEXGRID ROW
                 
               .TextMatrix(i, 0) = RS!Bed_type & " (" & RS!bed_TYPE_no & " )  - " & RS!bed_no
               .ColAlignment(1) = 0
               .TextMatrix(i, 1) = Format(RS!admission_date, "DD/MM/YYYY")
               
               ''''for staff bed discount
               
               If Len(txtstaff) > 0 And RS!Bed_type = "Cabin" And UTILITY.IsEntitledToStayInCabin(StaffClass) = False Then
                   IsBedDiscountIrregularityOccured = True
               Else
                   IsBedDiscountIrregularityOccured = False
               End If  '''' end of bed type
               
                Extra_bed_Flag_Indicator = Val(0 & RS!Extra_bed_flag)
                
                     
                     ''''''RELEASE FOR FLED OR HOLD BACKDATED  PATIENT
                      If RS.RecordCount = i Then ''' last or first row
'                            If PatientStatus = 1 Then '''CALLING THE OBJECT FROM FLED ENTRY
'                               INNER_CHARGE = UTILITY.GetBedChargeForRow(RS.RecordCount, i, Format(RS!admission_date, "DD/MM/YYYY"), Format(FLED_DATE, "DD/MM/YYYY"), RS!bed_charge, False)
                          If LockingFlag = True Then
                             PatientStatus = PatientStatusToBe
                             endDate = FLED_DATE
                          ElseIf LockingFlag = False And PatientStatus <> 0 Then
                             endDate = Format(RS!ed_dt, "DD/MM/YYYY")
                             OccuranceDateLabel.Caption = "Occurance Date : " & endDate
                          End If
                          
                          If PatientStatus = 2 Or PatientStatus = 3 Then  '''RELEASE FOR FLED OR HOLD BACKDATED  PATIENT
                              INNER_CHARGE = UTILITY.GetBedChargeForRow(RS.RecordCount, i, Format(RS!admission_date, "DD/MM/YYYY"), endDate, RS!bed_charge, False, False)
                           Else  '''RELEASE FOR NORMAL PATIENT
                                 If (firstRecordAdmissionDate = Format(RS!SYSDATE, "DD/MM/YYYY")) Then
                                    INNER_CHARGE = UTILITY.GetBedChargeForRow(RS.RecordCount, i, Format(RS!admission_date, "DD/MM/YYYY"), Format(RS!SYSDATE, "DD/MM/YYYY"), RS!bed_charge, True, False)
                                 Else
                                    INNER_CHARGE = UTILITY.GetBedChargeForRow(RS.RecordCount, i, Format(RS!admission_date, "DD/MM/YYYY"), Format(RS!SYSDATE, "DD/MM/YYYY"), RS!bed_charge, False, False)
                                 End If
                                 OccuranceDateLabel.Caption = "Today's Date : " & Format(RS!SYSDATE, "dd/mm/yyyy")
                            End If ''' END OF MISSION FLED_PAT
                   Else ''' if not last record
                         INNER_CHARGE = UTILITY.GetBedChargeForRow(RS.RecordCount, i, Format(RS!admission_date, "DD/MM/YYYY"), Format(RS!ed_dt, "DD/MM/YYYY"), RS!bed_charge, False, False)
                   End If
                    
 ''bed charge for last department serial or same serial i.e. same department
                   If txtDepartmentSerial = RS!DEPT_SERIAL Then ''MAX DEPARTMENT I.E .CURRENT DEPT
                         total_bed_charge = total_bed_charge + INNER_CHARGE
                         
                         
                  End If
               
               
               If RS!Bed_type = "Free-Bed" Then
                    .Row = i
                     .Col = 0
                    .CellForeColor = vbRed
                    .Col = 2
                    .CellForeColor = vbRed
                 End If
                .ColAlignment(2) = 0
                .TextMatrix(i, 2) = INNER_CHARGE
                .TextMatrix(i, 3) = RS!doc_dept
                i = i + 1
            RS.MoveNext
        Loop
    End With
 Else
    MSFlexGrid1.Rows = 1
 End If
 
'  MsgBox "CABIN DAYS :" & CABIN_DAYS
'  MsgBox "FREE BED DAYS  :" & FREE_BED_DAYS
'
  Conn.Close
  Set Conn = Nothing
  Set RS = Nothing
  Set cmd = Nothing
  Set UTILITY = Nothing
  End Sub
Private Sub LOAD_CURRENT_BED()
   Dim Conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim RS As New ADODB.Recordset
  
  If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
    End If
        cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        cmd.CommandText = "select BED_TYPE,Bed_type_no,BED_NO,extra_bed_flag,SERIAL_NO  From Indoor_pat_bed_info Where in_reg_no ='" & Trim(Cur_reg_no) & "' AND YRCODE ='" & Trim(cur_yr_code) & "'  AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM Indoor_pat_bed_info WHERE in_reg_no ='" & Trim(Cur_reg_no) & "' AND YRCODE='" & Trim(cur_yr_code) & "')"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        RS.CursorLocation = adUseClient

        RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False
          
       If RS.RecordCount > 0 Then
         txtbedType = "" & RS!Bed_type & "-" & RS!bed_TYPE_no & "-" & RS!bed_no
         txtExtraBedFlag = RS!Extra_bed_flag
         PATIENT_BED_SERIAL_NO = RS!SERIAL_NO
       End If
       
       If Conn.State = 1 Then
        Conn.Close
        Set Conn = Nothing
        Set RS = Nothing
        Set cmd = Nothing
     End If

End Sub
Private Sub Load_Patient_Info()
  Dim Conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim RS As New ADODB.Recordset
  
  If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
  End If
  cmd.ActiveConnection = Conn
  cmd.CommandType = adCmdText
  cmd.CommandText = "select pat_name,admission_date,STAFF_ID  From in_door_pat_info_main Where in_reg_no ='" & Trim(Cur_reg_no) & "' AND YRCODE='" & Trim(cur_yr_code) & "'"
      
  cmd.Properties("iRowsetChange") = True
  cmd.Properties("updatability") = 7
  RS.CursorLocation = adUseClient

  RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
  cmd.Properties("iRowsetChange") = False

  If RS.RecordCount > 0 Then
     txtNameRelease = RS!pat_name
     MaskEdBox1.Text = Format(RS!admission_date, "DD/MM/yyyy")
     txtstaff.Text = "" & RS!STAFF_ID
  End If
  
  Conn.Close
  Set Conn = Nothing
  Set RS = Nothing
  Set cmd = Nothing
     
End Sub
Private Sub FORMAT_GRID(MODE As Integer)
  If MODE = 1 Then
     With MSFlexGrid1
         .Rows = 1
         .Cols = 4
         .Col = 0: .Text = "Bed No"
         .Col = 1: .Text = " Admission Date "
         .Col = 2: .Text = " Charge "
         .Col = 3: .Text = "Department"
         .ColWidth(0) = 2100
         .ColWidth(1) = 1500
         .ColWidth(2) = 1000
         .ColWidth(3) = 1100
         
     End With
  End If
  If MODE = 2 Then
     With MSFlexGrid2
         .Rows = 1
         .Cols = 2
         .Col = 0: .Text = "Start Date"
         .Col = 1: .Text = "End Date"
         
         .ColWidth(0) = 1500
         .ColWidth(1) = 1500
         
     End With
  End If
  If MODE = 3 Then
     With MSFlexGrid3
          .Rows = 1
          .Cols = 2
          
          .Row = 0
          .Col = 0: .Text = "Test Name"
          .Col = 1: .Text = "Charge"
          
          .ColWidth(0) = 4500
          .ColWidth(1) = 900
     End With
  End If
  
End Sub
Private Sub LOAD_ADVANCE()
   Dim Conn As New ADODB.Connection
   Dim RS As New ADODB.Recordset
   Dim cmd As New ADODB.Command
  
   Conn.ConnectionString = strcn.Connection_String
   Conn.Open
   cmd.ActiveConnection = Conn
   cmd.CommandType = adCmdText
   cmd.CommandText = "select  nvl(sum(advance),0) as advance  From advance Where in_reg_no ='" & Trim(Cur_reg_no) & "' AND YRCODE= '" & Trim(cur_yr_code) & "'"
      
   cmd.Properties("iRowsetChange") = True
   cmd.Properties("updatability") = 7
   RS.CursorLocation = adUseClient
   RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
   cmd.Properties("iRowsetChange") = False

   If IsNull(RS!advance) = True Then
       txtAdvanceRelease = 0
   Else
       txtAdvanceRelease = RS!advance
       TXTaDVANCE2 = txtAdvanceRelease
   End If
   Conn.Close
   Set Conn = Nothing
   Set RS = Nothing
   Set cmd = Nothing
   
End Sub
Private Sub LOAD_MAX_ADMISSION_SERVICE_FEE()
  Dim Conn As New ADODB.Connection
    Dim RS As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    Conn.Open strcn.Connection_String
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "select MAX(Admission_charge)ADMISSION_CHARGE ,MAX(service_charge) SERVICE_CHARGE from INDOOR_PAT_BED_INFO where IN_REG_NO='" & Cur_reg_no & "' and YRCODE='" & cur_yr_code & "'"
    
    Set RS = cmd.Execute
    If Not RS.EOF Then
       txtWithAdmissionCharge = RS!ADMISSION_CHARGE
       txtServiceCharge = RS!service_charge
       txtAdmissionCharge2 = txtWithAdmissionCharge
       txtServiceCharge2 = txtServiceCharge
   End If
    
    
    Conn.Close
    Set Conn = Nothing
    Set RS = Nothing
    Set cmd = Nothing
  
End Sub
Private Sub PREVIOUS_DEPT_VALUE(MODE As Integer)
    On Error GoTo ERR_DESC
    Dim Conn As New ADODB.Connection
    Dim RS As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    Conn.Open strcn.Connection_String
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    If MODE = 0 Then
        cmd.CommandText = "select SUM(bed_sum) bed_sum ,SUM(operation_sum) operation_sum," & _
        "SUM(miscelleneous_charge) miscelleneous_charge ,SUM(extra_bed_charge) extra_bed_charge ,SUM(anesthesia_charge) anesthesia_charge," & _
        "SUM(delivery_charge) delivery_charge ,SUM(baby_care_charge) baby_care_charge ,SUM(neunetal_bed_charge) neunetal_bed_charge," & _
        "SUM(Exchange_transfusion_charge) Exchange_transfusion_charge ,SUM(photo_therapy_charge) photo_therapy_charge ,SUM(Blood_sugar_charge) Blood_sugar_charge," & _
        "SUM(medicine_charge) medicine_charge ,SUM(cardiology_charge) cardiology_charge ,SUM(nebuliser_charge) nebuliser_charge,SUM(incubator_charge) incubator_charge," & _
        "SUM(discount) discount from INDOOR_PAT_DEPT_INFO where IN_REG_NO='" & Cur_reg_no & "' and YRCODE='" & cur_yr_code & "'"
        
    ElseIf MODE = 1 Then
      
        cmd.CommandText = "select SUM(bed_sum) bed_sum ,SUM(operation_sum) operation_sum," & _
        "SUM(miscelleneous_charge) miscelleneous_charge ,SUM(extra_bed_charge) extra_bed_charge ,SUM(anesthesia_charge) anesthesia_charge," & _
        "SUM(delivery_charge) delivery_charge ,SUM(baby_care_charge) baby_care_charge ,SUM(neunetal_bed_charge) neunetal_bed_charge," & _
        "SUM(Exchange_transfusion_charge) Exchange_transfusion_charge ,SUM(photo_therapy_charge) photo_therapy_charge ,SUM(Blood_sugar_charge) Blood_sugar_charge," & _
        "SUM(medicine_charge) medicine_charge ,SUM(cardiology_charge) cardiology_charge ,SUM(nebuliser_charge) nebuliser_charge,SUM(incubator_charge) incubator_charge," & _
        "SUM(discount) discount from INDOOR_PAT_DEPT_INFO where IN_REG_NO='" & Cur_reg_no & "' and YRCODE='" & cur_yr_code & "' AND SERIAL_NO <> '" & Val(txtDepartmentSerial) & "' "
        
    End If
    
    Set RS = cmd.Execute
    If Not RS.EOF Then
       txtBedCharge1 = RS!bed_sum
       txtOperationCharge1 = RS!operation_sum
       txtMiscelleneousCharge1 = RS!miscelleneous_charge
       txtExtraBedCharge1 = RS!extra_bed_charge
       txtAnesthesiaCharge1 = RS!anesthesia_charge
       txtDeliveryCharge1 = RS!delivery_charge
       txtBabyCareCharge1 = RS!baby_care_charge
       txtNeunetalCharge1 = RS!neunetal_bed_charge
       txtExTrafusionCharge1 = RS!Exchange_transfusion_charge
       txtPhotoTherapyCharge1 = RS!photo_therapy_charge
       txtBloodSugarCharge1 = RS!blood_sugar_charge
       txtMedicineCharge1 = RS!medicine_charge
       txtCCuCharge1 = RS!cardiology_charge
       txtNebuliserCharge1 = RS!nebuliser_charge
       txtIncubatorCharge1 = RS!incubator_charge
       txtDisc1 = RS!DISCOUNT
    End If
    
    
    Conn.Close
    Set Conn = Nothing
    Set RS = Nothing
    Set cmd = Nothing
    Exit Sub
    
ERR_DESC:
      MsgBox Err.Description, vbCritical, "IT DIVISION,DNMIH"
End Sub
Private Sub CURRENT_DEPT_VALUE_FOR_FLED_PATIENT()
    On Error GoTo ERR_DESC
    Dim Conn As New ADODB.Connection
    Dim RS As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    Conn.Open strcn.Connection_String
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "select SUM(bed_sum) bed_sum ,SUM(operation_sum) operation_sum," & _
    "SUM(miscelleneous_charge) miscelleneous_charge ,SUM(extra_bed_charge) extra_bed_charge ,SUM(anesthesia_charge) anesthesia_charge," & _
    "SUM(delivery_charge) delivery_charge ,SUM(baby_care_charge) baby_care_charge ,SUM(neunetal_bed_charge) neunetal_bed_charge," & _
    "SUM(Exchange_transfusion_charge) Exchange_transfusion_charge ,SUM(photo_therapy_charge) photo_therapy_charge ,SUM(Blood_sugar_charge) Blood_sugar_charge," & _
    "SUM(medicine_charge) medicine_charge ,SUM(cardiology_charge) cardiology_charge ,SUM(nebuliser_charge) nebuliser_charge,SUM(incubator_charge) incubator_charge," & _
    "SUM(discount) discount from INDOOR_PAT_DEPT_INFO where IN_REG_NO='" & Cur_reg_no & "' and YRCODE='" & cur_yr_code & "' AND SERIAL_NO='" & txtDepartmentSerial & "'"
        
    Set RS = cmd.Execute
    If Not RS.EOF Then
       txtBedCharge2 = RS!bed_sum
       txtOperationCharge2 = RS!operation_sum
       txtMiscelleneousCharge2 = RS!miscelleneous_charge
       txtExtraBedCharge2 = RS!extra_bed_charge
       txtAnesthesiaCharge2 = RS!anesthesia_charge
       txtDeliveryCharge2 = RS!delivery_charge
       txtBabyCareCharge2 = RS!baby_care_charge
       txtNeunetalCharge2 = RS!neunetal_bed_charge
       txtExTrafusionCharge2 = RS!Exchange_transfusion_charge
       txtPhotoTherapyCharge2 = RS!photo_therapy_charge
       txtBloodSugarCharge2 = RS!blood_sugar_charge
       txtMedicineCharge2 = RS!medicine_charge
       txtCCuCharge2 = RS!cardiology_charge
       txtNebuliserCharge2 = RS!nebuliser_charge
       txtIncubatorCharge2 = RS!incubator_charge
       txtdisc = "" & RS!DISCOUNT
    End If
    
    
    Conn.Close
    Set Conn = Nothing
    Set RS = Nothing
    Set cmd = Nothing
    Exit Sub
    
ERR_DESC:
      MsgBox Err.Description, vbCritical, "IT DIVISION,DNMIH"
End Sub

Private Sub MSFlexGrid3_DblClick()
   MSFlexGrid3.clear
   FORMAT_GRID (3)
   LOAD_TEST_INFO

End Sub
Private Sub NewButton_Click()
   Unload Me
End Sub
Private Sub Text1_Change()
 If Not IsNumeric(Text1) Then
        Text1 = ""
 Else
      txtTotal = Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(Text1) + Val(txtDeliveryCharge) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator) + Val(txtCCU_Charge) + Val(txtNebuliser) + Val(txtmisce) + Val(txtMedicine_charge)
 End If
 
End Sub

Private Sub Text1_GotFocus()
        Text1.BackColor = &H96E4B1
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1)

End Sub

Private Sub Text1_LostFocus()
       If Text1 = Empty Then
       Text1 = 0
    End If
 Text1.BackColor = &H80000018
 
End Sub
Private Sub LOAD_ADDITION()
   txtTotalBedCharge = Val(txtBedCharge1) + Val(txtBedCharge2)
   txtTotalOpr = Val(txtOperationCharge1) + Val(txtOperationCharge2)
End Sub

Private Sub txtAnesthesiaCharge2_Change()
If Not IsNumeric(txtAnesthesiaCharge2) Then
     txtAnesthesiaCharge2 = ""
     Text1 = Val(txtAnesthesiaCharge1)
Else
  Text1 = Val(txtAnesthesiaCharge1) + Val(txtAnesthesiaCharge2)
End If
LOAD_TOTAL2
End Sub

Private Sub txtAnesthesiaCharge2_GotFocus()
  Label10.ForeColor = vbCyan
  Label10.FontSize = 9
  txtAnesthesiaCharge2.BackColor = &H80000018
  txtAnesthesiaCharge2.SelStart = 0
  txtAnesthesiaCharge2.SelLength = Len(txtAnesthesiaCharge2)
End Sub

Private Sub txtAnesthesiaCharge2_LostFocus()
  Label10.ForeColor = vbWhite
  txtAnesthesiaCharge2.BackColor = vbWhite
  Label10.FontSize = 8
  If Len(txtAnesthesiaCharge2) = 0 Then
     txtAnesthesiaCharge2 = 0
  End If
End Sub

Private Sub txtBabyCareCharge_Change()
  If Not IsNumeric(txtBabyCareCharge) Then
            txtBabyCareCharge = ""
  Else
       txtTotal = Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(Text1) + Val(txtDeliveryCharge) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator) + Val(txtCCU_Charge) + Val(txtNebuliser) + Val(txtmisce) + Val(txtMedicine_charge)
   End If
  
End Sub
Private Sub txtBabyCareCharge_GotFocus()
  If Val(txtBabyCareCharge) = 0 Or Val(txtBabyCareCharge) > 0 Then
        txtBabyCareCharge.ForeColor = vbBlack
        txtBabyCareCharge.Locked = False
         txtBabyCareCharge.BackColor = &H96E4B1
         txtBabyCareCharge.SelStart = 0
         txtBabyCareCharge.SelLength = Len(txtBabyCareCharge)

  End If
End Sub

Private Sub txtBabyCareCharge_LostFocus()
   If txtBabyCareCharge = Empty Then
             txtBabyCareCharge = 0
   End If
   txtBabyCareCharge.BackColor = &H80000018
End Sub

Private Sub txtBabyCareCharge2_Change()
 If Not IsNumeric(txtBabyCareCharge2) Then
    txtBabyCareCharge2 = ""
    txtBabyCareCharge = Val(txtBabyCareCharge1)
 Else
    txtBabyCareCharge = Val(txtBabyCareCharge1) + Val(txtBabyCareCharge2)
 End If
 LOAD_TOTAL2
End Sub

Private Sub txtBabyCareCharge2_GotFocus()
  Label18.ForeColor = vbCyan
  Label18.FontSize = 9
  txtBabyCareCharge2.SelStart = 0
  txtBabyCareCharge2.SelLength = Len(txtBabyCareCharge2)
  
End Sub

Private Sub txtBabyCareCharge2_LostFocus()
  Label18.ForeColor = vbWhite
  Label18.FontSize = 8
  If Len(txtBabyCareCharge2) = 0 Then
    txtBabyCareCharge2 = 0
  End If
End Sub

Private Sub txtBedCharge2_Change()
  If Not IsNumeric(txtBedCharge2) Then
     txtBedCharge2 = ""
     txtTotalBedCharge = Val(txtBedCharge1)
   Else
      txtTotalBedCharge = Val(txtBedCharge2) + Val(txtBedCharge1)
   End If
   LOAD_TOTAL2
End Sub

Private Sub txtBedCharge2_LostFocus()
   If Len(txtBedCharge2) = 0 Then
      txtBedCharge2 = 0
   End If
End Sub

Private Sub txtBloodSugarCharge2_Change()
  If Not IsNumeric(txtBloodSugarCharge2) Then
    txtBloodSugarCharge2 = ""
    txtBloodTher_charge = Val(txtBloodSugarCharge1)
 Else
    txtBloodTher_charge = Val(txtBloodSugarCharge1) + Val(txtBloodSugarCharge2)
 End If
 LOAD_TOTAL2
End Sub

Private Sub txtBloodSugarCharge2_GotFocus()
  Label22.ForeColor = vbCyan
  Label22.FontSize = 9
  txtBloodSugarCharge2.SelStart = 0
  txtBloodSugarCharge2.SelLength = Len(txtBloodSugarCharge2)
End Sub

Private Sub txtBloodSugarCharge2_LostFocus()
  Label22.ForeColor = vbWhite
  Label22.FontSize = 8
  If Len(txtBloodSugarCharge2) = 0 Then
     txtBloodSugarCharge2 = 0
  End If
End Sub

Private Sub txtBloodTher_charge_Change()
If Not IsNumeric(txtBloodTher_charge) Then
       txtBloodTher_charge = ""
 Else
     txtTotal = Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(Text1) + Val(txtDeliveryCharge) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator) + Val(txtCCU_Charge) + Val(txtNebuliser) + Val(txtmisce) + Val(txtMedicine_charge)
End If

 
End Sub

Private Sub txtBloodTher_charge_GotFocus()
          txtBloodTher_charge.BackColor = &H96E4B1
          txtBloodTher_charge.SelStart = 0
          txtBloodTher_charge.SelLength = Len(txtBloodTher_charge)

End Sub
Private Sub txtBloodTher_charge_LostFocus()
    If txtBloodTher_charge = Empty Then
       txtBloodTher_charge = 0
    End If
    
     txtBloodTher_charge.BackColor = &H80000018
End Sub
Private Sub txtCCU_Charge_Change()
If Not IsNumeric(txtCCU_Charge) Then
       txtCCU_Charge = ""
 Else
      txtTotal = Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(Text1) + Val(txtDeliveryCharge) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator) + Val(txtCCU_Charge) + Val(txtNebuliser) + Val(txtmisce) + Val(txtMedicine_charge)
 End If
 
End Sub
Private Sub txtCCU_Charge_GotFocus()
     txtCCU_Charge.BackColor = &H96E4B1
     txtCCU_Charge.SelStart = 0
     txtCCU_Charge.SelLength = Len(txtCCU_Charge)

End Sub
Private Sub txtCCU_Charge_LostFocus()
    If txtCCU_Charge = Empty Then
       txtCCU_Charge = 0
    End If
    txtCCU_Charge.BackColor = &H80000018
End Sub

Private Sub txtCCuCharge2_Change()
 If Not IsNumeric(txtCCuCharge2) Then
    txtCCuCharge2 = ""
    txtCCU_Charge = Val(txtCCuCharge1)
 Else
    txtCCU_Charge = Val(txtCCuCharge1) + Val(txtCCuCharge2)
 End If
 LOAD_TOTAL2
End Sub

Private Sub txtCCuCharge2_GotFocus()
  lab10.ForeColor = vbCyan
  lab10.FontSize = 9
  txtCCuCharge2.SelStart = 0
  txtCCuCharge2.SelLength = Len(txtCCuCharge2)
End Sub

Private Sub txtCCuCharge2_LostFocus()
  lab10.ForeColor = vbWhite
  lab10.FontSize = 8
  If Len(txtCCuCharge2) = 0 Then
     txtCCuCharge2 = 0
  End If
End Sub

Private Sub txtDeliveryCharge_Change()
If Not IsNumeric(txtDeliveryCharge) Then
   txtDeliveryCharge = ""
Else
      txtTotal = Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(Text1) + Val(txtDeliveryCharge) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator) + Val(txtCCU_Charge) + Val(txtNebuliser) + Val(txtmisce) + Val(txtMedicine_charge)
End If

End Sub
Private Sub txtDeliveryCharge_GotFocus()
  txtDeliveryCharge.BackColor = &H96E4B1
  txtDeliveryCharge.SelStart = 0
  txtDeliveryCharge.SelLength = Len(txtDeliveryCharge)

End Sub
Private Sub txtDeliveryCharge_LostFocus()
If (txtDeliveryCharge) = Empty Then
      txtDeliveryCharge = 0
      txtDeliveryCharge.BackColor = &H80000018
End If
    txtDeliveryCharge.BackColor = &H80000018
End Sub

Private Sub txtDeliveryCharge2_Change()
If Not IsNumeric(txtDeliveryCharge2) Then
       txtDeliveryCharge2 = ""
       txtDeliveryCharge = Val(txtDeliveryCharge1)
Else
  txtDeliveryCharge = Val(txtDeliveryCharge1) + Val(txtDeliveryCharge2)
End If
LOAD_TOTAL2
End Sub
Private Sub txtDeliveryCharge2_GotFocus()
  Label3.ForeColor = vbCyan
  Label3.FontSize = 9
  txtDeliveryCharge2.BackColor = &H80000018
  txtDeliveryCharge2.SelStart = 0
  txtDeliveryCharge2.SelLength = Len(txtDeliveryCharge2)
End Sub

Private Sub txtDeliveryCharge2_LostFocus()
  Label3.ForeColor = vbWhite
  Label3.FontSize = 8
  txtDeliveryCharge2.BackColor = vbWhite
  
  If Len(txtDeliveryCharge2) = 0 Then
    txtDeliveryCharge2 = 0
  End If
End Sub

Private Sub txtDisc_Change()
  If Not IsNumeric(txtdisc.Text) Then
     txtdisc = 0
  Else
     TxtTotalDisount = Val(txtDisc1) + Val(txtdisc)
     txtDueAmount = Val(txtNetTotal) - Val(TxtTotalDisount)
  End If
End Sub
Private Sub txtdisc_GotFocus()
txtdisc_percent = 0
txtdisc.BackColor = &H96E4B1
End Sub

Private Sub txtDisc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     cmdSAVE.SetFocus
  End If
End Sub

Private Sub txtDisc_LostFocus()
        txtdisc.BackColor = &HF7E3B3
End Sub
Private Sub txtdisc_percent_Change()
If Not IsNumeric(txtdisc_percent.Text) Then
        txtdisc_percent = 0
Else
    If txtdisc_percent = 0 Or txtdisc_percent = "" Then
        txtdisc = 0
 End If
 
 If txtdisc_percent <> 0 And IsNull(txtdisc_percent) = False Then '''IF AMOUNT IS MINUS
        txtdisc = Abs(Val(txtNetTotal) * Val(txtdisc_percent)) / 100
ElseIf txtdisc_percent <> 0 And IsNull(txtdisc_percent) = False Then
       txtdisc = Abs(Val(txtNetTotal) * Val(txtdisc_percent)) / 100
End If


End If
End Sub
Private Sub txtdisc_percent_GotFocus()
txtdisc = 0
txtdisc.BackColor = &H96E4B1
End Sub
Private Sub txtdisc_percent_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If Val(txtdisc_percent) = 0 Then
    cmdSAVE.SetFocus
   End If
  End If
End Sub

Private Sub txtdisc_percent_LostFocus()
txtdisc_percent.BackColor = &HF7E3B3
End Sub

Private Sub txtExtraBedCharge2_Change()
If Not IsNumeric(txtExtraBedCharge2) Then
       txtExtraBedCharge2 = ""
       txtExtraBedTotal = Val(txtExtraBedCharge1)
Else
  txtExtraBedTotal = Val(txtExtraBedCharge1) + Val(txtExtraBedCharge2)
End If

  LOAD_TOTAL2
End Sub

Private Sub txtExtraBedTotal_Change()
  txtTotalOpr_Change
End Sub
Private Sub txtExtraBedTotal_GotFocus()
  txtExtraBedTotal.SelStart = 0
  txtExtraBedTotal.SelLength = Len(txtExtraBedTotal)

End Sub

Private Sub txtExtraBedTotal_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     txtTotalOpr.SetFocus
  End If
End Sub

Private Sub txtExtraBedTotal_LostFocus()
      txtExtraBedTotal.BackColor = &H80000018
End Sub

Private Sub txtExTrafusionCharge2_Change()
 If Not IsNumeric(txtExTrafusionCharge2) Then
    txtExTrafusionCharge2 = ""
    txtExtTransfusion = Val(txtExTrafusionCharge1)
 Else
    txtExtTransfusion = Val(txtExTrafusionCharge1) + Val(txtExTrafusionCharge2)
 End If
 LOAD_TOTAL2
End Sub

Private Sub txtExTrafusionCharge2_GotFocus()
  Label20.ForeColor = vbCyan
  Label20.FontSize = 9
  txtExTrafusionCharge2.SelStart = 0
  txtExTrafusionCharge2.SelLength = Len(txtExTrafusionCharge2)
End Sub

Private Sub txtExTrafusionCharge2_LostFocus()
   Label20.ForeColor = vbWhite
   Label20.FontSize = 8
   If Len(txtExTrafusionCharge2) = 0 Then
      txtExTrafusionCharge2 = 0
   End If
End Sub

Private Sub txtExtTransfusion_Change()
If Not IsNumeric(txtExtTransfusion) Then
      txtExtTransfusion = ""
 Else
    txtTotal = Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(Text1) + Val(txtDeliveryCharge) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator) + Val(txtCCU_Charge) + Val(txtNebuliser) + Val(txtmisce) + Val(txtMedicine_charge)
End If

End Sub

Private Sub txtExtTransfusion_GotFocus()
        txtExtTransfusion.BackColor = &H96E4B1
        txtExtTransfusion.SelStart = 0
        txtExtTransfusion.SelLength = Len(txtExtTransfusion)

End Sub

Private Sub txtExtTransfusion_LostFocus()
  If txtExtTransfusion = Empty Then
     txtExtTransfusion = 0
   End If
   
     txtExtTransfusion.BackColor = &H80000018
End Sub



Private Sub txtIncubator_Change()
  If Not IsNumeric(txtIncubator) Then
         txtIncubator = ""
   Else
         txtTotal = Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(Text1) + Val(txtDeliveryCharge) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator) + Val(txtCCU_Charge) + Val(txtNebuliser) + Val(txtmisce) + Val(txtMedicine_charge)
  End If
 
End Sub

Private Sub txtIncubator_GotFocus()
     txtIncubator.SelStart = 0
     txtIncubator.SelLength = Len(txtIncubator)

End Sub

Private Sub txtIncubator_LostFocus()
   If txtIncubator = "" Then
     txtIncubator = 0
   End If
End Sub

Private Sub txtIncubatorCharge2_Change()
 If Not IsNumeric(txtIncubatorCharge2) Then
    txtIncubatorCharge2 = ""
    txtIncubator = Val(txtIncubatorCharge1)
 Else
  txtIncubator = Val(txtIncubatorCharge1) + Val(txtIncubatorCharge2)
 End If
 LOAD_TOTAL2
End Sub
Private Sub txtIncubatorCharge2_GotFocus()
  Label30.ForeColor = vbCyan
  Label30.FontSize = 9
  txtIncubatorCharge2.SelStart = 0
  txtIncubatorCharge2.SelLength = Len(txtIncubatorCharge2)
End Sub

Private Sub txtIncubatorCharge2_LostFocus()
  Label30.ForeColor = vbWhite
  Label30.FontSize = 8
  If Len(txtIncubatorCharge2) = 0 Then
     txtIncubatorCharge2 = 0
  End If
End Sub

Private Sub txtMedicine_charge_Change()
If Not IsNumeric(txtMedicine_charge.Text) Then
    txtMedicine_charge = 0
Else
   txtTotal = Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(Text1) + Val(txtDeliveryCharge) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator) + Val(txtCCU_Charge) + Val(txtNebuliser) + Val(txtmisce) + Val(txtMedicine_charge)
End If
   
End Sub

Private Sub txtMedicine_charge_GotFocus()
    txtMedicine_charge.BackColor = &H96E4B1
    txtMedicine_charge.SelStart = 0
    txtMedicine_charge.SelLength = Len(txtMedicine_charge)

End Sub



Private Sub txtMedicine_charge_LostFocus()
txtMedicine_charge.BackColor = &HF7E3B3
End Sub

Private Sub txtMedicineCharge2_Change()
  If Not IsNumeric(txtMedicineCharge2) Then
     txtMedicineCharge2 = ""
     txtMedicine_charge = Val(txtMedicineCharge1)
  Else
     txtMedicine_charge = Val(txtMedicineCharge1) + Val(txtMedicineCharge2)
  End If
  LOAD_TOTAL2
End Sub

Private Sub txtMedicineCharge2_GotFocus()
  Label23.ForeColor = vbCyan
  Label23.FontSize = 9
  txtMedicineCharge2.SelStart = 0
  txtMedicineCharge2.SelLength = Len(txtMedicineCharge2)
End Sub

Private Sub txtMedicineCharge2_LostFocus()
  Label23.ForeColor = vbWhite
  Label23.FontSize = 8
  
  If Len(txtMedicineCharge2) = 0 Then
     txtMedicineCharge2 = 0
  End If
End Sub

Private Sub txtmisce_Change()
If Not IsNumeric(txtmisce.Text) Then
    txtmisce = 0
Else
    txtTotal = Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(Text1) + Val(txtDeliveryCharge) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator) + Val(txtCCU_Charge) + Val(txtNebuliser) + Val(txtmisce) + Val(txtMedicine_charge)
End If
End Sub

Private Sub txtmisce_GotFocus()
  txtmisce.BackColor = &H96E4B1
  txtmisce.SelStart = 0
  txtmisce.SelLength = Len(txtmisce)

End Sub

Private Sub txtmisce_LostFocus()
txtmisce.BackColor = &HF7E3B3
End Sub



Private Sub txtMiscelleneousCharge2_Change()
  If Not IsNumeric(txtMiscelleneousCharge2) Then
     txtMiscelleneousCharge2 = ""
     txtmisce = Val(txtMiscelleneousCharge1)
  Else
     txtmisce = Val(txtMiscelleneousCharge1) + Val(txtMiscelleneousCharge2)
  End If
  LOAD_TOTAL2
End Sub

Private Sub txtMiscelleneousCharge2_GotFocus()
  Label8.ForeColor = vbCyan
  Label8.FontSize = 9
  txtMiscelleneousCharge2.SelStart = 0
  txtMiscelleneousCharge2.SelLength = Len(txtMiscelleneousCharge2)
  
End Sub

Private Sub txtMiscelleneousCharge2_LostFocus()
  Label8.ForeColor = vbWhite
  Label8.FontSize = 8
  If Len(txtMiscelleneousCharge2) = 0 Then
     txtMiscelleneousCharge2 = 0
  End If
End Sub

Private Sub txtNebuliser_Change()
  If Not IsNumeric(txtNebuliser) Then
          txtNebuliser = ""
  Else
    txtTotal = Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(Text1) + Val(txtDeliveryCharge) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator) + Val(txtCCU_Charge) + Val(txtNebuliser) + Val(txtmisce) + Val(txtMedicine_charge)
  End If
  
End Sub

Private Sub txtNebuliser_GotFocus()
    txtNebuliser.SelStart = 0
    txtNebuliser.SelLength = Len(txtNebuliser)

End Sub

Private Sub txtNebuliser_LostFocus()
  If txtNebuliser = "" Then
    txtNebuliser = 0
  End If
End Sub

Private Sub txtNebuliserCharge2_Change()
  If Not IsNumeric(txtNebuliserCharge2) Then
     txtNebuliserCharge2 = ""
     txtNebuliser = Val(txtNebuliserCharge1)
  Else
    txtNebuliser = Val(txtNebuliserCharge1) + Val(txtNebuliserCharge2)
  End If
  LOAD_TOTAL2
End Sub

Private Sub txtNebuliserCharge2_GotFocus()
  Label29.ForeColor = vbCyan
  Label29.FontSize = 9
  txtNebuliserCharge2.SelStart = 0
  txtNebuliserCharge2.SelLength = Len(txtNebuliserCharge2)
End Sub

Private Sub txtNebuliserCharge2_LostFocus()
  Label29.ForeColor = vbWhite
  Label29.FontSize = 8
  If Len(txtNebuliserCharge2) = 0 Then
     txtNebuliserCharge2 = 0
  End If
  
End Sub

Private Sub txtNetTotal_Change()
 If Len(txtstaff) > 0 And IsBedDiscountIrregularityOccured = False Then
   txtDueAmount = Val(txtNetTotal) - Val(TxtTotalDisount) - Val(txtAdvanceRelease)
 Else
   txtDueAmount = Val(txtNetTotal) - Val(TxtTotalDisount)
 End If
End Sub

Private Sub txtNeunetalCharge_Change()

If Not IsNumeric(txtNeunetalCharge) Then
      txtNeunetalCharge = ""
 Else
    txtTotal = Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(Text1) + Val(txtDeliveryCharge) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator) + Val(txtCCU_Charge) + Val(txtNebuliser) + Val(txtmisce) + Val(txtMedicine_charge)
End If

 
End Sub

Private Sub txtNeunetalCharge_GotFocus()
      txtNeunetalCharge.BackColor = &H96E4B1
      txtNeunetalCharge.SelStart = 0
      txtNeunetalCharge.SelLength = Len(txtNeunetalCharge)

End Sub

Private Sub txtNeunetalCharge_LostFocus()
 If txtNeunetalCharge = Empty Then
    txtNeunetalCharge = 0
 End If
   txtNeunetalCharge.BackColor = &H80000018
End Sub

Private Sub txtNeunetalCharge2_Change()
 If Not IsNumeric(txtNeunetalCharge2) Then
    txtNeunetalCharge2 = ""
    txtNeunetalCharge = Val(txtNeunetalCharge1)
 Else
    txtNeunetalCharge = Val(txtNeunetalCharge1) + Val(txtNeunetalCharge2)
 End If
 LOAD_TOTAL2
End Sub

Private Sub txtNeunetalCharge2_GotFocus()
  Label19.ForeColor = vbCyan
  Label19.FontSize = 9
  txtNeunetalCharge2.SelStart = 0
  txtNeunetalCharge2.SelLength = Len(txtNeunetalCharge2)
End Sub

Private Sub txtNeunetalCharge2_LostFocus()
   Label19.ForeColor = vbWhite
   Label19.FontSize = 8
   If Len(txtNeunetalCharge2) = 0 Then
      txtNeunetalCharge2 = 0
   End If
End Sub

Private Sub txtOperationCharge2_Change()
  If Not IsNumeric(txtOperationCharge2) Then
     txtOperationCharge2 = ""
     txtTotalOpr = Val(txtOperationCharge1)
  Else
     txtTotalOpr = Val(txtOperationCharge1) + Val(txtOperationCharge2)
  End If
  LOAD_TOTAL2
End Sub
Private Sub txtOperationCharge2_GotFocus()
  Label13.ForeColor = vbCyan
  Label13.FontSize = 9
  txtOperationCharge2.BackColor = &H80000018
  txtOperationCharge2.SelStart = 0
  txtOperationCharge2.SelLength = Len(txtOperationCharge2)
End Sub

Private Sub txtOperationCharge2_LostFocus()
   Label13.ForeColor = vbWhite
   Label13.FontSize = 8
   txtOperationCharge2.BackColor = vbWhite
   If Len(txtOperationCharge2) = 0 Then
      txtOperationCharge2 = 0
   End If
End Sub

Private Sub txtP_therapyCharge_Change()
If Not IsNumeric(txtP_therapyCharge) Then
      txtP_therapyCharge = ""
 Else
     txtTotal = Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(Text1) + Val(txtDeliveryCharge) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator) + Val(txtCCU_Charge) + Val(txtNebuliser) + Val(txtmisce) + Val(txtMedicine_charge)
End If

 End Sub

Private Sub txtP_therapyCharge_GotFocus()
        txtP_therapyCharge.BackColor = &H96E4B1
        txtP_therapyCharge.SelStart = 0
        txtP_therapyCharge.SelLength = Len(txtP_therapyCharge)

End Sub

Private Sub txtP_therapyCharge_LostFocus()
If txtP_therapyCharge = Empty Then
   txtP_therapyCharge = 0
End If
     txtP_therapyCharge.BackColor = &H80000018
End Sub



Private Sub txtPhotoTherapyCharge2_Change()
 If Not IsNumeric(txtPhotoTherapyCharge2) Then
    txtPhotoTherapyCharge2 = ""
    txtP_therapyCharge = Val(txtPhotoTherapyCharge1)
 Else
    txtP_therapyCharge = Val(txtPhotoTherapyCharge1) + Val(txtPhotoTherapyCharge2)
 End If
 LOAD_TOTAL2
End Sub

Private Sub txtPhotoTherapyCharge2_GotFocus()
  Label21.ForeColor = vbCyan
  Label21.FontSize = 9
  txtPhotoTherapyCharge2.SelStart = 0
  txtPhotoTherapyCharge2.SelLength = Len(txtPhotoTherapyCharge2)
  
End Sub

Private Sub txtPhotoTherapyCharge2_LostFocus()
   Label21.ForeColor = vbWhite
   Label21.FontSize = 8
   If Len(txtPhotoTherapyCharge2) = 0 Then
      txtPhotoTherapyCharge2 = 0
   End If
End Sub

Private Sub txtServiceCharge_Change()
  txtTotalOpr_Change
End Sub

Private Sub txtServiceCharge_GotFocus()
  txtServiceCharge.SelStart = 0
  txtServiceCharge.SelLength = Len(txtServiceCharge)

End Sub

Private Sub txtServiceCharge_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtExtraBedTotal.SetFocus
  End If
End Sub

Private Sub txtstaff_Change()
  If UTILITY.LOAD_STAFF(txtstaff.Text) = "0" Then
     Label16.Caption = "INVALID STAFF ID,...PLEASE VERIFY"
     Label16.ForeColor = vbRed
  Else
     txtstaff.Text = UCase(txtstaff.Text)
     Label16.Caption = UTILITY.LOAD_STAFF(txtstaff.Text)
     Label16.ForeColor = vbWhite
    
  End If
End Sub


Private Sub txtstaff_Click()
  txtstaff_Change
End Sub

Private Sub txtTotal_Change()
   If Len(txtstaff) > 0 And IsBedDiscountIrregularityOccured = False Then
     txtNetTotal = Val(txtTotal)
   Else
     txtNetTotal = Val(txtTotal) - Val(txtAdvanceRelease)
   End If
End Sub

Private Sub txtTotal2_Change()
  If Check1.Value = 1 Then
     txtdisc_percent_Change
  End If
    txtTotal = Val(txtTotal1) + Val(txtTotal2)
End Sub

Private Sub txtTotalBedCharge_Change()
  txtTotalOpr_Change
End Sub

Private Sub txtTotalBedCharge_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   txtWithAdmissionCharge.SetFocus
 End If
End Sub

Private Sub txtTotalOpr_Change()
 If Not IsNumeric(txtTotalOpr) Then
        txtTotalOpr = ""
 Else
    txtTotal = Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(Text1) + Val(txtDeliveryCharge) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator) + Val(txtCCU_Charge) + Val(txtNebuliser) + Val(txtmisce) + Val(txtMedicine_charge)
End If
End Sub

Private Sub txtTotalOpr_GotFocus()
       txtTotalOpr.BackColor = &H96E4B1
        
       txtTotalOpr.SelStart = 0
       txtTotalOpr.SelLength = Len(txtTotalOpr)

End Sub

Private Sub txtTotalOpr_LostFocus()
   If txtTotalOpr = Empty Then
        txtTotalOpr = 0
    End If
 txtTotalOpr.BackColor = &H80000018
 
End Sub

Private Sub txtWithAdmissionCharge_Change()
   txtTotalOpr_Change
End Sub
Private Sub txtWithAdmissionCharge_GotFocus()
  txtWithAdmissionCharge.SelStart = 0
  txtWithAdmissionCharge.SelLength = Len(txtWithAdmissionCharge)
End Sub

Private Sub txtWithAdmissionCharge_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtServiceCharge.SetFocus
 End If
End Sub
