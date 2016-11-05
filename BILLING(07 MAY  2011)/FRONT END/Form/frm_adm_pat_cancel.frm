VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmadm_pat_cancel 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5535
   ClientLeft      =   -105
   ClientTop       =   390
   ClientWidth     =   9315
   FillColor       =   &H80000001&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   33
      Top             =   5190
      Width           =   13365
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Software Programmer, IT Division, DNMIH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2940
         TabIndex        =   35
         Top             =   60
         Width           =   4725
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Developed && Maintenanced by:"
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
         Height          =   225
         Left            =   150
         TabIndex        =   34
         ToolTipText     =   "Developed && Maintenanced by:A.N.M. Atiqur Rahman,Software Programmer, IT Division, DNMIH"
         Top             =   90
         Width           =   2715
      End
   End
   Begin VB.TextBox TXTREC_NO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6360
      TabIndex        =   32
      Top             =   4290
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   7740
      TabIndex        =   3
      Top             =   4590
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   6510
      TabIndex        =   2
      Top             =   4590
      Width           =   1215
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   4590
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Height          =   1005
      Left            =   -30
      TabIndex        =   23
      Top             =   3300
      Width           =   9435
      Begin VB.TextBox txtCurpayment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   495
         Width           =   1875
      End
      Begin VB.TextBox txtTotalPayment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   7020
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   495
         Width           =   1875
      End
      Begin VB.TextBox TxtPreviousPayment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   330
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   495
         Width           =   1875
      End
      Begin VB.Label Label13 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2790
         TabIndex        =   30
         Top             =   390
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6240
         TabIndex        =   29
         Top             =   420
         Width           =   255
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Return"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   7020
         TabIndex        =   28
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back to Patient"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   3690
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   1890
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   0
      TabIndex        =   22
      Top             =   -90
      Width           =   9315
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADMISSION CANCELLATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   315
         Left            =   2580
         TabIndex        =   31
         Top             =   330
         Width           =   4755
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -360
         Picture         =   "frm_adm_pat_cancel.frx":0000
         Top             =   30
         Width           =   11820
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   2625
      Left            =   -60
      TabIndex        =   13
      Top             =   690
      Width           =   9405
      Begin VB.TextBox txtAgeInTest 
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
         Height          =   330
         Left            =   4875
         Locked          =   -1  'True
         MaxLength       =   17
         TabIndex        =   8
         Top             =   1275
         Width           =   555
      End
      Begin VB.TextBox txtAddrInTest 
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
         Left            =   330
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1920
         Width           =   8625
      End
      Begin VB.TextBox txtPat_ID1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc6"
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
         Height          =   330
         Left            =   330
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   510
         Width           =   4110
      End
      Begin VB.TextBox txtNameInTest 
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
         Height          =   330
         Left            =   330
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1260
         Width           =   4125
      End
      Begin VB.ComboBox cboDept 
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
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
         Height          =   315
         Left            =   4875
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "cboInTestDept"
         Top             =   510
         Width           =   2040
      End
      Begin VB.ComboBox cboInTestSex 
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
         Height          =   315
         ItemData        =   "frm_adm_pat_cancel.frx":5982
         Left            =   5520
         List            =   "frm_adm_pat_cancel.frx":598C
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1275
         Width           =   1365
      End
      Begin VB.ComboBox cboInTestReligion 
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
         Height          =   315
         ItemData        =   "frm_adm_pat_cancel.frx":5996
         Left            =   7020
         List            =   "frm_adm_pat_cancel.frx":59A9
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1260
         Width           =   1920
      End
      Begin MSComCtl2.DTPicker Dt_date 
         Height          =   330
         Left            =   6990
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   510
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   49152
         Format          =   58916865
         CurrentDate     =   37114
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   4875
         TabIndex        =   21
         Top             =   1005
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   6990
         TabIndex        =   20
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reg. No "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   330
         TabIndex        =   19
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   360
         TabIndex        =   18
         Top             =   1605
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   5520
         TabIndex        =   17
         Top             =   1005
         Width           =   375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   285
         TabIndex        =   16
         Top             =   1005
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   4875
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7020
         TabIndex        =   14
         Top             =   1005
         Width           =   885
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   240
      Top             =   0
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
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
      Left            =   360
      Top             =   0
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   120
      Top             =   0
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   3420
      Top             =   4710
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   4770
      Top             =   4620
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   360
      Top             =   0
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   120
      Top             =   0
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
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   5190
      Top             =   4530
      Width           =   3825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   -6660
      TabIndex        =   12
      Top             =   4770
      Width           =   270
   End
End
Attribute VB_Name = "frmadm_pat_cancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Con As New MyConnection
Dim Conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset
Dim RS1 As New Recordset
Dim Conn2 As New Connection
Dim rs2 As New Recordset
Public strUid As String
Dim UTILITY As New clsUtility
Dim VoucherNumber
Public strcn        As New MyConnection
Private Sub cmdExit_Click()
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub
Private Sub cmdPrint_Click()
  TXTREC_NO.Visible = True
  
  If TXTREC_NO = "" Then
     TXTREC_NO.SetFocus
  Else
    rptMode = 30
    Viewer.Show vbModal
  End If
End Sub
Private Sub cmdSave_Click()
If UTILITY.User_Shift_validation(frmMAIN.lbluser_id, frmMAIN.lblUserType.Caption) = False Then
    MsgBox "Dear. " & frmMAIN.lblName & "  Your Shift has been Expired.. " & vbCrLf & " " & vbCrLf & "Please Contact With Administrator", vbInformation, " IT, DNMIH."
    Exit Sub
End If

If Val(txtCurpayment) = Empty Then
    txtCurpayment = 0
    MsgBox "Nothing To Save", vbInformation, " IT, DNMIH."
End If

    
    Adodc4.ConnectionString = strcn.Connection_String
    Adodc4.RecordSource = "SELECT CANCELLATION_FLAG  AS FLAG FROM IN_DOOR_PAT_INFO_MAIN WHERE IN_REG_NO='" & frmadm_pat_cancel.txtPat_ID1 & "' AND YRCODE='" & Trim(frmadmissioncancellation.CBOYRCODE.Text) & "'"
    Adodc4.Refresh
    
    If Adodc4.Recordset!FLAG = "1" Then
         MsgBox "Admission has already been Cancelled", vbCritical, " IT, DNMIH."
         Exit Sub
    End If

  
    Call save_admcancellation
    MsgBox "Operation successful", vbInformation + vbOKOnly, "Save..."
    print_cancellation
    CMDEXIT.SetFocus
End Sub
Private Sub print_cancellation()
    Dim Conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset
   
   If Conn.State = 0 Then
      Conn.Open strcn.Connection_String
   End If
   
   Set cmd.ActiveConnection = Conn
   cmd.CommandType = adCmdText
            
   Dim Report6   As New CrystalReporttran
   Dim Param1 As New Parameter
  
   Set Param1 = cmd.CreateParameter("param1", adSingle, adParamInput, 100, frmadm_pat_cancel.TXTREC_NO)
   cmd.Parameters.Append Param1 '
                 
   cmd.Properties("PLSQLRSet") = True
            
   cmd.CommandText = "{CALL Rpt_in_dr_info_adm_print(?)}"
   Set RS = cmd.Execute
   cmd.Properties("PLSQLRSet") = False
   Report6.Text4.Width = 4000
   Report6.Text4.SetText ("Admission Cancellation")
          
   Report6.Database.SetDataSource RS

   Report6.PrintOut
   RS.Close
   
   If Conn.State = 1 Then
      Conn.Close
   End If

End Sub
Private Sub save_admcancellation()
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
       
  If Conn.State = 0 Then
      Conn.Open strcn.Connection_String
  End If
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSingle, adParamInput, 30, frmadmissioncancellation.txtReg_noInTest)
    cmd.Parameters.Append Param1 'IN_REG_NO
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 30, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param2 'U_id default Sumon
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 30, frmMAIN.lblBooth)
    cmd.Parameters.Append Param3 'booth
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 10, frmadmissioncancellation.CBOYRCODE)
    cmd.Parameters.Append Param4 'readvance
    
    If Len(txtCurpayment) = 0 Then
        txtCurpayment = 0
    End If
    
    Set Param5 = cmd.CreateParameter("param5", adSingle, adParamInput, 10, Val(txtCurpayment))
    cmd.Parameters.Append Param5 'previous advance
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 20, Trim(CboDept))
    cmd.Parameters.Append Param6 'department


    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL adm_cancellation(?,?,?,?,?,?)}"
    
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    
   If Conn.State = 1 Then
      Conn.Close
      Set Conn = Nothing
      Set RS = Nothing
      Set cmd = Nothing
   End If
   
   Adodc4.ConnectionString = strcn.Connection_String
   Adodc4.RecordSource = "SELECT MAX(REC_NO) AS REC_NO FROM RECEIPT_NO_COUNTER"
   Adodc4.Refresh
   If Adodc4.Recordset.RecordCount > 0 Then
      TXTREC_NO.Text = Adodc4.Recordset!REC_NO
    End If
  End Sub
Private Sub Form_Activate()
    txtCurpayment.SetFocus
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
End If
End Sub
Private Sub flush_grid()
    Adodc4.ConnectionString = strcn.Connection_String
    Adodc4.RecordSource = "select advance,dt from advance where in_reg_no='" & Trim(frmadmissioncancellation.txtReg_noInTest) & "' AND YRCODE='" & Trim(frmadmissioncancellation.CBOYRCODE) & "'"
    Adodc4.Refresh
'
End Sub
Private Sub Form_Load()
     Dim temp
    If Conn2.State = 0 Then
        Conn2.ConnectionString = strcn.Connection_String
        Conn2.Open
      End If
        cmd.ActiveConnection = Conn2
        cmd.CommandType = adCmdText
        cmd.CommandText = "select pat_name,pat_guard_name,sex,age,religion,addr1,phone,doc_dept,admission_date  From in_door_pat_info_main Where in_reg_no ='" & Trim(frmadmissioncancellation.txtReg_noInTest.Text) & "' AND YRCODE='" & Trim(frmadmissioncancellation.CBOYRCODE) & "'"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        rs2.CursorLocation = adUseClient

        rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False

                  
       If rs2.RecordCount > 0 Then
             txtNameInTest = rs2!pat_name
             If Not IsNull(rs2!addr1) Then
                 txtAddrInTest = rs2!addr1
            End If
       
         txtAgeInTest = rs2!age
         Dt_date.Value = rs2!admission_date
         cboInTestSex.Text = rs2!sex
         cboInTestReligion = rs2!religion
         CboDept = rs2!doc_dept
       End If
       
       txtPat_ID1 = frmadmissioncancellation.txtReg_noInTest
       total_adv
   txtCurpayment = TxtPreviousPayment
    txtTotalPayment = TxtPreviousPayment
    
    If Conn2.State = 1 Then
        Conn2.Close
        Set Conn2 = Nothing
    End If
  
   End Sub
Private Sub total_adv()
        Adodc5.ConnectionString = strcn.Connection_String
        Adodc5.RecordSource = "select  nvl(sum(advance),0)as advance from advance where in_reg_no ='" & Trim(frmadmissioncancellation.txtReg_noInTest.Text) & "' AND YRCODE='" & Trim(frmadmissioncancellation.CBOYRCODE) & "'"
        Adodc5.Refresh
        
        If Adodc5.Recordset.RecordCount > 0 Then
            TxtPreviousPayment = Adodc5.Recordset!advance
        End If
        
End Sub
