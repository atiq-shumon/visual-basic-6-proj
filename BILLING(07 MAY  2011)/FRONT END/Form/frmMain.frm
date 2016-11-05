VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H006D1812&
   Caption         =   "The State Medical Faculty of Bangladesh"
   ClientHeight    =   8310
   ClientLeft      =   -4140
   ClientTop       =   735
   ClientWidth     =   8880
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   8310
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00F3B498&
      Height          =   705
      Left            =   0
      TabIndex        =   8
      Top             =   -90
      Width           =   12975
      Begin VB.CommandButton cmdBtntool 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   7
         Left            =   7500
         Picture         =   "frmMain.frx":334B6
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Press to Exit"
         Top             =   150
         Width           =   885
      End
      Begin VB.CommandButton cmdBtntool 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   6570
         Picture         =   "frmMain.frx":338F8
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Press to Log off"
         Top             =   150
         Width           =   885
      End
      Begin VB.CommandButton cmdBtntool 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   495
         Index           =   8
         Left            =   5640
         Picture         =   "frmMain.frx":33D3A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Press to Log in"
         Top             =   150
         Width           =   885
      End
      Begin VB.CommandButton cmdBtntool 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   495
         Index           =   6
         Left            =   4710
         Picture         =   "frmMain.frx":3417C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   150
         Width           =   885
      End
      Begin VB.CommandButton cmdBtntool 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   495
         Index           =   5
         Left            =   3780
         Picture         =   "frmMain.frx":345BE
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Press to create user"
         Top             =   150
         Width           =   885
      End
      Begin VB.CommandButton cmdBtntool 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   495
         Index           =   4
         Left            =   2850
         Picture         =   "frmMain.frx":34E88
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   150
         Width           =   885
      End
      Begin VB.CommandButton cmdBtntool 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   1920
         Picture         =   "frmMain.frx":35752
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   150
         Width           =   885
      End
      Begin VB.CommandButton cmdBtntool 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   990
         Picture         =   "frmMain.frx":35B94
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Press to  Search Student"
         Top             =   150
         Width           =   885
      End
      Begin VB.CommandButton cmdBtntool 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   60
         Picture         =   "frmMain.frx":35E9E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Student Information Entry"
         Top             =   150
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FCEAE2&
         Height          =   345
         Left            =   8850
         TabIndex        =   20
         Top             =   210
         Width           =   900
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FCEAE2&
         Height          =   345
         Left            =   9960
         TabIndex        =   19
         Top             =   210
         Width           =   1125
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   570
      Top             =   3540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36768
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37494
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":377AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38088
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":384DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3932C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3977E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A022
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2370
      Top             =   3690
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
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   750
      Top             =   2310
   End
   Begin VB.TextBox txtUserId 
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   930
      TabIndex        =   0
      Top             =   5100
      Width           =   2145
   End
   Begin VB.TextBox txtPassWord 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   930
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5700
      Width           =   2145
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log on Time :"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   3
      Left            =   10740
      TabIndex        =   22
      Top             =   690
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log on Time :"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   9510
      TabIndex        =   21
      Top             =   690
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblUid 
      BackStyle       =   0  'Transparent
      Caption         =   "Emdad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4500
      TabIndex        =   18
      Top             =   690
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Id :"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   1
      Left            =   3780
      TabIndex        =   7
      Top             =   660
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   8880
      Picture         =   "frmMain.frx":3A474
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   7260
      TabIndex        =   6
      Top             =   690
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   6150
      TabIndex        =   5
      Top             =   660
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label User 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1620
      TabIndex        =   4
      Top             =   780
      Width           =   90
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "User ID :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   150
      TabIndex        =   3
      Top             =   5130
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   30
      TabIndex        =   2
      Top             =   5745
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   8325
      Left            =   -30
      Picture         =   "frmMain.frx":46723
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12075
   End
   Begin VB.Menu mnuinformation 
      Caption         =   "&[Entry]"
      Enabled         =   0   'False
      Begin VB.Menu mnuReg_entry 
         Caption         =   "Application For Registration"
         Shortcut        =   ^R
      End
      Begin VB.Menu sep_34 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAppforA_card 
         Caption         =   "Application for Admit Card"
      End
      Begin VB.Menu fdsgsfd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdmit_card 
         Caption         =   "Admit Card Issue"
         Shortcut        =   ^A
      End
      Begin VB.Menu dfgdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnust_maks 
         Caption         =   "Student marks Entry(Written)"
      End
      Begin VB.Menu gfdsgdfsg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStd_mark_prac 
         Caption         =   "Student marks Entry(Practical)"
      End
      Begin VB.Menu mnuSert 
         Caption         =   "-"
      End
      Begin VB.Menu mnuESE 
         Caption         =   "&Examiner Schedule Entry"
         Visible         =   0   'False
      End
      Begin VB.Menu gfdsgdfs 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuGeneral_setup 
      Caption         =   "  &[General Setup]"
      Enabled         =   0   'False
      Begin VB.Menu mnuUserCreation 
         Caption         =   "User Creation"
      End
      Begin VB.Menu gfsdgfds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDegreeinformation 
         Caption         =   "Degree Information"
      End
      Begin VB.Menu g_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInstitute_info 
         Caption         =   "Institute Information"
      End
      Begin VB.Menu sep_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBoard_info 
         Caption         =   "Board Information setup"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMain_course 
         Caption         =   "Main Course Setup"
      End
      Begin VB.Menu sep_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSub_course 
         Caption         =   "Level Information"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTabgrp 
         Caption         =   "Section Setup"
      End
      Begin VB.Menu fdsfsd22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSubject 
         Caption         =   "Subject Information"
      End
      Begin VB.Menu fdsabvfdgfdtre 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTabulator 
         Caption         =   "Tabulator Information"
      End
      Begin VB.Menu fghgfhgf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExamtypesetup 
         Caption         =   "Exam Type Setup"
      End
      Begin VB.Menu fdsgdfsgds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuacademic_yr 
         Caption         =   "Academic Year Setup"
      End
      Begin VB.Menu fdsfsdfsdrew 
         Caption         =   "-"
      End
      Begin VB.Menu mnusetup 
         Caption         =   "&Examiner Information Setup"
         Visible         =   0   'False
      End
      Begin VB.Menu gfdsgsd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRC 
         Caption         =   "& Registration Cancellation"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   " &[View] "
      Enabled         =   0   'False
      Begin VB.Menu mnuToolbar 
         Caption         =   "Toolbar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "   [ Utility]"
      Enabled         =   0   'False
      Begin VB.Menu mnuSS 
         Caption         =   "Student Search"
         Shortcut        =   ^F
      End
      Begin VB.Menu dsafdsafsare 
         Caption         =   "-"
      End
      Begin VB.Menu mnulogoff 
         Caption         =   "Log Off"
         Shortcut        =   ^L
      End
      Begin VB.Menu fdsfsd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUser 
         Caption         =   "Change Password"
      End
      Begin VB.Menu fdsfadsagds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDatabackup 
         Caption         =   "Data Backup"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "     &[ Report]"
      Enabled         =   0   'False
      Begin VB.Menu mnuSub_info_report 
         Caption         =   "Subject Information Report"
         Begin VB.Menu mnuEngish 
            Caption         =   "In English"
         End
         Begin VB.Menu dsfsdf 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBeng 
            Caption         =   "In Bengali"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu dfsgdfsgfds 
         Caption         =   "-"
      End
      Begin VB.Menu mnustd_gen_info 
         Caption         =   "Student General Information"
      End
      Begin VB.Menu sep1000001 
         Caption         =   "-"
      End
      Begin VB.Menu mnustd_register 
         Caption         =   "Student Register"
         Enabled         =   0   'False
      End
      Begin VB.Menu fdsafds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSt_reg_at_a_glance 
         Caption         =   "Student Registration at glance"
         Begin VB.Menu mnuInst_wise 
            Caption         =   "Institute wise"
         End
      End
      Begin VB.Menu fdgdfgdfgdfgfd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuadmit_c_i_at 
         Caption         =   "Admit Card Issue At a Glance"
         Begin VB.Menu mnuInst_wise_admit 
            Caption         =   "Institute Wise"
         End
      End
      Begin VB.Menu fsdfsdfewwewe 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTS 
         Caption         =   "Tabulation Sheet"
      End
      Begin VB.Menu hgfdhdgf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMSP 
         Caption         =   "Marksheet Preparation"
         Begin VB.Menu mnuSS1stYR 
            Caption         =   "Student Specific(1st Year)"
         End
         Begin VB.Menu dfgsfdgfdsgdsafsddfsgfdsgdfs 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSS2ND 
            Caption         =   "Student specific(2nd Year)"
         End
         Begin VB.Menu gfdgsfdsg 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSp 
            Caption         =   "Student Specific(3rd Year)"
         End
         Begin VB.Menu rgfdsgfsd 
            Caption         =   "-"
         End
         Begin VB.Menu mnuprelli 
            Caption         =   "Student Specific(Prelliminary)"
         End
         Begin VB.Menu gddfdsgfdsgsd 
            Caption         =   "-"
         End
         Begin VB.Menu mnufinal 
            Caption         =   "Student Specific(Final)"
         End
      End
      Begin VB.Menu tyettetteree 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEXSCH 
         Caption         =   "Examiner Schedule"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " &[ Help]"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu fdsfsdafsd 
         Caption         =   "-"
      End
      Begin VB.Menu mnucontents 
         Caption         =   "Contents"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdBtntool_Click(Index As Integer)
 Select Case Index
    Case 0
         Form7.Show 1
    Case 1
         Form35.Show 1
    Case 4
         '''Form30.Show 1
    Case 5
         Form22.Show 1
    Case 7
       Unload Me
    Case 8
         Unload Me
         frmMain.Show 1
         txtUserId = ""
         txtPassWord = ""
 End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If
End Sub


Private Sub Image3_Click()

End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show vbModal
End Sub

Private Sub mnuacademic_yr_Click()
 Form14.Show vbModal
End Sub

Private Sub mnuAdmit_card_Click()
   Form26.Show vbModal
End Sub

Private Sub mnuAppforA_card_Click()
   Form6.Show vbModal
End Sub

Private Sub mnuBeng_Click()
   rptmode = 211
 CRViewer1.Show vbModal
End Sub

Private Sub mnuBoard_info_Click()
 Form3.Show vbModal
End Sub

Private Sub mnuDegreeinformation_Click()
    Form1.Show vbModal
End Sub

Private Sub mnuEngish_Click()
   rptmode = 1
   CRViewer1.Show vbModal
End Sub

Private Sub mnuESE_Click()
  Form30.Show vbModal
End Sub

Private Sub mnuExamtypesetup_Click()
 Form13.Show vbModal
End Sub

Private Sub mnuexit_Click()
    Unload Me
End Sub
Private Sub txtpass_KeyPress(KeyAscii As Integer)
    If Trim(txtPass.Text) = "harun" Or Trim(txtPass.Text) = "mahbub" Or Trim(txtPass.Text) = "emdad" Then
        mnuinformation.Enabled = True
        mnuentry.Enabled = True
        mnureport.Enabled = True
        Frame1.Visible = False
    End If
End Sub

Private Sub txtuid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub
Private Sub txtuid_LostFocus()
    If Val(txtuid.Text) = 2003 Or Val(txtuid.Text) = 2001 Or Trim(txtuid.Text) = "emdad" Then
        txtPass.SetFocus
        Else
        MsgBox "Invalid user id", vbCritical, "Warning..."
        txtuid.SetFocus
        Exit Sub
    End If
End Sub

Private Sub mnuEXSCH_Click()
  Form31.Show 1
End Sub

Private Sub mnufinal_Click()
  Form28.Option1(6).Value = True
    Form28.Show 1
End Sub

Private Sub mnuInst_wise_admit_Click()
  Form25.Show vbModal
End Sub

Private Sub mnuInst_wise_Click()
    Form24.Show vbModal
End Sub

Private Sub mnuInst_wise1_Click()

End Sub

Private Sub mnuInstitute_info_Click()
 Form2.Show vbModal
End Sub

Private Sub mnulogoff_Click()
  Unload Me
  frmMain.Show vbModal
  txtUserId = ""
  txtPassWord = ""
End Sub

Private Sub mnuMain_course_Click()
 Form4.Show vbModal
End Sub

Private Sub mnuprelli_Click()
    Form28.Option1(5).Value = True
    Form28.Show 1
     
End Sub


Private Sub mnuRC_Click()
  Form34.Show vbModal
End Sub

Private Sub mnuReg_entry_Click()
 
 Form7.Show 1
End Sub

Private Sub mnusetup_Click()
  Form29.Show 1
End Sub

Private Sub mnuSp_Click()
 Form28.Option1(2).Value = True
 Form28.Show 1
 End Sub

Private Sub mnuSS_Click()
  Form35.Show 1
End Sub

Private Sub mnuSS1stYR_Click()
     Form28.Option1(0).Value = True
     Form28.Show 1
     
End Sub

Private Sub mnuSS2ND_Click()
        Form28.Option1(1).Value = True
        Form28.Show 1
        
End Sub

Private Sub mnust_maks_Click()
    tabulator_form = 1
    form15.Show vbModal
End Sub

Private Sub mnuSt_reg_at_a_glance_Click()
 ''Form23.Show vbModal
End Sub

Private Sub mnustd_gen_info_Click()
  Form17.Show vbModal
End Sub

Private Sub mnuStd_mark_prac_Click()
  tabulator_form = 2
  form15.Show vbModal
End Sub

Private Sub mnustd_register_Click()
 Form16.Show vbModal
End Sub

Private Sub mnuSub_course_Click()
 Form5.Show vbModal
End Sub

Private Sub mnuSubject_Click()
 Form10.Show vbModal
End Sub

Private Sub mnuTabgrp_Click()
   Form20.Show vbModal
End Sub

Private Sub mnuTabulator_Click()
    Form12.Show vbModal
End Sub

Private Sub mnuToolbar_Click()
  If mnuToolbar.Checked = True Then
     mnuToolbar.Checked = False
     Frame1.Visible = False
     frmMain.lblUid.Top = 30
     frmMain.Label4(1).Top = 30
     frmMain.Label4(0).Top = 30
     frmMain.Label4(2).Top = 30
     frmMain.Label4(3).Top = 30
     frmMain.Label5.Top = 30
     
  Else
   mnuToolbar.Checked = True
   Frame1.Visible = True
      frmMain.lblUid.Top = 720
     frmMain.Label4(1).Top = 720
     frmMain.Label4(0).Top = 720
     frmMain.Label4(2).Top = 720
     frmMain.Label4(3).Top = 720
     frmMain.Label5.Top = 720
 End If
End Sub

Private Sub mnuTS_Click()
    Form27.Show vbModal
End Sub

Private Sub mnuUser_Click()
    Form32.Show 1
End Sub

Private Sub mnuUserCreation_Click()
    Form22.Show vbModal
End Sub

Private Sub Timer1_Timer()
   On Error GoTo err_desc
     lblTime = Format(Now, "hh:mm:ss AM/PM")
     Exit Sub
     
err_desc:
   MsgBox Err.Description, vbCritical, cmp
   
     
End Sub

Private Sub txtPassWord_GotFocus()
   txtPassWord.BackColor = &H80000018
End Sub

Private Sub txtPassWord_LostFocus()
   txtPassWord.BackColor = vbWhite
End Sub

Private Sub txtUserId_GotFocus()
   txtUserId.BackColor = &H80000018
End Sub

Private Sub txtUserId_LostFocus()
 txtUserId.BackColor = vbWhite
'On Error GoTo err_desc
    If Len(Trim(txtUserId.Text)) = 0 Then Exit Sub
    
    
    
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "select pass,type,user_name from security where user_id='" & Trim(txtUserId.Text) & "'"
    Adodc1.Refresh
    
    If Adodc1.Recordset.RecordCount > 0 Then
       strPass = Adodc1.Recordset!Pass
       userType = Adodc1.Recordset!Type
       UserName = Adodc1.Recordset!user_name
       strUid = Trim(txtUserId.Text)
       txtPassWord.SetFocus

    Else
       MsgBox "Invalid User!!!!!!!!", vbCritical, cmp
       txtUserId.SetFocus
       Exit Sub
    End If
    
  Exit Sub
err_desc:
     MsgBox Err.Description, vbInformation, "Daffodil Software Ltd."
     txtUserId.SetFocus
End Sub

Private Sub txtPassWord_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Len(Trim(txtUserId.Text)) = 0 Then
            MsgBox "User ID Required", vbCritical, cmp
            txtUserId.SetFocus
        End If
       If Len(Trim(txtPassWord.Text)) = 0 Then Exit Sub
       If txtPassWord.Text = Trim(strPass) Then
          Label2.Visible = False
          Label3.Visible = False
          txtUserId.Visible = False
          lblUid.Visible = True
          txtPassWord.Visible = False
          Label4(0).Visible = True
          Label4(2).Visible = True
          Label4(3).Visible = True
          Label4(3).Caption = Format(Now, "hh:mm:ss AM/PM")
          lblUid.Caption = txtUserId.Text
          Label5.Visible = True
          Label4(1).Visible = True
          mnuinformation.Enabled = True
          mnuView.Enabled = True
          mnuUtility.Enabled = True
          mnureport.Enabled = True
          mnuGeneral_setup.Enabled = True
          Label5.Caption = Adodc1.Recordset!user_name
          
          enable_validation
          If UCase(validation_var) <> UCase("admin") Then
              mnuGeneral_setup.Enabled = False
          Else
              mnuGeneral_setup.Enabled = True
          End If
          Dim i As Integer
          
          For i = 0 To 8
             cmdBtntool(i).Enabled = True
          Next i
          
       Else
          MsgBox "  Invalid Password!!         ", vbInformation, cmp
          txtPassWord.SetFocus
          txtPassWord = ""
          Exit Sub
       End If
    End If
End Sub

Private Sub enable_validation()
   Adodc1.connectionstring = strcn.Connection
   Adodc1.RecordSource = "select type from security where user_id='" & frmMain.txtUserId & "'"
   Adodc1.Refresh
   
   If Adodc1.Recordset.RecordCount > 0 Then
      validation_var = Adodc1.Recordset!Type
   End If
   End Sub
