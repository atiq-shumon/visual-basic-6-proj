VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   $"frmMain.frx":0000
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   630
   ClientWidth     =   11880
   HelpContextID   =   999
   Icon            =   "frmMain.frx":00AA
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   11880
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.ComboBox UserRoleCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmMain.frx":0974
      Left            =   2100
      List            =   "frmMain.frx":0981
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   8070
      Width           =   1965
   End
   Begin VB.TextBox txtLogOn 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   0
      Left            =   2160
      TabIndex        =   5
      Top             =   7290
      Width           =   1860
   End
   Begin VB.TextBox txtLogOn 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2160
      PasswordChar    =   "#"
      TabIndex        =   6
      Top             =   7695
      Width           =   1860
   End
   Begin VB.Timer tmrLogOn 
      Interval        =   500
      Left            =   315
      Top             =   3600
   End
   Begin VB.Frame fraLogOn 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Log On"
      ForeColor       =   &H00808080&
      Height          =   90
      Left            =   7785
      TabIndex        =   4
      Top             =   1215
      Visible         =   0   'False
      Width           =   285
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   315
      Top             =   2970
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrClock 
      Interval        =   10
      Left            =   11430
      Top             =   675
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   495
      Top             =   1665
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2319
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3AD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DED
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4245
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4561
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5299
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B75
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6451
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":732D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7781
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7BD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7EED
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8207
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8521
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":897B
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9255
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":96A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9AF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9F4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A825
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B0FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B9D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C2B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CB8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D467
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DD41
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E61B
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EEF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F7CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":104A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11183
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11E5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12B37
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13811
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13C63
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":140B5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbrStanderd 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   " Empoyee Information  "
            ImageIndex      =   28
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "  Pay Preparation    "
            ImageIndex      =   15
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PayPrep"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Bonus"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Disburse"
                  Text            =   "Disbursement"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "  Bonus && Other Benefits  "
            ImageIndex      =   33
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "  Disbursement  "
            ImageIndex      =   32
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "  Salary Advance    "
            ImageIndex      =   36
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "  Loan  "
            ImageIndex      =   37
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "  Reports  "
            ImageIndex      =   24
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "EmpRpt"
                  Text            =   "Empoyee Information"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PayRpt"
                  Text            =   "Payroll"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PFRpt"
                  Text            =   "Provident Fund"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "  Search..."
            ImageIndex      =   38
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "  Log Off "
            ImageIndex      =   40
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   " Exit "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillColor       =   &H00400000&
         ForeColor       =   &H80000008&
         Height          =   490
         Left            =   7875
         ScaleHeight     =   495
         ScaleWidth      =   2985
         TabIndex        =   2
         Top             =   45
         Width           =   2985
         Begin VB.Label lblClock 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "12:00:00 AM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   360
            Left            =   435
            TabIndex        =   3
            Top             =   75
            Width           =   1995
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   8010
         ScaleHeight     =   510
         ScaleWidth      =   2220
         TabIndex        =   1
         Top             =   1800
         Width           =   2220
      End
   End
   Begin MSComctlLib.ImageList Img 
      Index           =   0
      Left            =   495
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14507
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":149CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14E93
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15359
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1581F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15CE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":161AB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Img 
      Index           =   1
      Left            =   1080
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16671
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16B37
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16FFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":174C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17989
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17E4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18315
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18667
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":189B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18D0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1905D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":193AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19701
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19A53
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19DA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A0F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A5BD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Img 
      Index           =   4
      Left            =   2790
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AA83
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AF49
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B40F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B8D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BD9B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Img 
      Index           =   3
      Left            =   3375
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C261
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C5B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C905
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CC57
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CFA9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Img 
      Index           =   2
      Left            =   1620
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D2FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D64D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D99F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DCF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E043
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E395
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E6E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EA39
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1ED8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F0DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F42F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Img 
      Index           =   5
      Left            =   2205
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F781
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FC47
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2010D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":205D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20A99
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20F5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21425
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":218EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21DB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22277
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2273D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22C03
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LogIn as:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   960
      TabIndex        =   13
      Top             =   8100
      Width           =   960
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Personnel Management Information System (PMIS)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   360
      TabIndex        =   12
      Top             =   9930
      Width           =   8895
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   720
      Top             =   6960
      Width           =   4125
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "IT Division, DNMIH"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   465
      Left            =   10470
      TabIndex        =   11
      Top             =   1170
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Powered By :      Software Programmer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   8370
      TabIndex        =   10
      Top             =   810
      Width           =   6105
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Height          =   240
      Index           =   1
      Left            =   990
      TabIndex        =   9
      Top             =   7695
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      Height          =   330
      Index           =   0
      Left            =   2115
      Shape           =   4  'Rounded Rectangle
      Top             =   7245
      Width           =   1950
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      Height          =   330
      Index           =   1
      Left            =   2115
      Shape           =   4  'Rounded Rectangle
      Top             =   7650
      Width           =   1950
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
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
      Height          =   240
      Index           =   0
      Left            =   990
      TabIndex        =   8
      Top             =   7290
      Width           =   795
   End
   Begin VB.Shape shpLogOn 
      BorderColor     =   &H00FFC0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   6030
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   11520
      Left            =   -510
      Picture         =   "frmMain.frx":230C9
      Top             =   600
      Width           =   15360
   End
   Begin VB.Menu mnuMan 
      Caption         =   "  &Software Management  "
      Begin VB.Menu mnuUser 
         Caption         =   "    User && Privileges"
         Shortcut        =   {F8}
      End
      Begin VB.Menu sep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCngPW 
         Caption         =   "    Change Password"
         Shortcut        =   {F11}
      End
      Begin VB.Menu sep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrinter 
         Caption         =   "    Printer preference"
         Shortcut        =   {F7}
      End
      Begin VB.Menu sep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "   Exit"
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "  &General Setup  "
      Begin VB.Menu mnuComp_Info 
         Caption         =   "    Organization Setup"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu sep20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPScale 
         Caption         =   "    Payscale"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHR 
         Caption         =   "    House Rent Allowance"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBns 
         Caption         =   "    Other Parameters"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOffice_Time 
         Caption         =   "    Holiday,Leave && Office Time"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuline 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclosing 
         Caption         =   "   Closing Balance Entry(For PF)"
      End
      Begin VB.Menu mnunext 
         Caption         =   "-"
      End
      Begin VB.Menu mnufiscal 
         Caption         =   "   Fiscal Year SetUp"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuleaveapp 
      Caption         =   "Leave Information"
      Begin VB.Menu mnuleavemenu 
         Caption         =   "   Earn Leave Setup "
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnu55 
         Caption         =   "-"
      End
      Begin VB.Menu mnuleaveapplication 
         Caption         =   "   Leave Application"
         Shortcut        =   ^{F7}
      End
   End
   Begin VB.Menu mnuInPr 
      Caption         =   "Increment & &Promotion"
      Begin VB.Menu mnuincrement 
         Caption         =   "   Increment Information "
      End
      Begin VB.Menu mnuline10 
         Caption         =   "-"
      End
      Begin VB.Menu mnupromotion 
         Caption         =   "   Promotion Information"
      End
   End
   Begin VB.Menu Mymenu 
      Caption         =   "  &Entry  "
      WindowList      =   -1  'True
      Begin VB.Menu emp_info 
         Caption         =   "    Employee Information"
         Shortcut        =   ^E
      End
      Begin VB.Menu sep8 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAttn 
         Caption         =   "    Attendance Register"
         Shortcut        =   ^T
         Visible         =   0   'False
      End
      Begin VB.Menu Sep98 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuleave 
         Caption         =   "    Leave Register"
         Shortcut        =   ^L
         Visible         =   0   'False
      End
      Begin VB.Menu sep97 
         Caption         =   "-"
      End
      Begin VB.Menu Payprep 
         Caption         =   "    Salary Preparation"
         Shortcut        =   ^S
      End
      Begin VB.Menu sep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnusalaraynextmonth 
         Caption         =   "    Salary Preparation for the Next Month"
      End
      Begin VB.Menu mnuline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnusalydisbursement 
         Caption         =   "    Salary Disbursment"
      End
      Begin VB.Menu mnuline3 
         Caption         =   "-"
      End
      Begin VB.Menu Bns 
         Caption         =   "    Bonus Preparation"
         Shortcut        =   ^B
      End
      Begin VB.Menu sep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOT 
         Caption         =   "    Overtime Preparation"
         Shortcut        =   ^O
      End
      Begin VB.Menu sep99 
         Caption         =   "-"
      End
      Begin VB.Menu Disb 
         Caption         =   "    Disbursement"
         Shortcut        =   ^D
      End
      Begin VB.Menu sep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalAdv 
         Caption         =   "    Salary Advance && Loan"
         Shortcut        =   ^A
      End
      Begin VB.Menu m0005 
         Caption         =   "-"
      End
      Begin VB.Menu m0001 
         Caption         =   "   Gratuity Fund"
      End
      Begin VB.Menu m0002 
         Caption         =   "-"
      End
      Begin VB.Menu m0003 
         Caption         =   "   Provident Fund"
      End
      Begin VB.Menu m0004 
         Caption         =   "-"
      End
      Begin VB.Menu mnuX 
         Caption         =   "    Exit"
         Shortcut        =   ^X
         Visible         =   0   'False
      End
      Begin VB.Menu mnuJobEnd 
         Caption         =   "    Job Ending"
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "Utility"
      Begin VB.Menu mnuIncomeTaxCalc 
         Caption         =   "Income Tax Calculator"
      End
   End
   Begin VB.Menu mnuRpt 
      Caption         =   "Reports"
      Begin VB.Menu mnuRptMan 
         Caption         =   "    Report Manager"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuline02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuemployee 
         Caption         =   "    Employee Report"
      End
      Begin VB.Menu mnudsfsd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuretired 
         Caption         =   "    Employee Report (Retiered)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "  &Help  "
      Begin VB.Menu mnuHlpCont 
         Caption         =   "    Contents..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu sep18 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "    Search..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu Sep19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "    About..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SecureX As New Security
Dim Pass As String
Dim U_Id As String
'Open_Screen is a module level subroutine
'it resides in the Fx_Sub.bas file
'it takes two arguments and checks permission to a
'particular form or scope for a particular user.
Private Sub Bns_Click()
Dim f As New frmBonusPreparation
f.Show 1
End Sub
Private Sub Command1_Click()
    Enable_Menu True
End Sub
Private Sub Command2_Click()
    Enable_Menu False
End Sub
Private Sub Disb_Click()
 ' Disbursement
   'Open_Screen Form13, U_Id
'    Dim f As New frmPaySlipInfoReport
'    f.Show 1
   MsgBox "Opearation Restricted", vbInformation, organizationInfo
End Sub
Private Sub emp_info_Click()
  'Employee Information
    Form2.Show 1
    End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then End
End Sub
Private Sub Form_Load()
   On Error Resume Next
    Set_MenuImage   ' set bitmap into sub menus
    Company_Nm_Add
    tmrLogOn.Enabled = True
    UserRoleCombo.ListIndex = 0
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu Mymenu
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MsgBox vbCrLf + "Thank You for using this Software" + vbCrLf + vbCrLf + "May Allah Bless You." + vbCrLf + vbCrLf + "Courtesy:" + vbCrLf + "Software Programmer," + vbCrLf + "IT Division,DNMIH", vbInformation, "IT Division,DNMIH"
    Destroy Me
End Sub

Private Sub m0001_Click()
Dim f As New frmGratuity
f.Show 1
End Sub

Private Sub m0003_Click()
Dim f As New frmPF
f.Show 1
End Sub

Private Sub mnuBns_Click()
'Bonus Setup
    Form12.Show 1
End Sub

Private Sub mnuclosing_Click()
Dim f As New frmOpeningBalance
f.Show 1
End Sub

Private Sub mnuComp_Info_Click()
'Company Information
    Form9.Show 1
End Sub

Private Sub mnuemployee_Click()
Dim f As New frmEmployeeRpt
f.Show 1
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnufiscal_Click()
Dim f As New frmFiscalYear
f.Show 1
End Sub

Private Sub mnuHR_Click()
'House rent setup
    Form5.Show
End Sub

Private Sub mnuJob_Click()
'Job Type & Designation Setup
    Open_Screen Form9, U_Id
End Sub

Private Sub mnuLn_Click()
'Loan Application
    Open_Screen Form11, U_Id
End Sub

Private Sub mnuLnSt_Click()
'Loan Setup
   Open_Screen Form10, U_Id
End Sub

Private Sub mnuPF_Click()
'PF parameter setup
    Open_Screen Form12, U_Id
End Sub

Private Sub mnuIncomeTaxCalc_Click()
  frmITCalc.Show 1
End Sub

Private Sub mnuincrement_Click()
Dim f As New frmIncrement
f.Show 1
End Sub

Private Sub mnuJobEnd_Click()
     MsgBox "Operation Restricted", vbInformation, "IT Division, DNMIH"
     'Open_Screen Form23, U_Id
End Sub

Private Sub mnuleaveapplication_Click()
    'Open_Screen frmLeaveApplication, U_Id
    Dim f As New frmLeaveApplication
    f.Show 1
End Sub

Private Sub mnuleavemenu_Click()
Dim f As New frmLeaveRegistryForEmp
f.Show 1
End Sub

Private Sub mnuOffice_Time_Click()
 Form10.Show 1
  
End Sub

Private Sub mnuOT_Click()
   ' Open_Screen Form22, U_Id
   Dim f As New Form22
   f.Show 1
End Sub


Private Sub mnuPrinter_Click()
'Printer Selection
    CommonDialog1.Action = 5
End Sub
Private Sub mnuProd_info_Click()
   ' Open_Screen Form6, U_Id
End Sub

Private Sub mnupromotion_Click()
Dim f As New frmPromotionInfo
f.Show 1
End Sub

Private Sub mnuPScale_Click()
'Payscale Setup
    Form7.Show 1
End Sub

Private Sub mnuredf_Click()
Dim f As New frmTowhomitConcern
f.Show 1
End Sub

Private Sub mnuretired_Click()
Dim f As New Form4
f.Show 1
End Sub

Private Sub mnuRptMan_Click()
'Report Manager
    Form21.Show 1
End Sub

Private Sub mnuSalAdv_Click()
   On Error Resume Next
    'Salary Advance
   Form11.Show 1
End Sub

Private Sub mnusalaraynextmonth_Click()
Dim f As New frmSalaryfornextmonth
f.Show 1
End Sub

Private Sub mnusalydisbursement_Click()
'Dim f As New frmSalaryDisburs
'f.Show 1
 Dim f As New frmPaySlipInfoReport
        f.Show 1

End Sub

Private Sub mnuSearch_Click()
   On Error Resume Next
    'Search Screen
    Open_Screen frmSearch, U_Id
    
End Sub

Private Sub mnuSubs_Click()
    'Subscription Rate Setup
    'Open_Screen form14, U_Id
End Sub


Private Sub mnutest_Click()

    rptmode = 19
    'EmployeeName
    Form20.Show vbModal
End Sub

Private Sub mnuUser_Click()
'User and privilege
    Form90.Show 1
End Sub

Private Sub Payprep_Click()
'Pay preparation
  If UserRoleCombo.Text <> "Personnel" Then
   Form3.Show 1
  End If
 End Sub

Private Sub tmrLogOn_Timer()
        LogOn   'LogOn   ' disables menu and toolbar & shows log on frame
        tmrLogOn.Enabled = False
End Sub

Private Sub txtLogOn_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
           Case 0
             If txtLogOn(0) <> Empty And KeyAscii = 13 Then
                 txtLogOn(1).SetFocus
              End If
           Case 1
             If txtLogOn(1) <> Empty And KeyAscii = 13 Then
                U_Id = Trim(txtLogOn(0))
                Pass = Trim(txtLogOn(1))
                UserRoleCombo.SetFocus
        End If
    End Select
  
                
   

End Sub

Private Sub UserRoleCombo_keypress(KeyAscii As Integer)
    ' Validates user access during log on
    If KeyAscii = 13 Then
                If Len(txtLogOn(0)) = 0 Then
                   MsgBox "User Id Required", vbInformation, "Software Programmer,IT,DNMIH"
                   txtLogOn(0).SetFocus
                   Exit Sub
               ElseIf Len(txtLogOn(1)) = 0 Then
                  MsgBox "User Name Required", vbInformation, "Software Programmer,IT,DNMIH"
                  txtLogOn(1).SetFocus
                   Exit Sub
               End If
    
                With SecureX                'instance of security class
                        .Connstring = strCN.Connection_String
                        
                'ValidateUser method checks and lets authentic
                'user gain access to the software
                       If .ValidateUser(U_Id, Pass, UserRoleCombo.Text) = False Then
                            SetFocus_To txtLogOn(1)
                       Else
                            ''  Validation fails
                            ''  txtLogOn(0) = Empty
                            txtLogOn(1) = Empty

                            Enable_Menu True
                            txtLogOn(0).Visible = False
                            txtLogOn(1).Visible = False
                            Label5.Visible = False
                            UserRoleCombo.Visible = False
                            Label1(0).Visible = False
                            Label1(1).Visible = False
                            shpLogOn.Visible = False
                            Shape1(0).Visible = False
                            Shape1(1).Visible = False
                            If UserRoleCombo.Text = "Admin" Then
                              mnuMan.Enabled = True
                              mnuSetup.Enabled = True
                              mnuUtility.Enabled = True
                              mnusalaraynextmonth.Enabled = True
                              Payprep.Enabled = True
                            ElseIf UserRoleCombo.Text = "Accounts" Then
                              mnuMan.Enabled = False
                              mnuUtility.Enabled = True
                              mnuSetup.Enabled = False
                            ElseIf UserRoleCombo.Text = "Personnel" Then
                              mnuMan.Enabled = True
                              mnuSetup.Enabled = True
                              mnuUtility.Enabled = False
                              mnusalaraynextmonth.Enabled = False
                              Payprep.Enabled = False
                              mnuMan.Enabled = False
                              mnuSetup.Enabled = False
                              mnuOT.Enabled = False
                              mnuSalAdv.Enabled = False
                              Bns.Enabled = False
                              Disb.Enabled = False
                            Else
                              mnuMan.Enabled = False
                              mnuSetup.Enabled = False
                            End If
                            UserRole = UserRoleCombo.Text
                       End If

                End With
        
    End If
End Sub

Private Sub tlbrStanderd_ButtonClick(ByVal Button As MSComctlLib.Button)

''Toolbar--------------------


    Select Case Button.Index

        Case 1: emp_info_Click
        
        Case 3: Payprep_Click
'        Case 5: Bns_Click
'        Case 7: Disb_Click
'        Case 9: mnuSalAdv_Click
'        Case 11: mnuLn_Click
'        Case 13: mnuRptMan_Click
'        Case 15: mnuSearch_Click
        Case 17: LogOff
        Case 19: Close_Msg Me

    End Select

End Sub

Private Sub tmrClock_Timer()
  'Shows clock on the main screen
    lblClock = Format(Now, "hh:mm:ss AM/PM")
End Sub

Private Sub Set_MenuImage()

''called during form load event
''which sets bitmap into submenus

   On Error Resume Next

    Dim hMenu As Long
    Dim C As Integer
    Dim SubMenu As Long
    Dim i As Integer
    Dim Menu_ID As Long
    Dim X As Long

    hMenu = GetMenu(hwnd)

    For C = 0 To 5
        SubMenu = GetSubMenu(hMenu, C)
        For i = 1 To 18
            Menu_ID = GetMenuItemID(SubMenu, i - 1)
            X = SetMenuItemBitmaps(SubMenu, Menu_ID, 0, Img(C).ListImages(i).Picture, Img(C).ListImages(i).Picture)
        Next i
    Next C

End Sub

Public Sub Enable_Menu(Oparation_Mode As Boolean)

'called during Log On and Log Off of the software
'which enables or disables menus and toolbar
'Oparation_Mode=yes enables menus and toolbar and vice varsa

On Error Resume Next
    Dim MyObj As Object

    For Each MyObj In Screen.ActiveForm
        If TypeOf MyObj Is Menu Then
            If Oparation_Mode = True Then
               MyObj.Enabled = True
                tlbrStanderd.Enabled = True
            Else
               MyObj.Enabled = False
               tlbrStanderd.Enabled = False
            End If
        End If
    Next

End Sub
Public Sub LogOff()

'called during  click event of toolbar button "Log Off"
'which on confirmation disables menus and toolbar
    If MsgBox("Do you really want to log off current session?", vbYesNo + vbQuestion + vbDefaultButton1, "Log Off") = vbYes Then
         U_Id = Empty
         Enable_Menu False
         fraLogOn.Visible = True
         shpLogOn.Visible = True
         SetFocus_To txtLogOn(0)
    End If

End Sub

Public Sub LogOn()
'called during form1 activate event
'which disables menus and toolbar
         Enable_Menu False
         fraLogOn.Visible = True
         shpLogOn.Visible = True
         SetFocus_To txtLogOn(0)
End Sub

