VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMAIN 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   Icon            =   "MAIN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10830
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2340
      Top             =   3120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   810
      Top             =   3420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":0A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1936
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "SEARCH  PATIENT(Ctrl+F)"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Shift wise Patient Admission Statement"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      MouseIcon       =   "MAIN.frx":2810
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1200
      Top             =   8340
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
   Begin VB.TextBox txtBooth 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   8580
      TabIndex        =   8
      Top             =   3510
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1050
      Top             =   8970
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
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
   Begin VB.TextBox txtPassWord 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
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
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   7350
      PasswordChar    =   "?"
      TabIndex        =   2
      Top             =   5880
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7350
      TabIndex        =   0
      Top             =   5370
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Powered by: IT Division, DNMIH"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   5790
      TabIndex        =   20
      Top             =   4680
      Width           =   3885
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   2  'Dash
      BorderWidth     =   3
      Height          =   1515
      Left            =   5280
      Top             =   5040
      Width           =   4695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "53/1  Johnson Road ,  Dhaka-1100"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   315
      Left            =   10260
      TabIndex        =   19
      Top             =   10200
      Width           =   3945
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Dhaka National Medical Institute Hospital"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   5700
      TabIndex        =   18
      Top             =   9690
      Width           =   8805
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Billing MIS"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   1410
      TabIndex        =   17
      Top             =   750
      Width           =   3675
   End
   Begin VB.Image Image2 
      Height          =   915
      Left            =   330
      Picture         =   "MAIN.frx":36EA
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1125
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "By: Software Programmer,IT,DNMIH"
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
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Top             =   1530
      Width           =   3705
   End
   Begin VB.Label lblUserType 
      BackStyle       =   0  'Transparent
      Caption         =   "lblusertype"
      Height          =   285
      Left            =   7650
      TabIndex        =   15
      Top             =   4500
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Label lblShift 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      Height          =   315
      Left            =   7890
      TabIndex        =   14
      Top             =   5940
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label LBLDATE 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   555
      Left            =   7830
      TabIndex        =   13
      Top             =   1020
      Width           =   1260
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booth No:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11070
      TabIndex        =   11
      Top             =   2340
      Width           =   1035
   End
   Begin VB.Label lblBooth 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   405
      Left            =   12210
      TabIndex        =   10
      Top             =   2340
      Width           =   1335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Booth No"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5070
      TabIndex        =   9
      Top             =   5430
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   405
      Left            =   7230
      TabIndex        =   7
      Top             =   2220
      Width           =   1965
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name: "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5670
      TabIndex        =   6
      Top             =   2205
      Width           =   1275
   End
   Begin VB.Label lbluser_id 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   1290
      TabIndex        =   5
      Top             =   2250
      Width           =   1335
   End
   Begin VB.Label lbluser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Id:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   2250
      Width           =   825
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TYPE PASSWORD :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   5580
      TabIndex        =   3
      Top             =   5910
      Width           =   1725
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TYPE USER ID:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   5595
      TabIndex        =   1
      Top             =   5400
      Width           =   1290
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   13605
      Y1              =   45
      Y2              =   60
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -720
      Picture         =   "MAIN.frx":DF56
      Top             =   -570
      Width           =   15360
   End
   Begin VB.Menu mnuEntry 
      Caption         =   "   ~  &ENTRY  ~"
      Begin VB.Menu mnuPatient_Information 
         Caption         =   " &PATIENT INFORMATION"
         Begin VB.Menu mnuadmission 
            Caption         =   " &ADMISSION"
            Shortcut        =   ^A
         End
         Begin VB.Menu fghg 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTransfer 
            Caption         =   "WARD/ BED TRANSFER"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuDeptTrans 
            Caption         =   "DEPARTMENT TRANSFER"
            Shortcut        =   ^D
         End
         Begin VB.Menu sepdept 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOperation 
            Caption         =   " &OPERATION"
            Shortcut        =   +^{F12}
            Visible         =   0   'False
         End
         Begin VB.Menu mnunabuliser 
            Caption         =   " Nebulizer Charge"
            Shortcut        =   ^N
            Visible         =   0   'False
         End
         Begin VB.Menu mnuExtraBed 
            Caption         =   " &EXTRA BED"
            Shortcut        =   ^E
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCCUbed 
            Caption         =   " &CCU BED"
            Shortcut        =   ^C
         End
         Begin VB.Menu kjhkjhkjh 
            Caption         =   "-"
         End
         Begin VB.Menu mnutest 
            Caption         =   " &INDOOR DIAGNOSTIC TEST"
            Shortcut        =   ^T
         End
         Begin VB.Menu hgg 
            Caption         =   "-"
         End
         Begin VB.Menu mnuReAdvance 
            Caption         =   " &RE-ADVANCE ENTRY"
            Shortcut        =   ^M
         End
         Begin VB.Menu gfgfg 
            Caption         =   "-"
         End
         Begin VB.Menu mnuadmissioncancellatino 
            Caption         =   " &ADMISSION  CANCELLATION"
            Shortcut        =   ^Z
         End
         Begin VB.Menu gfdgfg 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRelease 
            Caption         =   " &PATIENT RELEASE"
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu hghghg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOutDoorPatient 
         Caption         =   " &OPD PATIENT"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuOPDPAT 
         Caption         =   "OUTCASE PATIENT"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnufdasfsda 
         Caption         =   "-"
      End
      Begin VB.Menu printSetup 
         Caption         =   " PRINTER SETUP"
         Visible         =   0   'False
      End
      Begin VB.Menu GFHJGDJH 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEXIT 
         Caption         =   " &EXIT"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuGeneralSetup 
      Caption         =   "   ~ &ADMIN  AREA ~"
      Begin VB.Menu mnuDoctor_Information 
         Caption         =   " &DOCTORS' INFORMATION"
      End
      Begin VB.Menu vvvv 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBed_Info 
         Caption         =   " &BED INFORMATION"
      End
      Begin VB.Menu nnn 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOperationInformation 
         Caption         =   " &Operation Information"
         Visible         =   0   'False
      End
      Begin VB.Menu dsads 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuChild 
         Caption         =   " &Paediaetric  Department"
         Visible         =   0   'False
      End
      Begin VB.Menu sdfd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTest_Information 
         Caption         =   " &DIAGNOSTIC TEST INFORMATION"
      End
      Begin VB.Menu mm 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShift 
         Caption         =   " &SHIFT SETUP"
         Visible         =   0   'False
      End
      Begin VB.Menu dfsdfds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddUser 
         Caption         =   " &NEW USER SETUP"
         Visible         =   0   'False
      End
      Begin VB.Menu gfhfh 
         Caption         =   "-"
      End
      Begin VB.Menu fdgfdgfg 
         Caption         =   "-"
      End
      Begin VB.Menu mnufledpat 
         Caption         =   "IRREGULAR PATIENT INFO. ENTRY"
      End
      Begin VB.Menu fdsafdsa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuD_Backup 
         Caption         =   " DATA BACKUP"
         Visible         =   0   'False
      End
      Begin VB.Menu hjhbnmjbhyty 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisc_Edit 
         Caption         =   " Special Discount "
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu MNUDOT 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "    ~  &OTHERS  ~"
      Begin VB.Menu mnuRemoveUser 
         Caption         =   "&REMOVE USER"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu gfg 
         Caption         =   "-"
      End
      Begin VB.Menu MnuChangePass 
         Caption         =   "CHANGE PASSWORD"
      End
      Begin VB.Menu fdfd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuothercharge 
         Caption         =   "OTHER CHARGE ENTRY"
      End
      Begin VB.Menu bvnbnfghxcx 
         Caption         =   "-"
      End
      Begin VB.Menu fdsafds 
         Caption         =   "&DIAGNOSTIC REFUND"
      End
      Begin VB.Menu fdsfsdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSP 
         Caption         =   "&SEARCH PATIENT"
         Shortcut        =   ^F
      End
      Begin VB.Menu fdsfd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReceipt_printing 
         Caption         =   "RECEIPT PRINTING"
      End
      Begin VB.Menu fdgfg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUser 
         Caption         =   "LOG OFF CUR. USER"
         Shortcut        =   ^L
      End
      Begin VB.Menu GFSDGFSDG 
         Caption         =   "-"
      End
      Begin VB.Menu mnutestcancellation 
         Caption         =   " &TEST CANCELLATION"
         Enabled         =   0   'False
      End
      Begin VB.Menu dsfdsfsdfds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWorkingSchedule 
         Caption         =   "ROSTER DUTY SETUP"
         Enabled         =   0   'False
      End
      Begin VB.Menu TFHDFGGF 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPatient 
         Caption         =   " &EDIT PATIENT INFORMATION"
         Enabled         =   0   'False
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "    ~  &REPORTS  ~"
      Begin VB.Menu mnu_Test_Information 
         Caption         =   "Test Information"
         Visible         =   0   'False
      End
      Begin VB.Menu GFDSGFDS 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDoctor_Info 
         Caption         =   "Doctors' Information"
         Visible         =   0   'False
      End
      Begin VB.Menu gfhgfhfghrt 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCurrentPatient 
         Caption         =   "Current Patient Statement"
      End
      Begin VB.Menu dfsfadsad 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDWDI 
         Caption         =   "&Diagnostic  Income(Department Wise)"
      End
      Begin VB.Menu ghfhgfhfg 
         Caption         =   "-"
      End
      Begin VB.Menu msdlfjlkd 
         Caption         =   "DISCOUNT STATEMENT"
         Begin VB.Menu mnuSummary 
            Caption         =   "DISCOUNT SUMMARY"
         End
         Begin VB.Menu hfghgfhg 
            Caption         =   "-"
         End
         Begin VB.Menu MNUDID 
            Caption         =   "DISCOUNT IN DETAILS"
         End
      End
      Begin VB.Menu ghjhgjhj 
         Caption         =   "-"
      End
      Begin VB.Menu mnuemployeeSpecific 
         Caption         =   "Employee Specific"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDetail 
         Caption         =   "Detail"
         Visible         =   0   'False
      End
      Begin VB.Menu fghfghgf 
         Caption         =   "&COLLECTION  STATEMENT"
      End
      Begin VB.Menu mnucollection_stat 
         Caption         =   "Collection Statistics"
         Visible         =   0   'False
      End
      Begin VB.Menu GGGG 
         Caption         =   "-"
      End
      Begin VB.Menu MNUDI 
         Caption         =   "&DEPARTMENTAL INCOME STATEMENT"
      End
      Begin VB.Menu fsdafsda 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnugrpstatement 
         Caption         =   "Group wise Statement"
         Visible         =   0   'False
      End
      Begin VB.Menu SEP0000012 
         Caption         =   "-"
      End
      Begin VB.Menu SAP999991 
         Caption         =   "-"
      End
      Begin VB.Menu MNUPAAA 
         Caption         =   "PATIENT ADMISSION"
         Begin VB.Menu MNUAS 
            Caption         =   "ADMISSION SUMMARY"
         End
         Begin VB.Menu SEP00001 
            Caption         =   "-"
         End
         Begin VB.Menu MNUAD 
            Caption         =   "ADMISSION DETAILS"
         End
      End
      Begin VB.Menu SAP00004 
         Caption         =   "-"
      End
      Begin VB.Menu mnugrprecipt 
         Caption         =   "GROUP WISE RECEIPT COLLECTION"
      End
      Begin VB.Menu sep01 
         Caption         =   "-"
      End
      Begin VB.Menu MNUAR 
         Caption         =   "ADVANCE REGISTER"
      End
      Begin VB.Menu mnuAPR 
         Caption         =   "&Abscond Patient Report"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "   ~  &ABOUT  ~"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuDsl 
      Caption         =   "&DNMIH"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "   ~  &HELP  ~"
      Visible         =   0   'False
   End
   Begin VB.Menu MNUBAKP 
      Caption         =   "~ BACKUP ~"
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cmd As New Command
Dim RS As New Recordset
Public strUid As String
Dim SFillSysObj As New Scripting.FileSystemObject
Dim UTILITY As New clsUtility
Dim tx As TextStream
Dim PASSWORD_TYPED_TIMES As Integer
Public strcn        As New MyConnection



Private Sub fdsafds_Click()
  frm_Diagnostic_refund.Show vbModal
End Sub

Private Sub fghfghgf_Click()
    frmDaily_collection_statement.Show vbModal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
       mnuEntry.Visible = False
       MnuGeneralSetup.Visible = False
       mnuReports.Visible = False
       mnuUtility.Visible = False
       lblDate.Caption = Date
       Locate_Booth
       PASSWORD_TYPED_TIMES = 1
       canonPrinterName = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Len(lbluser_id.Caption) > 0 Then
    MsgBox vbCrLf + "Thank You for using this Software" + vbCrLf + vbCrLf + "May Allah Bless You." + vbCrLf + vbCrLf + "Courtesy:" + vbCrLf + "Software Programmer," + vbCrLf + "IT Division,DNMIH", vbInformation, "IT Division,DNMIH"
    Unload Me
  End If
End Sub

Private Sub mnu_Test_Information_Click()
    Rpt_test_info.Show vbModal
End Sub

Private Sub MNUAD_Click()
    frmpatient_history.Show 1
End Sub

Private Sub mnuAddUser_Click()
     frmSecurity.Show vbModal
End Sub

Private Sub mnuAdmission_Click()
        frmIndoor_main.Show vbModal
End Sub

Private Sub mnuAdmission2_Click()

End Sub

Private Sub mnuadmissioncancellatino_Click()
  frmadmissioncancellation.Show vbModal
End Sub

Private Sub mnuAPR_Click()
   frm_abscond.Show 1
End Sub

Private Sub mnuAR_Click()
  Rpt_advance_reg.Show 1
End Sub

Private Sub mnuARRNW_Click()
  Rpt_advance_reg_REG.Show 1
End Sub



Private Sub MNUAS_Click()
  RPT_PAT_ADMISSION.Show 1
End Sub

Private Sub MNUBAKP_Click()
  mnuD_Backup_Click
End Sub

Private Sub MNUBDPR_Click()
'   PatientStatus = 2
'   frmfled.Label7.Caption = "BACK-DATED PAT. RELEASE"
'   frmfled.Label1(2).Caption = "RELEASE DATE"
'   frmfled.Show vbModal
'   frmfled.txtRegNoRelease = ""
'
End Sub

Private Sub mnuBed_Info_Click()
Bed_info.Show vbModal
End Sub
Private Sub mnuCCUbed_Click()
   formIndicator = 1
   frmReadvancepayment.Show vbModal
End Sub

Private Sub MnuChangePass_Click()
  Change_pass.Show 1
End Sub

Private Sub mnuChild_Click()
frmChild_dept.Show vbModal
End Sub

Private Sub mnucollection_stat_Click()
   frmDaily_collection_stat.Show vbModal
End Sub

Private Sub mnuCurrentPatient_Click()
   CurrentPatientUI.Show 1
End Sub

Private Sub mnuD_Backup_Click()
On Error GoTo ERR_DESC
  If SFillSysObj.FileExists("C:\\WINNT\\bkp.bat") Then
     SFillSysObj.DeleteFile ("C:\\WINNT\\bkp.bat")
  End If
  If SFillSysObj.FileExists("C:\\WINNT\\tmpSQl.sql") Then
     SFillSysObj.DeleteFile ("C:\\WINNT\\tmpSQl.sql")
  End If
  If Not SFillSysObj.FolderExists("G:\BACKUP") Then
     SFillSysObj.CreateFolder ("G:\BACKUP")
  End If
  
   If Not SFillSysObj.FileExists("C:\\WINNT\\tmpSQl.sql") Then
      With SFillSysObj
          .CreateTextFile ("C:\\WINNT\\tmpSQl.sql")
          Set tx = .OpenTextFile("C:\\WINNT\\tmpSQl.sql", ForWriting)
      End With
      tx.WriteLine ("conn system/hansaworld@bank")
      tx.WriteLine ("$EXP USERID=HOSPITAL_BILLING/NETWORK@BANK FILE='G:\BACKUP\Bill_BACKUP'")
      tx.WriteLine ("$EXP USERID=payroll/payroll@BANK FILE='G:\BACKUP\Payroll_BACKUP'")
      ''tx.WriteLine ("$EXP USERID=acct_07_08/dn_acct@BANK FILE='G:\BACKUP\accounts07_08BACKUP'")
      ''tx.WriteLine ("$EXP USERID=acct_08_09/dn_acct@BANK FILE='G:\BACKUP\accounts08_09BACKUP'")
     ''' tx.WriteLine ("$EXP USERID=acct_09_10/dn_acct@BANK FILE='G:\BACKUP\accounts09_10BACKUP'")
      tx.WriteLine ("$EXP USERID=acct_10_11/dn_acct@BANK FILE='G:\BACKUP\accounts10_11BACKUP'")
      tx.WriteLine ("$EXP USERID=NATIONAL_INVENTORY/dn_inventory@BANK FILE='G:\BACKUP\Inventory_BACKUP'")
      tx.WriteLine ("$EXP USERID=PPBF/dn_PPBF@BANK FILE='G:\BACKUP\PPBF_BACKUP'")
      tx.WriteLine ("$EXP USERID=PPBF0809/dn_PPBF@BANK FILE='G:\BACKUP\PPBF_BACKUP0809'")
      tx.WriteLine ("$EXP USERID=POPULAR_BILLING/NETWORK@BANK FILE='G:\BACKUP\POPULAR_BACKUP'")
      tx.WriteLine ("Exit")
       
  End If
  tx.Close
  If Not SFillSysObj.FileExists("C:\\WINNT\\bkp.bat") Then
                     With SFillSysObj
                           .CreateTextFile ("C:\\WINNT\\bkp.bat")
                            Set tx = .OpenTextFile("C:\\WINNT\\bkp.bat", ForWriting)
                       End With
                               
          tx.WriteLine ("SQLPLUS  HOSPITAL_BILLING/NETWORK@BANK @C:\\WINNT\\tmpSQL.SQL")
          tx.WriteLine ("Exit")
          tx.Close
    End If ''''end of export11AM.bat
   
  
 
   Shell "C:\\WINNT\\bkp.bat"
   
   
'  SFillSysObj.DeleteFile ("C:\\bkp.bat")
'  SFillSysObj.DeleteFile ("C:\\tmpSQL.SQL")
   Exit Sub
ERR_DESC:
      MsgBox Err.Description, vbCritical, " IT, DNMIH"

End Sub

Private Sub mnuDaily_Statement_Click()

End Sub

Private Sub mnuDeptTrans_Click()
   frmDeptTransfer.Show
End Sub

Private Sub mnuDetail_Click()
  Rptdiscount_detail.Show vbModal
End Sub

Private Sub mnuDI_Click()
  Rpt_Indoor_door_info.Show 1
End Sub



Private Sub MNUDID_Click()
  Rptdiscount_detail.Show 1
End Sub

Private Sub mnuDoctor_Info_Click()
Rpt_doc_info.Show vbModal
End Sub

Private Sub mnuDoctor_Information_Click()
    Doctors_info.Show vbModal
End Sub



Private Sub mnuDsl_Click()
'    FRMABOUT.Show vbModal
End Sub

Private Sub mnuDWDI_Click()
  frmDiagnostic_Income.Show 1
End Sub

Private Sub mnuEditPatient_Click()
   frmReg_for_EDIT_PAT.Show vbModal
End Sub

Private Sub mnuemployeeSpecific_Click()
  Rpt_discount_staff.Show vbModal
End Sub

Private Sub mnuEXIT_Click()
    End
End Sub

Private Sub mnuExtraBed_Click()
frmExtraBed.Show vbModal
'frmExtraBed.txtRegNoExtraBed = ""
'frmExtraBed.txtRegNoExtraBed.SetFocus

End Sub

Private Sub mnufind_Click()
  frmpatient_search.Show vbModal
End Sub

Private Sub mnufledpat_Click()
LockingFlag = True
frmIrregularPatientEntry.Show vbModal
frmIrregularPatientEntry.txtRegNoRelease = ""
frmIrregularPatientEntry.Label7.Caption = "ABSCOND  PAT. INFO. ENTRY"
'frmfled.txtRegNoRelease.SetFocus

End Sub

Private Sub mnugrprecipt_Click()
   Rpt_receipt_group.Show vbModal
End Sub

Private Sub mnugrpstatement_Click()
    frmgroup_statement.Show vbModal
 End Sub

Private Sub mnuIndoor_Click()
Rpt_Indoor_door_info.Show vbModal
End Sub

Private Sub mnunabuliser_Click()
 frmReg_nebuliser.Show vbModal
End Sub

Private Sub mnuOPDPAT_Click()
  Dim pat_info_OPD As New Pat_Info_out
  OPD_OUT_INDICATION = "OUT"
  With pat_info_OPD
      .Label9(0).Caption = "OUT-CASE PATIENT DIAGNOSTIC INFORMATION ENTRY"
      .CboDept.clear
      .CboDept.AddItem "Out-Case"
      .CboDept.Text = .CboDept.List(0)
      pat_info_OPD.Show 1
  End With
End Sub
Private Sub mnuOperation_Click()
 frmOpr_no.Show
End Sub

Private Sub mnuOperationInformation_Click()
 frmOperation_info.Show vbModal
End Sub

Private Sub mnuothercharge_Click()
   frm_otherchargereceipt.Show 1
End Sub

Private Sub mnuOutDoorPatient_Click()
 OPD_OUT_INDICATION = "OPD"
 Pat_Info_out.Show 1
End Sub

Private Sub Mnupa_Click()
  frmpatient_history.Show vbModal
End Sub

Private Sub mnupathological_Click()
    Rpt_out_door_info_summary.Show vbModal
End Sub

Private Sub mnuReAdvance_Click()
  formIndicator = 0
  frmReadvancepayment.Show 1
End Sub

Private Sub mnuReceipt_printing_Click()
   frmUlitity_release.Show 1

End Sub

Private Sub mnuRelease_Click()
     LockingFlag = False
     frmRelease.Show vbModal
     frmRelease.txtRegNoRelease = ""
End Sub

Private Sub mnuShift_Click()
frmShiftSetup.Show vbModal
End Sub

Private Sub mnuStatistics_Click()
   frmSTATISTICS.Show vbModal
End Sub

Private Sub mnuSP_Click()
   Bed_status.Show 1
End Sub

Private Sub mnuSummary_Click()
   Rpt_discount.Show vbModal
End Sub

Private Sub mnuTest_Click()
 frmReg_no.Show
 frmReg_no.txtReg_noInTest = ""
 frmReg_no.txtReg_noInTest.SetFocus
End Sub

Private Sub mnuTest_Information_Click()
    Test_info_main.Show vbModal
End Sub

Private Sub mnutestcancellation_Click()
frmtest_cancel_entry.Show vbModal
End Sub

Private Sub mnuTransfer_Click()
   frmReg_for_Bed_Transf.Show vbModal
End Sub

Private Sub mnuUser_Click()
Dim reply As String
    reply = MsgBox("Are you sure to log Off?", vbQuestion + vbYesNo, "Logging off...")
    If reply = vbYes Then
        Unload Me
        frmMAIN.Show
    End If
End Sub

Private Sub mnuwith_Click()
   Rpt__REC_DET.Show vbModal
End Sub

Private Sub mnuwithout_Click()
  Rpt_IN_out_door_info_RECEIPT.Show vbModal
End Sub

Private Sub mnuWorkingSchedule_Click()
frmWorkingSchedule.Show vbModal

End Sub



Private Sub nmuOutDoor_Click()
    Rpt_out_door_info.Show vbModal
End Sub

Private Sub printSetup_Click()
    CommonDialog1.Action = 5
End Sub

  Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(Text1) <> 0 Then
            txtpassword.SetFocus
        Else
            MsgBox "Please enter any User Id.", vbInformation, " IT, DNMIH"
        End If
    End If
End Sub

Private Sub Timer1_Timer()
  lblDate.Caption = Format(Now, " DD MMM YYYY        hh:mm AM/PM")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.index
         Case 1
            Call USER_COLLECTION
         Case 2
            Bed_status.Show 1
         Case 3
            Rpt_SHIFTWISE_PAT_ADMISSION.Show 1
  End Select
End Sub
Private Sub USER_COLLECTION()
  With Rpt_IN_out_door_info_RECEIPT
         .Option1(5).Value = True
         .rptOutCombo.Text = frmMAIN.lbluser_id
         .CboName.ListIndex = .rptOutCombo.ListIndex
         .Check1.Value = 1
         .Check2.Value = 1
         .cboShift.List(0) = Trim(frmMAIN.lblShift)
         .Show 1
          
    End With
End Sub
Private Sub txtPassWord_KeyPress(KeyAscii As Integer)
    Dim validation As Integer
  If KeyAscii = 13 Then
    If Text1 = "" Then
       MsgBox "Please Put User Id", vbCritical, "IT,DNMIH"
       Text1.SetFocus
       Exit Sub
    Else
      
            Adodc1.ConnectionString = strcn.Connection_String
            Adodc1.RecordSource = "Select user_id, user_name, user_password,user_type From Security Where (User_Id = '" & Text1 & "')"
            Adodc1.Refresh

            If Adodc1.Recordset.EOF = True Then  '''whether not id matched
                   MsgBox "Incorrect User ID", vbCritical, "Warning"
                   Text1 = ""
                   txtpassword = ""
                   Text1.SetFocus
                   Exit Sub
            Else
               If txtpassword = Adodc1.Recordset!user_password Then
                    If UTILITY.User_Shift_validation(Adodc1.Recordset!user_id, Adodc1.Recordset!user_type) = True Then
                             UserRole = UCase(Adodc1.Recordset!user_type)
                             If UCase(Adodc1.Recordset!user_type) = UCase("Admin") Then
                                            txtpassword.Visible = False
                                            Text1.Visible = False
                                            Label1.Visible = False
                                            Label2.Visible = False
                                            Label3.Visible = False
                                            txtBooth.Visible = False
                                            Shape1.Visible = False
                                            Label10.Visible = False
                                            lbluser_id.Caption = Adodc1.Recordset!user_id
                                            lblName.Caption = Adodc1.Recordset!user_name
                                            lblUserType.Caption = Adodc1.Recordset!user_type
                                            lblBooth.Caption = Val(txtBooth)
                                            mnuEntry.Visible = True
                                            mnuWorkingSchedule.Enabled = True
                                            mnutestcancellation.Enabled = True
                                            mnuEditPatient.Enabled = True
                                            MnuGeneralSetup.Visible = True
                                            mnuReports.Visible = True
                                            MNUBAKP.Visible = True
                                            mnuUtility.Visible = True
                                            Toolbar1.Buttons(1).Enabled = True
                                            Toolbar1.Buttons(2).Enabled = True
                                            Toolbar1.Buttons(3).Enabled = True

                                      ElseIf UCase(Adodc1.Recordset!user_type) <> UCase("Admin") Then
                                            txtpassword.Visible = False
                                            Text1.Visible = False
                                            Label1.Visible = False
                                            Label2.Visible = False
                                            Label3.Visible = False
                                            MNUBAKP.Visible = False
                                            txtBooth.Visible = False
                                            Shape1.Visible = False
                                            Label10.Visible = False
                                            lbluser_id.Caption = Adodc1.Recordset!user_id
                                            lblName.Caption = Adodc1.Recordset!user_name
                                            lblUserType.Caption = Adodc1.Recordset!user_type
                                            lblBooth.Caption = Val(txtBooth)
                                            mnuEntry.Visible = True
                                            mnuReports.Visible = True
                                            mnuUtility.Visible = True
                                            MnuChangePass = True
                                            Toolbar1.Buttons(1).Enabled = True
                                            Toolbar1.Buttons(2).Enabled = True
                                            mnuEditPatient.Enabled = False
                                            Toolbar1.Buttons(3).Enabled = True
                                         End If '''END OF USER TYPE
                                 
                                 
                                 
                                 Else
                                      MsgBox "Dear. " & Adodc1.Recordset!user_name & "  Your Shift is not assigned " & vbCrLf & " " & vbCrLf & "Please Contact With Administrator", vbInformation, " IT, DNMIH."
               
                              End If  ''end of shift validation
                              
                      Else
                               
                               MsgBox "WRONG PASSWORD", vbCritical, "Incorrect"
                               txtpassword.Text = ""
                               PASSWORD_TYPED_TIMES = PASSWORD_TYPED_TIMES + 1
                               If PASSWORD_TYPED_TIMES > 4 Then
                                  End
                               End If
                     End If   '''end of User Password
    
            End If      '''end of else whether not id matched
      End If
  End If  '''end of keyascii


End Sub



