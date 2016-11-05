VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmIrregularPatientEntry 
   Appearance      =   0  'Flat
   BackColor       =   &H80000013&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6990
   FillColor       =   &H80000000&
   ForeColor       =   &H80000000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   -60
      TabIndex        =   13
      Top             =   4380
      Width           =   7335
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Powered by :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   960
         TabIndex        =   15
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Software Programmer, IT, DNMIH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   2400
         TabIndex        =   14
         Top             =   120
         Width           =   4365
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select Patient Status"
      ForeColor       =   &H000000C0&
      Height          =   885
      Left            =   -30
      TabIndex        =   10
      Top             =   750
      Width           =   7275
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Account Hold/Backdated"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   3030
         TabIndex        =   12
         Top             =   390
         Width           =   2925
      End
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Absconded"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Index           =   0
         Left            =   540
         TabIndex        =   11
         Top             =   390
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   0
      TabIndex        =   7
      Top             =   3690
      Width           =   7455
      Begin VB.Image Image2 
         Height          =   855
         Left            =   -3630
         Picture         =   "frmfleed.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   11910
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   -90
      Width           =   7455
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "IRREGULAR  PAT. INFO. ENTRY"
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
         Left            =   1020
         TabIndex        =   8
         Top             =   240
         Width           =   4755
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -180
         Picture         =   "frmfleed.frx":5982
         Stretch         =   -1  'True
         Top             =   60
         Width           =   11070
      End
   End
   Begin VB.ComboBox CBOYRCODE 
      Height          =   315
      ItemData        =   "frmfleed.frx":B304
      Left            =   2370
      List            =   "frmfleed.frx":B306
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1920
      Width           =   2835
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   60
      Top             =   -210
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
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5280
      TabIndex        =   3
      Top             =   3090
      Width           =   1065
   End
   Begin VB.TextBox txtRegNoRelease 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2370
      TabIndex        =   2
      Top             =   3090
      Width           =   2835
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   2370
      TabIndex        =   1
      Top             =   2505
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   255
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mm-YYYY"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2220
      TabIndex        =   18
      Top             =   3180
      Width           =   75
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2220
      TabIndex        =   17
      Top             =   2580
      Width           =   75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2220
      TabIndex        =   16
      Top             =   1920
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OCCURANCE DATE "
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
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   2580
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FISCAL YEAR"
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
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1980
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " REG NO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   3135
      Width           =   885
   End
End
Attribute VB_Name = "frmIrregularPatientEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Con As New MyConnection
Dim Conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset
Public UTILITY As New clsUtility
Public strUid As String
Public strcn        As New MyConnection
Private Sub Command1_Click()
  On Error GoTo ERR_DESC
    Dim MSG As String
    Dim FLED_OBJ As New frmDeptTransferPatientRelease
    Cur_reg_no = frmIrregularPatientEntry.txtRegNoRelease
    cur_yr_code = frmIrregularPatientEntry.CBOYRCODE
    FLED_DATE = Format(MaskEdBox1, "DD/MM/YY")
   
   
   MSG = UTILITY.GetPatientCurrentStatusInStringValue(Cur_reg_no, cur_yr_code)
   If UTILITY.IsAdmissionDateLess(admissionDate, FLED_DATE) = True Then
      MsgBox "Admission Date can't be less than Occurance Date", vbInformation, "IT, DNMIH"
   Exit Sub
   End If
   
       Select Case MSG
          Case 0
              Unload Me
              IRREGULAR_CASE = 1
              With FLED_OBJ
             
                    .Label7.Caption = "IRREGULAR PATIENT FINAL CALCULATION "
             
                 .cmdSAVE.Caption = "LOCK"
                 .cmdSAVE.FontBold = True
                 .Show 1
            End With
            txtRegNoRelease.Text = ""
            Exit Sub
        Case Else
           MsgBox MSG, vbInformation, "IT DIVISION,DNMIH"
           txtRegNoRelease.Text = ""
           Exit Sub
           End Select

Exit Sub
ERR_DESC:
        MsgBox Err.Description, vbInformation, "IT DIVISION,DHMIH"

End Sub

Private Sub Form_Activate()
      txtRegNoRelease = ""
      txtRegNoRelease.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
       Unload Me
  End If

End Sub

Private Sub Form_Load()
  MaskEdBox1.Text = Format(Date, "DD/MM/YY")
  PopulateFiscalYear
  PatientStatusToBe = 2
End Sub
Private Sub PopulateFiscalYear()
   Dim yearList() As String
   Dim i As Integer
   yearList = UTILITY.GetFiscalYears()
   For i = LBound(yearList) To UBound(yearList)
       CBOYRCODE.AddItem yearList(i)
   Next i
   CBOYRCODE.ListIndex = 0
End Sub
Private Sub MaskEdBox1_GotFocus()
  With MaskEdBox1
       .SelStart = 0
       .SelLength = Len(.Text)
  End With
End Sub
Private Sub Option1_Click()
     
End Sub

Private Sub Option_Click(Index As Integer)
     Select Case Index
            Case 0
                  PatientStatusToBe = 2
            Case 1
                  PatientStatusToBe = 3
     End Select
End Sub

Private Sub txtRegNoRelease_Change()
 If Not IsNumeric(txtRegNoRelease) Then
   txtRegNoRelease = ""
  End If
End Sub

Private Sub txtRegNoRelease_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
            Command1_Click
      End If
      If KeyAscii = 27 Then
         Unload Me
      End If
      
End Sub
