VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDeptTransfer 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6225
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H80000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   345
      Left            =   1740
      TabIndex        =   0
      Top             =   1290
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   -60
      TabIndex        =   7
      Top             =   2340
      Width           =   7155
      Begin VB.Image Image2 
         Height          =   855
         Left            =   -1020
         Picture         =   "frmDeptTransfer.frx":0000
         Top             =   -90
         Width           =   11820
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   -60
      TabIndex        =   6
      Top             =   -90
      Width           =   6615
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PATIENT DEPT. TRANSFER"
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
         Left            =   930
         TabIndex        =   8
         Top             =   210
         Width           =   4245
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -4620
         Picture         =   "frmDeptTransfer.frx":5982
         Top             =   60
         Width           =   11820
      End
   End
   Begin VB.ComboBox CBOYRCODE 
      Height          =   315
      ItemData        =   "frmDeptTransfer.frx":B304
      Left            =   1740
      List            =   "frmDeptTransfer.frx":B306
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   810
      Width           =   2835
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   60
      Top             =   2400
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
      Left            =   4590
      TabIndex        =   3
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtRegNoRelease 
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
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1740
      TabIndex        =   2
      Top             =   1800
      Width           =   2835
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSFER DATE "
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
      Left            =   120
      TabIndex        =   9
      Top             =   1350
      Width           =   1620
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
      Left            =   120
      TabIndex        =   5
      Top             =   870
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "REG NO :"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1860
      Width           =   915
   End
End
Attribute VB_Name = "frmDeptTransfer"
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
If MaskEdBox1.Text = "__/__/__" Then
   MsgBox "Please put a valid date here", vbInformation, "DNMIH"
   MaskEdBox1.SetFocus
End If

 FLED_DATE = MaskEdBox1.Text
 Dim MSG As String
 MSG = UTILITY.GetPatientCurrentStatusInStringValue(txtRegNoRelease, CBOYRCODE)
 
 If UTILITY.IsAdmissionDateLess(admissionDate, FLED_DATE) = True Then
      MsgBox "Admission Date can't be less than Occurance Date", vbInformation, "IT, DNMIH"
   Exit Sub
   End If
   
 Select Case MSG
        Case 0
            frmDepartmentTransfer.Show 1
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
      
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
       Unload Me
  End If

End Sub
Private Sub Form_Load()
   MaskEdBox1.Text = Format(Date, "DD/MM/YY")
  
   PopulateFiscalYear
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
  MaskEdBox1.SelStart = 0
  MaskEdBox1.SelLength = Len(MaskEdBox1.Text)
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     txtRegNoRelease.SetFocus
  End If
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
