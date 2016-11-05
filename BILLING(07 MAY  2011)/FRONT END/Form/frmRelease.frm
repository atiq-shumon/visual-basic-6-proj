VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRelease 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6075
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H80000000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox DischargeTypeCombo 
      Appearance      =   0  'Flat
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
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2100
      Width           =   2805
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   -60
      TabIndex        =   7
      Top             =   2970
      Width           =   6585
      Begin VB.Image Image2 
         Height          =   855
         Left            =   0
         Picture         =   "frmRelease.frx":0000
         Top             =   90
         Width           =   11820
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   -120
      TabIndex        =   6
      Top             =   -90
      Width           =   6675
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PATIENT RELEASE ENTRY"
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
         Left            =   450
         TabIndex        =   8
         Top             =   210
         Width           =   4245
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   150
         Picture         =   "frmRelease.frx":5982
         Top             =   -60
         Width           =   11820
      End
   End
   Begin VB.ComboBox CBOYRCODE 
      Height          =   315
      ItemData        =   "frmRelease.frx":B304
      Left            =   2070
      List            =   "frmRelease.frx":B306
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   900
      Width           =   2835
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   60
      Top             =   2760
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
      Left            =   4920
      TabIndex        =   3
      Top             =   2130
      Width           =   975
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
      Left            =   2085
      TabIndex        =   1
      Top             =   1440
      Width           =   2835
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Discharge Type"
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
      Left            =   180
      TabIndex        =   9
      Top             =   2130
      Width           =   1755
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
      Left            =   180
      TabIndex        =   5
      Top             =   960
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
      Left            =   180
      TabIndex        =   0
      Top             =   1545
      Width           =   915
   End
End
Attribute VB_Name = "frmRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strUid As String
Public strcn        As New MyConnection
Public UTILITY As New clsUtility

Private Sub Command1_Click()
 On Error GoTo ERR_DESC
 Dim MSG As String
  Cur_reg_no = frmRelease.txtRegNoRelease.Text
  cur_yr_code = frmRelease.CBOYRCODE.Text
  dischargeType = UTILITY.GetDischargeTypeInCode(frmRelease.DischargeTypeCombo.Text)
  
 MSG = UTILITY.GetPatientCurrentStatusInStringValue(Cur_reg_no, cur_yr_code)
 
 
 Select Case PatientStatus
        
        Case 0 '''admitted patient come to release
                  IRREGULAR_CASE = 0
                  Unload Me
                   With frmDeptTransferPatientRelease
                      .Label7.Caption = "REGULAR PATIENT RELEASE INFO. ENTRY"
                      .Show 1
                    End With
                 Exit Sub
        Case 2 ''absconded patient come to release
               IRREGULAR_CASE = 1
                Unload Me
                With frmDeptTransferPatientRelease
                      .Label7.Caption = "ABSCONDED PATIENT RELEASE INFO. ENTRY"
                      .Show 1
                    End With
                 Exit Sub
        Case 3 ''hold patient come to release
             IRREGULAR_CASE = 1
             Unload Me
              With frmDeptTransferPatientRelease
                      .Label7.Caption = "Reg.HOLD PATIENT RELEASE INFO. ENTRY"
                      .Show 1
                    End With
                 Exit Sub
       Case Else  ''' don't do anything just show message
            MsgBox MSG, vbInformation, "IT DIVISION,DNMIH"
            txtRegNoRelease.Text = ""
            Exit Sub
End Select

Exit Sub
ERR_DESC:
        MsgBox Err.Description, vbInformation, "IT DIVISION,DHMIH"
        
End Sub

Private Sub DischargeTypeCombo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Command1_Click
   End If
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
       Unload Me
  End If

End Sub
Private Sub Form_Load()
   txtRegNoRelease = ""
   
   PopulateFiscalYear
   PopulateDischargeType
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
Private Sub PopulateDischargeType()
    Dim typeList() As String
    Dim i As Integer
    typeList = UTILITY.GetDischargeTypes()
    
    For i = LBound(typeList) To UBound(typeList)
        DischargeTypeCombo.AddItem typeList(i)
    Next i
    DischargeTypeCombo.ListIndex = 0
    
End Sub
Private Sub txtRegNoRelease_Change()
 If Not IsNumeric(txtRegNoRelease) Then
   txtRegNoRelease = ""
  End If
End Sub
Private Sub txtRegNoRelease_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
            DischargeTypeCombo.SetFocus
      End If
      If KeyAscii = 27 Then
         Unload Me
      End If
      
End Sub

