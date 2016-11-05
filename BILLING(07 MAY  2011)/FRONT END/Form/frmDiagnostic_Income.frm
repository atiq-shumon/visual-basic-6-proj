VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDiagnostic_Income 
   Appearance      =   0  'Flat
   BackColor       =   &H80000014&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   450
      Top             =   3240
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
   Begin VB.TextBox txtMainName 
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
      Left            =   1230
      TabIndex        =   13
      Top             =   1890
      Width           =   4155
   End
   Begin VB.ComboBox cboMainCode 
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
      ItemData        =   "frmDiagnostic_Income.frx":0000
      Left            =   210
      List            =   "frmDiagnostic_Income.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1890
      Width           =   1005
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Test Head Wise"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   11
      Top             =   1380
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Test Wise"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3930
      TabIndex        =   10
      Top             =   960
      Width           =   1785
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Dept Details"
      BeginProperty Font 
         Name            =   "Georgia"
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
      Left            =   2100
      TabIndex        =   9
      Top             =   960
      Width           =   1785
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Dept Summary"
      BeginProperty Font 
         Name            =   "Georgia"
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
      Left            =   180
      TabIndex        =   8
      Top             =   960
      Value           =   -1  'True
      Width           =   1875
   End
   Begin VB.CommandButton CMDEXIT 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   4050
      TabIndex        =   4
      ToolTipText     =   "CLOSE"
      Top             =   3810
      Width           =   1215
   End
   Begin VB.CommandButton CMDREPORT 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   2820
      TabIndex        =   3
      ToolTipText     =   "VIEW REPORT"
      Top             =   3810
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Top             =   2730
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   12582912
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   345
      Left            =   3180
      TabIndex        =   2
      Top             =   2730
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   12582912
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0FF&
      FillColor       =   &H00FFFF00&
      Height          =   1005
      Left            =   -30
      Top             =   780
      Width           =   5505
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   3330
      Width           =   75
   End
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   2760
      Top             =   3750
      Width           =   2565
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TO  DATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   3240
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FROM DATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   270
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DIAGNOSTIC INCOME REPORT"
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
      Left            =   330
      TabIndex        =   0
      Top             =   150
      Width           =   4755
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   1
      Left            =   -390
      Picture         =   "frmDiagnostic_Income.frx":0004
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   11610
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   0
      Left            =   0
      Picture         =   "frmDiagnostic_Income.frx":5986
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11610
   End
End
Attribute VB_Name = "frmDiagnostic_Income"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UTILITY As New clsUtility

Private Sub cboMainCode_Click()
   LoadMainCode (2)
End Sub

Private Sub cmdExit_Click()
 Unload Me
End Sub
Private Sub CMDREPORT_Click()
    If UTILITY.START_END_VALIDATION(MaskEdBox1, MaskEdBox2) = False Then
      Label2.Caption = "Start Date can't be greater(>) than End date..Verify"
      MaskEdBox1.SetFocus
      Exit Sub
   End If
  Screen.MousePointer = vbHourglass
  Label2.Caption = "Please wait while processing...."
  If Option1(0).Value = True Then
      rptMode = 500
  ElseIf Option1(1).Value = True Then
       rptMode = 501
  ElseIf Option1(2).Value = True Then
       rptMode = 505
  ElseIf Option1(3).Value = True Then
      rptMode = 506
  End If
  
  Viewer.Show vbModal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     SendKeys (Chr(9))
  End If
End Sub
Private Sub Form_Load()
   MaskEdBox1.Text = Format(Date, "DD/MM/YY")
   MaskEdBox2.Text = Format(Date, "DD/MM/YY")
  
End Sub

Private Sub MaskEdBox1_Change()
  Label2.Caption = ""
End Sub

Private Sub MaskEdBox2_GotFocus()
  With MaskEdBox2
       .SelStart = 0
       .SelLength = Len(MaskEdBox2)
       
  End With
End Sub

Private Sub MaskEdBox1_GotFocus()
  With MaskEdBox1
       .SelStart = 0
       .SelLength = Len(MaskEdBox1)
  End With
  
End Sub


Private Sub Option1_Click(Index As Integer)
  Select Case Index
         Case 0, 1, 2
              cboMainCode.Enabled = False
              txtMainName.Enabled = False
         Case 3
               cboMainCode.Enabled = True
               txtMainName.Enabled = True
               LoadMainCode (1)
               cboMainCode.ListIndex = 0
               
  End Select

End Sub
Private Sub LoadMainCode(MODE As Integer)
     Select Case MODE
            Case 1
               Adodc1.ConnectionString = strcn.Connection_String
               Adodc1.RecordSource = "select distinct(m_code)from test_info_main"
               Adodc1.Refresh
      

              If Adodc1.Recordset.RecordCount > 0 Then
                  cboMainCode.clear
                   Adodc1.Recordset.MoveFirst
                  While Adodc1.Recordset.EOF = False
                    cboMainCode.AddItem Adodc1.Recordset!M_Code
                    Adodc1.Recordset.MoveNext
                  Wend
              End If
            Case 2
               Adodc1.ConnectionString = strcn.Connection_String
               Adodc1.RecordSource = "select m_NAME  from test_info_main WHERE m_CODE='" & cboMainCode & "'"
               Adodc1.Refresh
      

              If Adodc1.Recordset.RecordCount > 0 Then
                  
                   Adodc1.Recordset.MoveFirst
                  While Adodc1.Recordset.EOF = False
                    txtMainName = Adodc1.Recordset!m_name
                    Adodc1.Recordset.MoveNext
                  Wend
              End If
        End Select
    
End Sub
