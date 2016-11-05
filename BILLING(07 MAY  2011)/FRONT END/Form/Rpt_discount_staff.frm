VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Rpt_discount_staff 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Discount ' Report on Staff/Member"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3000
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
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   30
      TabIndex        =   2
      Top             =   -60
      Width           =   4185
      Begin VB.TextBox TXTSTF_NAME 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   90
         TabIndex        =   11
         Top             =   1230
         Width           =   3855
      End
      Begin VB.ComboBox txtstaffID 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   840
         Width           =   2805
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Top             =   1830
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   38197
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   2370
         TabIndex        =   4
         Top             =   1830
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   38197
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Select ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   285
         Left            =   150
         TabIndex        =   9
         Top             =   930
         Width           =   915
      End
      Begin VB.Shape Shape2 
         Height          =   495
         Left            =   1200
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Report on"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1260
         TabIndex        =   8
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "       ID  Specific"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   1140
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2370
         TabIndex        =   6
         Top             =   1590
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   5
         Top             =   1590
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   600
      Picture         =   "Rpt_discount_staff.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   2310
      Width           =   510
   End
   Begin VB.CommandButton cmdPreview 
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
      Left            =   90
      Picture         =   "Rpt_discount_staff.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Preview"
      Top             =   2310
      Width           =   510
   End
   Begin VB.Shape Shape1 
      Height          =   585
      Left            =   30
      Top             =   2220
      Width           =   1155
   End
End
Attribute VB_Name = "Rpt_discount_staff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDEXIT_Click()

Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If

End Sub
Private Sub getemp_name()
Dim var_name
On Error GoTo Errdesc
Dim conn10 As New Connection
Dim rs10 As New Recordset
Dim cmd As New Command

If conn10.State = 0 Then
conn10.ConnectionString = strcn.Connection_String
conn10.Open
End If
var_name = Rpt_discount_staff.txtstaffID.Text

cmd.ActiveConnection = conn10
cmd.CommandType = adCmdText
cmd.CommandText = "select payroll.emp_info.emp_nm  from payroll.emp_info  where upper(payroll.emp_info.emp_id)='" & Trim(var_name) & "'"
rs10.CursorLocation = adUseClient
rs10.Open cmd.CommandText, conn10, adOpenDynamic, adLockOptimistic
    If rs10.RecordCount > 0 Then
        empname = rs10.Fields(0)
    End If
Exit Sub
If conn10.State = 1 Then
    conn10.Close
    Set conn10 = Nothing
End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, " IT, DNMIH"
End Sub
Private Sub cmdPreview_Click()
   Screen.MousePointer = vbHourglass
   '''''getemp_name
  If txtstaffID = "" Then
     MsgBox "Please Enter a valid Staff Id", vbInformation, " IT, DNMIH."
     Exit Sub
  Else
       rptMode = 14
       Viewer.Show vbModal
 End If
       
End Sub

Private Sub Form_Load()
 DTPicker1.Value = Date
 DTPicker2.Value = Date
 Adodc1.ConnectionString = strcn.Connection_String
      Adodc1.RecordSource = "select PAYROLL.EMP_INFO.EMP_ID  AS EMP_ID from PAYROLL.EMP_INFO ORDER BY EMP_ID"
      Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        While Adodc1.Recordset.EOF = False
          txtstaffID.AddItem Adodc1.Recordset!EMP_ID
          
            Adodc1.Recordset.MoveNext
        Wend
    End If
    
 
' rptMode = 1
' Option1(0).Value = True
' Combo1.Text = "Medicine"
End Sub

'Private Sub Option1_Click(Index As Integer)
'Select Case Index
'    Case 0
'        If Option1(0).Value = True Then
'              IntOption = 1
'
''            Option1(1).Enabled = False
'            Combo1.Enabled = False
'        Else
''            Option1(1).Enabled = True
'            Combo1.Enabled = True
'
'        End If
'    Case 1
'        If Option1(1).Value = True Then
'             IntOption = 2
'
''            Option1(1).Enabled = True
'            Combo1.Enabled = True
'        Else
''            Option1(1).Enabled = False
'            Combo1.Enabled = False
'
'        End If
'End Select
'End Sub
Private Sub Text1_Change()

End Sub

Private Sub txtstaffID_Click()
  Adodc1.ConnectionString = strcn.Connection_String
      Adodc1.RecordSource = "select PAYROLL.EMP_INFO.EMP_NM  AS EMP_name from PAYROLL.EMP_INFO where payroll.emp_info.emp_id='" & Trim(txtstaffID.Text) & "'"
      Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
          TXTSTF_NAME.Text = Adodc1.Recordset!EMP_name
    End If
 
End Sub
