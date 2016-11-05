VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Rptdiscount_detail 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMDEXIT 
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   5970
      TabIndex        =   9
      ToolTipText     =   "CLOSE"
      Top             =   2970
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   4740
      TabIndex        =   8
      ToolTipText     =   "VIEW REPORT"
      Top             =   2970
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   2265
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7845
      Begin VB.OptionButton Option_col_staff 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Caption         =   "COLLEGE STAFF"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   225
         Left            =   210
         TabIndex        =   13
         Top             =   300
         Width           =   2145
      End
      Begin VB.OptionButton Option_gen_pat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Caption         =   "POOR PATIENT(FREE-BED ONLY)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   435
         Left            =   2760
         TabIndex        =   12
         Top             =   180
         Width           =   3585
      End
      Begin VB.OptionButton OTHER_THAN_FREE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Caption         =   "POOR PATIENT (OTHER THAN FREE-BED)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   2760
         TabIndex        =   11
         Top             =   660
         Width           =   4335
      End
      Begin VB.OptionButton option_menber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Caption         =   "COMMITTEE MEMBER"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   1050
         Width           =   2355
      End
      Begin VB.OptionButton option_hos_staff 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Caption         =   "HOSPITAL STAFF"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   690
         Width           =   2265
      End
      Begin VB.TextBox txtstaffId 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6750
         TabIndex        =   4
         Top             =   1140
         Visible         =   0   'False
         Width           =   885
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   360
         TabIndex        =   1
         Top             =   1635
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   60751873
         CurrentDate     =   38197
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   4860
         TabIndex        =   2
         Top             =   1635
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   60751873
         CurrentDate     =   38197
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0FF&
         Height          =   1395
         Left            =   -540
         Top             =   -30
         Width           =   8565
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Staff Id"
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
         Left            =   6810
         TabIndex        =   5
         Top             =   1410
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-----DATE RANGE-----"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2850
         TabIndex        =   3
         Top             =   1710
         Width           =   1815
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Reports in Details"
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
      Left            =   1440
      TabIndex        =   10
      Top             =   120
      Width           =   4755
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   -600
      Picture         =   "Rpt_discount_DETAIL.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11610
   End
   Begin VB.Shape Shape1 
      Height          =   435
      Left            =   4680
      Top             =   2940
      Width           =   2535
   End
End
Attribute VB_Name = "Rptdiscount_detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()

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
var_name = Rpt_discount_staff.txtstaffId.Text

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
  If Rptdiscount_detail.option_hos_staff = True Then
      Option_discount = 1
  ElseIf Rptdiscount_detail.Option_col_staff = True Then
     Option_discount = 2
  ElseIf Rptdiscount_detail.Option_gen_pat = True Then
     Option_discount = 3
 ElseIf Rptdiscount_detail.option_menber = True Then
     Option_discount = 4
 ElseIf Rptdiscount_detail.OTHER_THAN_FREE = True Then
     Option_discount = 5
 End If
 If Option_discount < 1 Or Option_discount > 5 Then
      MsgBox "Please Select An Option ", vbInformation, " IT, DNMIH."
     Exit Sub
Else
   '''getemp_name
'  If txtstaffId = "" Then
'     MsgBox "Please Enter a valid Staff Id", vbInformation, " IT, DNMIH."
'     Exit Sub
'  Else
       rptMode = 20
       Viewer.Show vbModal
 End If
       
End Sub

Private Sub Form_Load()
 DTPicker1.Value = Date
 DTPicker2.Value = Date
 
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

