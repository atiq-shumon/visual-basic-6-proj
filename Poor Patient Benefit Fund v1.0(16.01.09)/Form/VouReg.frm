VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form9 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voucher Register"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "VouReg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   825
      Left            =   -30
      TabIndex        =   16
      Top             =   -120
      Width           =   6825
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Register"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4170
         TabIndex        =   17
         Top             =   210
         Width           =   2565
      End
   End
   Begin VB.TextBox txtUnitCode 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3915
      TabIndex        =   15
      Top             =   3330
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtVouType 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3285
      TabIndex        =   14
      Top             =   3330
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtQueryPart 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3555
      TabIndex        =   13
      Top             =   3195
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.ComboBox cboReportType 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   315
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Voucher No. Wise"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   2460
      Width           =   1725
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Voucher Date Wise"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   1725
      Width           =   1725
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "All"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   360
      TabIndex        =   6
      Top             =   3180
      Width           =   1725
   End
   Begin VB.TextBox txtvou_no 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4575
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2415
      Width           =   1740
   End
   Begin VB.CommandButton cmdPREVIEW 
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
      Left            =   5265
      Picture         =   "VouReg.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Preview"
      Top             =   3045
      Width           =   510
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
      Left            =   5805
      Picture         =   "VouReg.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Exit"
      Top             =   3045
      Width           =   510
   End
   Begin MSComCtl2.DTPicker dted_dt 
      Height          =   285
      Left            =   4815
      TabIndex        =   3
      Top             =   1725
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   55508993
      CurrentDate     =   36961
   End
   Begin MSComCtl2.DTPicker dtst_dt 
      Height          =   285
      Left            =   2835
      TabIndex        =   2
      Top             =   1725
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   55508993
      CurrentDate     =   36961
   End
   Begin VB.Shape Shape3 
      Height          =   465
      Left            =   5220
      Top             =   3030
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Report Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   315
      TabIndex        =   12
      Top             =   855
      Width           =   885
   End
   Begin VB.Shape Shape1 
      Height          =   315
      Index           =   2
      Left            =   315
      Top             =   3135
      Width           =   1875
   End
   Begin VB.Shape Shape1 
      Height          =   315
      Index           =   1
      Left            =   315
      Top             =   2415
      Width           =   1875
   End
   Begin VB.Shape Shape2 
      Height          =   360
      Index           =   1
      Left            =   4770
      Top             =   1680
      Width           =   1545
   End
   Begin VB.Shape Shape2 
      Height          =   360
      Index           =   0
      Left            =   2790
      Top             =   1680
      Width           =   1545
   End
   Begin VB.Shape Shape1 
      Height          =   315
      Index           =   0
      Left            =   315
      Top             =   1680
      Width           =   1875
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4500
      TabIndex        =   11
      Top             =   1770
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2340
      TabIndex        =   10
      Top             =   1770
      Width           =   345
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher #"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3735
      TabIndex        =   9
      Top             =   2460
      Width           =   750
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   645
      Index           =   2
      Left            =   180
      Top             =   1545
      Width           =   6315
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   645
      Index           =   3
      Left            =   180
      Top             =   2235
      Width           =   6315
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   645
      Index           =   4
      Left            =   180
      Top             =   2955
      Width           =   6315
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   5
      Left            =   180
      Top             =   765
      Width           =   6315
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strTitle, strCase As String
Private Sub cboReportType_Click()
    Select Case cboReportType.ListIndex
        Case 0
            txtVouType.Text = "JV"
        Case 1
            txtVouType.Text = "CP"
        Case 2
            txtVouType.Text = "CR"
        Case 3
            txtVouType.Text = "BP"
        Case 4
            txtVouType.Text = "BR"
        Case 5
            txtVouType.Text = "JV,CP,CR,BP,BR"
    End Select
End Sub

Private Sub cmdEXIT_Click()
    Unload Me
End Sub

Private Sub cmdPREVIEW_Click()
  Screen.MousePointer = vbHourglass
    
    'voucher date wise
    '*****************************************
    '*****************************************
    If Me.Option1.Value = True Then
        If dtst_dt.Value > dted_dt.Value Then
            MsgBox "Improper date range", vbCritical
            dtst_dt.SetFocus
            Exit Sub
        End If
        
            txtQueryPart.Text = "where vou_type in (" & Trim(Me.txtVouType.Text) & _
            ") and vou_date between ''" & Format(dtst_dt.Value, "yyyy-mm-dd") & "'' and ''" & _
            Format(dted_dt.Value, "yyyy-mm-dd") & "''"
            
            strTitle = "  from  " & Me.dtst_dt.Value & "  to  " & Me.dted_dt.Value
            
            CRViewer1.Show vbModal
    End If
    '*****************************************
    '*****************************************
    'Voucher no wise
    If Me.Option2.Value = True Then
        If Len(Trim(Me.txtVOU_NO.Text)) = 0 Then
            MsgBox "Voucher # required", vbCritical
            Me.txtVOU_NO.SetFocus
            Exit Sub
        End If
    
        CRViewer1.Show vbModal
    End If
    '*****************************************
    '*****************************************
    'all
    If Me.Option3.Value = True Then
    
        Me.txtQueryPart.Text = "where vou_type in (" & Trim(Me.txtVouType.Text) & ")"
        strTitle = ""
        CRViewer1.Show vbModal
        
    End If
    '*****************************************
    '*****************************************
End Sub

Private Sub dted_dt_CloseUp()
'    dted_dt.MaxDate = objectCompSetup.ed_dt
'    dted_dt.MinDate = objectCompSetup.st_dt
End Sub

Private Sub dted_dt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub dted_dt_LostFocus()
    dted_dt_CloseUp
End Sub

Private Sub dtst_dt_CloseUp()
'    dtst_dt.MaxDate = objectCompSetup.ed_dt
'    dtst_dt.MinDate = objectCompSetup.st_dt
End Sub

Private Sub dtst_dt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub dtst_dt_LostFocus()
    dtst_dt_CloseUp
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub Form_Load()
    rptMode = 2
    
'    objectCompSetup.Flush_Comp (strcn)
    
    dtst_dt.Value = Date
    dted_dt.Value = Date
    
    With Me.cboReportType
    .AddItem "Journal Voucher"
    .AddItem "Cash Payment Voucher"
    .AddItem "Cash Receipt Voucher"
    .AddItem "Bank Payment Voucher"
    .AddItem "Bank Receipt Voucher"
    .AddItem "All Transaction"
    End With
    
End Sub

Private Sub Option1_Click()
    dtst_dt.Enabled = True
    dted_dt.Enabled = True
    txtVOU_NO.Enabled = False
    strCase = "Vou_Date"
End Sub

Private Sub Option2_Click()
    dtst_dt.Enabled = False
    dted_dt.Enabled = False
    txtVOU_NO.Enabled = True
    strCase = "Vou_No"
    
End Sub

Private Sub Option3_Click()
    dtst_dt.Enabled = False
    dted_dt.Enabled = False
    txtVOU_NO.Enabled = False
    strCase = "All"
End Sub

