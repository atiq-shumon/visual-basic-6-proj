VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form form34 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5925
   ForeColor       =   &H000000C0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H8000000C&
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4860
      TabIndex        =   5
      ToolTipText     =   "Click to  Edit Information"
      Top             =   6750
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   6315
      Left            =   -30
      TabIndex        =   2
      Top             =   390
      Width           =   6405
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5835
         Left            =   60
         TabIndex        =   1
         Top             =   450
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   10292
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorSel    =   12632319
         ForeColorSel    =   12582912
         BackColorBkg    =   -2147483624
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   21
         Left            =   60
         TabIndex        =   0
         Top             =   120
         Width           =   5865
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000001&
      FillColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   -30
      ScaleHeight     =   435
      ScaleWidth      =   12045
      TabIndex        =   3
      Top             =   -30
      Width           =   12105
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Search Screen"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   345
         Index           =   2
         Left            =   1920
         TabIndex        =   4
         Top             =   -30
         Width           =   2175
      End
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mnuRFS 
         Caption         =   "Refresh"
      End
      Begin VB.Menu Sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDL 
         Caption         =   "Delete"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuUpdateSerial 
      Caption         =   "Update Serial"
      Visible         =   0   'False
      Begin VB.Menu mnuUSE 
         Caption         =   "Update Serial"
      End
   End
End
Attribute VB_Name = "form34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objcom As New DSLComFram.clsCommon
Dim objRs As New ADODB.Recordset
Private Sub CmdEdit_Click()
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
   If Len(List1.Text) > 0 Then
     form14.txtfields(0) = Trim(Get_Description(List1.Text))
   End If
     Unload Me
  End If
  form14.txtfields(0).SetFocus
End Sub

Private Sub MaskEdBox1_GotFocus()
  MaskEdBox1.SelStart = 0
  MaskEdBox1.SelLength = Len(MaskEdBox1.Text)
  MaskEdBox1.SetFocus

End Sub



Private Sub MaskEdBox2_GotFocus()
  MaskEdBox2.SelStart = 0
  MaskEdBox2.SelLength = Len(MaskEdBox1.Text)
  MaskEdBox2.SetFocus
End Sub

Private Sub Form_Load()
    Call format_grid
    If IssueSearchMode = 0 Or IssueSearchMode = 1 Then
     load_issueto (IssueSearchMode)
    Else
      load_issueto (5)
    End If
End Sub
Private Sub load_issueto(Index As Integer)
 
 Select Case Index
        Case 0
           Set objRs = objcom.Get_RS("SELECT emp_id,emp_nm  from payroll.emp_info where upper(emp_id) not like upper('M%') or upper(emp_id) not like upper('c%') order by emp_id", objmyCon)
         
             With MSFlexGrid1
                             Do Until objRs.EOF
                                .Rows = i + 1
                                .TextMatrix(i, 0) = Trim(objRs(0))
                                .TextMatrix(i, 1) = "" & Trim(objRs(1))
                               
                                i = i + 1
                                objRs.MoveNext
                            Loop
            End With
        Case 1
           Set objRs = objcom.Get_RS("SELECT emp_id,emp_nm  from payroll.emp_info where upper(emp_id)  like upper('M%')  order by emp_id", objmyCon)
          With MSFlexGrid1
                             Do Until objRs.EOF
                                .Rows = i + 1
                                .TextMatrix(i, 0) = Trim(objRs(0))
                                .TextMatrix(i, 1) = "" & Trim(objRs(1))
                               
                                i = i + 1
                                objRs.MoveNext
                            Loop
            End With
       Case 2
           Set objRs = objcom.Get_RS("SELECT bed_no,bed_type,BED_EXT_COL  from hospital_billing.bed_info  order by bed_type", objmyCon)
           With MSFlexGrid1
                             Do Until objRs.EOF
                                .Rows = i + 1
                                .TextMatrix(i, 0) = Trim(objRs(0))
                                .TextMatrix(i, 1) = "" & Trim(objRs(1))
                               
                                i = i + 1
                                objRs.MoveNext
                            Loop
            End With
                
'       Case 2
'           Set objRs = objcom.Get_RS("SELECT emp_id,emp_nm  from payroll.emp_info where upper(emp_id)  like upper('c%')  order by emp_id", objmyCon)
'           CboSupplier.Clear
'           If Not objRs.EOF Then
'              objRs.MoveFirst
'              Do Until objRs.EOF
'                 CboSupplier.AddItem Trim(objRs(1)) + "~" + Trim(objRs(0))
'                 objRs.MoveNext
'              Loop
'           End If
'
         Case 3
           Set objRs = objcom.Get_RS("SELECT refer_code,doc_dept  from hospital_billing.doctor_info  order by refer_code", objmyCon)
           With MSFlexGrid1
                             Do Until objRs.EOF
                                .Rows = i + 1
                                .TextMatrix(i, 0) = Trim(objRs(0))
                                .TextMatrix(i, 1) = "" & Trim(objRs(1))
                               
                                i = i + 1
                                objRs.MoveNext
                            Loop
            End With
           
           
      Case 4
'           Set objRs = objcom.Get_RS("SELECT refer_code,doc_dept  from hospital_billing.doctor_info  order by refer_code", objmyCon)
           CboSupplier.Clear
'           If Not objRs.EOF Then
'              objRs.MoveFirst
'              Do Until objRs.EOF
'                 CboSupplier.AddItem Trim(objRs(1)) + "~" + Trim(objRs(0))
'                 objRs.MoveNext
'              Loop
'           End If
     Case 5
           Set objRs = objcom.Get_RS("SELECT emp_id,emp_nm  from payroll.emp_info order by emp_id", objmyCon)
         
             With MSFlexGrid1
                             Do Until objRs.EOF
                                .Rows = i + 1
                                .TextMatrix(i, 0) = Trim(objRs(0))
                                .TextMatrix(i, 1) = "" & Trim(objRs(1))
                               
                                i = i + 1
                                objRs.MoveNext
                            Loop
            End With
    Case 6
           Set objRs = objcom.Get_RS("SELECT emp_id,emp_nm  from payroll.emp_info where upper(emp_nm)  like upper('%" & Trim(txtfields(21)) & "%') order by emp_id", objmyCon)
             format_grid
             MSFlexGrid1.Clear
                      
                      With MSFlexGrid1
                             Do Until objRs.EOF
                                .Rows = i + 1
                                .TextMatrix(i, 0) = Trim(objRs(0))
                                .TextMatrix(i, 1) = "" & Trim(objRs(1))
                                i = i + 1
                                objRs.MoveNext
                            Loop
            End With
           
  End Select
  
End Sub
Private Sub MSFlexGrid1_DblClick()
     MSFlexGrid1_KeyPress (13)
End Sub
Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
On Error Resume Next
  If KeyAscii = 13 Then
   If Len(MSFlexGrid1.Text) > 0 Then
     frmIssue.txtfields(7) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
     frmIssue.txtfields(7).SetFocus
   End If
     Unload Me
  End If
  frmIssue.txtfields(7).SetFocus
End Sub

Private Sub format_grid()
  With MSFlexGrid1
   .Clear
    .Rows = 0
    .Cols = 2
    .ColAlignment(0) = 0
    .ColWidth(0) = 900
    .ColWidth(1) = 6000
   
    .Rows = 1000
End With

End Sub

Private Sub txtfields_Change(Index As Integer)
  load_issueto (6)
End Sub

