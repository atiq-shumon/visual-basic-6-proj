VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form form33 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Issue Search"
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
      TabIndex        =   13
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   3000
         Left            =   3150
         TabIndex        =   7
         Top             =   2700
         Visible         =   0   'False
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   5292
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   14737632
         ForeColor       =   8388608
         BackColorFixed  =   14737632
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483635
         BackColorBkg    =   16777215
         FocusRect       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4935
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   8705
         _Version        =   393216
         Rows            =   0
         Cols            =   3
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
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Caption         =   "Date && Issue Wise"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   2640
         TabIndex        =   9
         Top             =   180
         Width           =   2115
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Caption         =   "Date Wise"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Top             =   150
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.TextBox txtfields 
         Height          =   315
         Index           =   2
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   21
         Left            =   2550
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   3315
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   285
         Left            =   1260
         TabIndex        =   0
         Top             =   540
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   285
         Left            =   3540
         TabIndex        =   10
         Top             =   540
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To  Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2730
         TabIndex        =   12
         Top             =   540
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   11
         Top             =   540
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   5
         Top             =   990
         Visible         =   0   'False
         Width           =   750
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
         Caption         =   "Issue Search"
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
         Width           =   1890
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   1920
      TabIndex        =   16
      Top             =   6810
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Issues :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   60
      TabIndex        =   15
      Top             =   6810
      Width           =   1170
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
Attribute VB_Name = "form33"
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

Private Sub Form_Load()
  MaskEdBox1.Text = Format(Date, "dd/mm/yy")
  MaskEdBox2.Text = Format(Date, "dd/mm/yy")
  
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

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  
    If MaskEdBox1 <> "__/__/__" Then
            If Check_ValidDate(MaskEdBox1) = False Then
                MaskEdBox1.Text = "__/__/__"
                MaskEdBox1.SetFocus
                Exit Sub
            Else
               MaskEdBox2.SetFocus
            End If
      End If

End If
End Sub

Private Sub MaskEdBox2_GotFocus()
  MaskEdBox2.SelStart = 0
  MaskEdBox2.SelLength = Len(MaskEdBox1.Text)
  MaskEdBox2.SetFocus
End Sub

Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If MaskEdBox2 <> "__/__/__" Then
            If Check_ValidDate(MaskEdBox2) = False Then
                MaskEdBox2.Text = "__/__/__"
                MaskEdBox2.SetFocus
                Exit Sub
            Else
               If txtfields(2).Visible = True Then
                  txtfields(2).SetFocus
               Else
                 Call show_salesid(0, MaskEdBox1, MaskEdBox2, "")
                 MSFlexGrid1.Col = 0
                 MSFlexGrid1.Row = 0
                 MSFlexGrid1.SetFocus
               End If
               
            End If
      End If
    
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
   MSFlexGrid1_KeyPress (13)
End Sub
Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
On Error Resume Next
  If KeyAscii = 13 Then
   If Len(MSFlexGrid1.Text) > 0 Then
     frmIssue.txtfields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
   End If
     Unload Me
  End If
  frmIssue.txtfields(0).SetFocus
End Sub

Private Sub MSFlexGrid3_DblClick()
   If Len(Trim(MSFlexGrid3.Text)) <> 0 Then
       txtfields(2).Text = MSFlexGrid3.Text
       Call txtFields_LostFocus(2)
    Else
      txtfields(2).SetFocus
    End If
    MSFlexGrid3.Visible = False
End Sub

Private Sub MSFlexGrid3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     MSFlexGrid3_DblClick
  End If
End Sub

Private Sub Option1_Click(Index As Integer)
  Select Case Index
         Case 0
               Label2(5).Visible = False
               txtfields(2).Visible = False
               txtfields(21).Visible = False
               MaskEdBox1.SetFocus
         Case 1
               Label2(5).Visible = True
               txtfields(2).Visible = True
               txtfields(21).Visible = True
               MaskEdBox1.SetFocus
             
  End Select
End Sub
Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    Select Case Index
           Case 2
           txtFields_LostFocus (2)
    End Select
  End If
End Sub
Private Sub show_salesid(mode As Integer, date1 As Date, date2 As Date, var_customer As String)
 Dim obj_sales As New ADODB.Recordset
 Dim issueLike As String
 issueLike = "-" & CategoryCode & "-"
 If mode = 0 Then '''date to date
  format_grid
  i = 0
   Set obj_sales = objcom.Get_RS("SELECT m.IssueId,m.indent_no,s.type_name From IssueMain m ,item_issue_type s where to_number(m.IssueTYpe)= to_number(s.type_code)and m.issueid like '%" & issueLike & "%' and ( to_char(m.IssueDate,'dd-mon-yyyy') between '" & Format(MaskEdBox1.Text, "dd-mmm-yyyy") & "'  and '" & Format(MaskEdBox2.Text, "dd-mmm-yyyy") & "')  order by m.issueid desc", objmyCon)
   If Not obj_sales.EOF Then
    i = 0
'    and (to_date(to_char(m.IssueDate,'dd-mon-yyyy'),'dd-mon-yyyy') >= to_date('" & Format(MaskEdBox1.Text, "dd mmm yyyy") & "') and to_date(to_char(m.IssueDate,'dd-mon-yyyy'),'dd-mon-yyyy') <= to_date('" & Format(MaskEdBox2.Text, "dd mmm yyyy") & "')
    With MSFlexGrid1
         Do Until obj_sales.EOF
            .Rows = i + 1
            .TextMatrix(i, 0) = Trim(obj_sales(0))
            .ColAlignment(1) = 0
            .TextMatrix(i, 1) = "" & Trim(obj_sales(1))
            .TextMatrix(i, 2) = obj_sales(2)
            i = i + 1
            obj_sales.MoveNext
        Loop
       End With
    Label4.Caption = i
    MSFlexGrid1.Rows = 25 + i
Else
    Label4.Caption = 0
    MSFlexGrid1.Rows = 25 + i
 End If

  End If
   If mode = 1 Then '''date to date and customer wise
   format_grid
   i = 0
    Set obj_sales = objcom.Get_RS("SELECT m.IssueId,m.indent_no,s.type_name From IssueMain m ,item_issue_type s where to_number(m.IssueTYpe)= to_number(s.type_code)and m.issueid like '%" & issueLike & "%' and to_number(m.IssueTYpe)=to_number('" & Trim(txtfields(2).Text) & "') and  ( to_char(m.IssueDate,'dd-mon-yyyy') between '" & Format(MaskEdBox1.Text, "dd-mmm-yyyy") & "'  and '" & Format(MaskEdBox2.Text, "dd-mmm-yyyy") & "')  order by m.issueid desc", objmyCon)
    If Not obj_sales.EOF Then
    i = 0
    With MSFlexGrid1
         Do Until obj_sales.EOF
            .Rows = i + 1
            .TextMatrix(i, 0) = Trim(obj_sales(0))
            .ColAlignment(1) = 0
            .TextMatrix(i, 1) = "" & Trim(obj_sales(1))
            .TextMatrix(i, 2) = obj_sales(2)
            i = i + 1
            obj_sales.MoveNext
        Loop
       End With
    Label4.Caption = i
    MSFlexGrid1.Rows = 25 + i
Else
    Label4.Caption = 0
    MSFlexGrid1.Rows = 25 + i
 End If

  End If
   Set obj_sales = Nothing
   
End Sub
Private Sub format_grid()
  With MSFlexGrid1
    .Rows = 0
    .Cols = 3
    .ColWidth(0) = 1200
    .ColWidth(1) = 1800
    .ColWidth(2) = 2500
    .Rows = 20
End With

End Sub
Private Sub txtFields_LostFocus(Index As Integer)
  If Len(Trim(txtfields(2).Text)) = 0 Then Exit Sub

                  Set objRs = objcom.Get_RS("SELECT type_code,type_name from item_issue_type where type_code='" & txtfields(2) & "'", objmyCon)

                 If Not objRs.EOF Then
                     txtfields(2).Text = objRs(0).value
                     txtfields(21).Text = objRs(1).value
                  Else
                     MSFlexGrid3.Left = txtfields(2).Left
                     MSFlexGrid3.Top = txtfields(2).Top
                     MSFlexGrid3.TabIndex = txtfields(2).TabIndex + 1
                     Call getSupplier_code(Trim(txtfields(2).Text))
                    Exit Sub
                 End If
            Call show_salesid(1, MaskEdBox1, MaskEdBox2, txtfields(2))
            MSFlexGrid1.Col = 0
            MSFlexGrid1.Row = 0
            MSFlexGrid1.SetFocus
End Sub
Private Sub getSupplier_code(str As String)
 On Error GoTo err_loop
    MSFlexGrid3.Clear
    MSFlexGrid3.Rows = 0

    MSFlexGrid3.ColWidth(0) = "1200"
    MSFlexGrid3.ColAlignment(0) = 1

    MSFlexGrid3.ColWidth(1) = "10000"

  Set objRs = objcom.Get_RS("SELECT type_code,type_name from item_issue_type where upper(type_name) like upper('" & str & "%') order by type_code", objmyCon)

    If Not objRs.EOF Then
        Do Until objRs.EOF
            MSFlexGrid3.AddItem objRs(0) & vbTab & objRs(1)
            objRs.MoveNext
       Loop
    End If

    MSFlexGrid3.Visible = True
  MSFlexGrid3.SetFocus
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
End Sub
