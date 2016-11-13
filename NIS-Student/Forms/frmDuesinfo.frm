VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDuesinfo 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   -30
      TabIndex        =   5
      Top             =   5460
      Width           =   8085
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H8000000C&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5010
         TabIndex        =   10
         ToolTipText     =   "Click to Update information"
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H8000000C&
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
         Height          =   405
         Left            =   6990
         TabIndex        =   9
         ToolTipText     =   "Click to Exit"
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H8000000C&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6000
         TabIndex        =   8
         ToolTipText     =   "Click to Delete"
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H8000000C&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3990
         TabIndex        =   7
         ToolTipText     =   "Click to Save"
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton cmdnew 
         BackColor       =   &H8000000C&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         TabIndex        =   6
         ToolTipText     =   "Click to insert new information"
         Top             =   210
         Width           =   975
      End
      Begin VB.Shape Shape1 
         Height          =   465
         Left            =   2970
         Top             =   180
         Width           =   5025
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   795
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   7965
      TabIndex        =   3
      Top             =   0
      Width           =   8025
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dues Amount Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEF0F7&
         Height          =   285
         Left            =   2430
         TabIndex        =   11
         Top             =   150
         Width           =   2925
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   -90
         Picture         =   "frmDuesinfo.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   705
      Left            =   -60
      TabIndex        =   2
      Top             =   810
      Width           =   8115
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   2970
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   0
         Top             =   150
         Width           =   5055
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1590
         TabIndex        =   1
         Top             =   150
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student  ID && Name"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   210
         Width           =   1410
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3885
      Left            =   0
      TabIndex        =   12
      Top             =   1500
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   6853
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   15005934
      BackColorSel    =   -2147483624
      ForeColorSel    =   16711680
      BackColorBkg    =   15724265
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmDuesinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbSubjectType_Change()
 
      cmbSubjectType.Text = "Compulsory"
        
End Sub



Private Sub cmbSubjectType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Combo2.SetFocus
   End If
End Sub

Private Sub cmdDelete_Click()
   If Len(txtfields(0)) <> 0 And Len(txtfields(1)) <> 0 Then
       Dim rs1 As New ADODB.Recordset
       Set rs1 = getdata("select M_code from Subject_Info_sub where M_code ='" & Trim(txtfields(0).Text) & "'")
      If Not rs1.EOF Then
         MsgBox "Data exists in Subject information Sub..First delete there", vbInformation, cmp
         Exit Sub
      End If
      If (MsgBox("Are You sure to delete ?", vbYesNo + vbCritical) = vbYes) Then
            Dim rs As New ADODB.Recordset
            Dim cmd As New ADODB.Command
            Dim con As New ADODB.connection
            con.Open GConnString
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SubjectInformation_main"
            cmd(1) = "D"
            cmd(2) = Format(Trim(txtfields(0)), "00000")
            cmd(3) = Trim(txtfields(1))
            cmd(4) = Trim(Combo2.Text)
            cmd(5) = Trim(cmbSubjectType.Text)
            cmd(6) = Trim(soft_user)
            cmd(7) = Format(Date, "dd mmm yyyy")
            cmd.Execute
            MsgBox "Deleted successfully.", vbInformation, cmp
            cmdnew.SetFocus
            Call ShowFlexData
   End If
   
 End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
'If Len(Combo1.Text) = 0 Then Exit Sub
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
'con.Open connectionstring.GConnString
con.Open GConnString
cmd.ActiveConnection = con
Set rs = getdata("select max(isnull(cast(M_code as int),0))+1 from SubjectInfomain")
If Not rs.EOF Then
    txtfields(0) = IIf(IsNull(rs(0)) = True, "00001", Format(rs(0), "00000"))
Else
    txtfields(0) = "00001"
End If

For i = 1 To 1
    txtfields(i) = ""
Next
Combo2.Text = " "
cmbSubjectType.Text = " "

txtfields(1).SetFocus

End Sub

Private Sub cmdSAVE_Click()
If Len(txtfields(0)) = 0 Then
    
    MsgBox "Please Enter subject Id.", vbInformation, App.Title
    cmdnew.SetFocus
    Exit Sub
End If
If Len(txtfields(1)) = 0 Then
    
    MsgBox "Please Enter subject Name.", vbInformation, App.Title
    txtfields(1).SetFocus
    Exit Sub
End If
If Combo2.Text = " " Then
    MsgBox "Please Enter Subject Unit.", vbInformation, App.Title
    Combo2.SetFocus
    Exit Sub
End If
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "SubjectInformation_main"
cmd(1) = "S"
cmd(2) = Format(Trim(txtfields(0)), "00000")
cmd(3) = Trim(txtfields(1))
cmd(4) = Trim(Combo2.Text)
cmd(5) = Trim(cmbSubjectType.Text)
cmd(6) = Trim(soft_user)
cmd(7) = Format(Date, "dd mmm yyyy")
cmd.Execute
MsgBox "Saved successfully.", vbInformation, cmp
cmdnew.SetFocus
Call ShowFlexData
End Sub

Private Sub cmdUpdate_Click()
  If Len(txtfields(0)) = 0 Then
    
    MsgBox "Please Enter subject Id.", vbInformation, App.Title
    cmdnew.SetFocus
    Exit Sub
End If
If Len(txtfields(1)) = 0 Then
    
    MsgBox "Please Enter subject Name.", vbInformation, App.Title
    txtfields(1).SetFocus
    Exit Sub
End If
If Combo2.Text = " " Then
    MsgBox "Please Enter Subject Unit.", vbInformation, App.Title
    Combo2.SetFocus
    Exit Sub
End If
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "SubjectInformation_main"
cmd(1) = "U"
cmd(2) = Format(Trim(txtfields(0)), "00000")
cmd(3) = Trim(txtfields(1))
cmd(4) = Trim(Combo2.Text)
cmd(5) = Trim(cmbSubjectType.Text)
cmd(6) = Trim(soft_user)
cmd(7) = Format(Date, "dd mmm yyyy")
cmd.Execute
MsgBox "Updated successfully.", vbInformation, cmp
cmdnew.SetFocus
Call ShowFlexData

End Sub

Private Sub Combo1_Click()
Call ShowFlexData
Call cmdnew_Click
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdnew.SetFocus
End If
End Sub

Private Sub Combo1_LostFocus()
Call ShowFlexData

End Sub

Private Sub Combo2_Change()
   Combo2 = "1"
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdsave.SetFocus
End If
End Sub



Private Sub Command1_Click()
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
  txtfields(0).Text = frmCollection_info.txtfields(0)
  txtfields(1).Text = frmCollection_info.txtfields(2)
  
With MSFlexGrid1
    .Rows = 1
    .Cols = 10
    .Col = 0: .Text = "  Fee Code #"
    .Col = 1: .Text = "Fee Title"
'    .Col = 2: .Text = "     Subject Unit "
'    .Col = 3: .Text = "     Compolsary/Optional "
    .ColWidth(0) = 900
    .ColWidth(1) = 2500 ''4500
    .ColWidth(2) = 1000
'    .ColWidth(3) = 1300
'
End With
ShowFlexData
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo errdes
txtfields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
Combo2 = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
cmbSubjectType = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)

Exit Sub
errdes:
  MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub txtfields_Change(Index As Integer)
            Select Case Index
                   Case 2
                         If Not IsNumeric(txtfields(2).Text) Then
                               txtfields(2) = ""
                         End If
            End Select
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
'If KeyAscii = 13 Then
'  Select Case Index
'
'    Case 1
'       cmbSubjectType.SetFocus
'    Case 2
'        Combo2.SetFocus
'
'End Select
'End If
End Sub
Public Function getdata(SQLString As String) As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = SQLString
 Set rs = cmd.Execute
Set getdata = rs
End Function

Private Sub txtfields_LostFocus(Index As Integer)
txtfields(0) = Format(txtfields(0), "00000")
Dim rs As New ADODB.Recordset

Select Case Index
    Case 0
        If Len(Trim(txtfields(0))) = 0 Then Exit Sub
      
            txtfields(0) = Format(txtfields(0), "00000")
          
            Set rs = getdata("SELECT * from SubjectInfo WHERE (SubjectID = '" & txtfields(0) & "') and ClassID= '" & Combo1.Text & "'")
                 If Not rs.EOF Then
                        Combo1.Text = rs!classId
                        txtfields(1) = rs!SubjectDsc
                        txtfields(2) = rs!totalmarks
                        Combo2.Text = rs!Subjectunit
               
                End If
                
    Case 2
        Dim SubMarks As Double
        If Len(Trim(txtfields(2))) = 0 Then Exit Sub
        If IsNumeric(txtfields(2)) = False Then
            MsgBox "Please Enter Numeric Value.", vbInformation, App.Title
            txtfields(2) = ""
            txtfields(2).SetFocus
            Exit Sub
        End If
End Select
End Sub
Private Sub ShowFlexData()
Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("SELECT a.Fee_Code,b.Fee_title ,a.Fee_amt,a.NoOfTimes From Fee_setup a,fee_info b where a.fee_code=b.fee_code and FeesStatus=1 and Class_id='" & frmCollection_info.Combo1 & "'")
If Not rs1.EOF Then

    i = 1
    With MSFlexGrid1
        Do Until rs1.EOF
            MSFlexGrid1.Rows = i + 1
                .TextMatrix(i, 0) = rs1!Fee_Code
                .TextMatrix(i, 1) = rs1!fee_title
                .TextMatrix(i, 2) = rs1!Fee_amt
                .TextMatrix(i, 3) = rs1!NoOfTimes
                .ColAlignment(4) = 0
                .TextMatrix(i, 4) = Month(Date)
                .TextMatrix(i, 5) = 12 / rs1!NoOfTimes
                .TextMatrix(i, 6) = (CInt(.TextMatrix(i, 4)) / CInt(.TextMatrix(i, 5)))
'                .TextMatrix(i, 2) = rs1!Subjectunit
'                 .TextMatrix(i, 3) = rs1!SubjectType
                i = i + 1
            rs1.MoveNext
        Loop
        
    End With
 Else
        MSFlexGrid1.Rows = 1
        
 End If

End Sub
