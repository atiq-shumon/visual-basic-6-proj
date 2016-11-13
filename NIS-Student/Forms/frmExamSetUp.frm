VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmExamSetUp 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8220
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   -30
      TabIndex        =   15
      Top             =   4830
      Width           =   8685
      Begin VB.CommandButton CmdEdit 
         BackColor       =   &H8000000C&
         Caption         =   "Edit"
         Height          =   375
         Left            =   5130
         TabIndex        =   20
         ToolTipText     =   "Click to Save"
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H8000000C&
         Caption         =   "Close"
         Height          =   375
         Left            =   7170
         TabIndex        =   8
         ToolTipText     =   "Click to Exit"
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H8000000C&
         Caption         =   "Delete"
         Height          =   375
         Left            =   6150
         TabIndex        =   7
         ToolTipText     =   "Click to Delete"
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H8000000C&
         Caption         =   "Save"
         Height          =   375
         Left            =   4140
         TabIndex        =   5
         ToolTipText     =   "Click to Save"
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton cmdnew 
         BackColor       =   &H8000000C&
         Caption         =   "New"
         Height          =   375
         Left            =   3120
         MaskColor       =   &H8000000C&
         TabIndex        =   6
         ToolTipText     =   "Click to insert new information"
         Top             =   210
         Width           =   975
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   3060
         Top             =   180
         Width           =   5115
      End
   End
   Begin VB.Frame Frame6 
      ForeColor       =   &H00C00000&
      Height          =   1995
      Left            =   0
      TabIndex        =   11
      Top             =   720
      Width           =   8235
      Begin VB.TextBox txtfields 
         BackColor       =   &H00CEF0F7&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3480
         TabIndex        =   1
         Top             =   330
         Width           =   4365
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   1185
      End
      Begin VB.TextBox txtfields 
         Height          =   465
         Index           =   4
         Left            =   960
         MaxLength       =   80
         TabIndex        =   4
         ToolTipText     =   "Insert Short Note"
         Top             =   1380
         Width           =   6885
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   1
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   3
         ToolTipText     =   "Insert Exam Type Name"
         Top             =   900
         Width           =   4365
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   2
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term Title"
         Height          =   195
         Index           =   2
         Left            =   2610
         TabIndex        =   17
         Top             =   330
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term code"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   1350
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type Title"
         Height          =   195
         Left            =   2610
         TabIndex        =   13
         Top             =   900
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type code"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   930
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   8595
      TabIndex        =   9
      Top             =   0
      Width           =   8655
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   10
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam  Type Set Up"
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
         Left            =   2580
         TabIndex        =   18
         Top             =   150
         Width           =   2160
      End
      Begin VB.Image Image1 
         Height          =   870
         Left            =   -30
         Picture         =   "frmExamSetUp.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   8235
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2115
      Left            =   0
      TabIndex        =   19
      Top             =   2730
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   3731
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
Attribute VB_Name = "frmExamSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con

 If Len(txtfields(3)) = 0 Then
        MsgBox "Type Code Mandatory..Press on NEW button to generate such Code", vbCritical, App.Title
        cmdnew.SetFocus
        Exit Sub
    End If
     If Len(txtfields(1)) = 0 Then
        MsgBox "Type Title Mandatory...", vbCritical, App.Title
        txtfields(1).SetFocus
        Exit Sub
    End If
   
    
If (MsgBox("Are You sure to delete ?", vbYesNo + vbCritical) = vbYes) Then
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Exam_Type_Info_Save"
    cmd(1) = "D"
    cmd(2) = Trim(Combo1.Text)
    cmd(3) = Trim(txtfields(3))
    cmd(4) = Trim(txtfields(1))
    cmd(5) = Trim(txtfields(2))
    cmd(6) = soft_user
    cmd(7) = Format(Date, "DD MMM YYYY")
    cmd.Execute
    MsgBox "Delete Successfully Exam Type Information.", vbInformation, App.Title
    
    For i = 2 To 3
     txtfields(i) = ""
    Next
    Call ShowFlexData
Else
    Exit Sub
End If
End Sub
Private Sub CmdEdit_Click()
  Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con

 If Len(txtfields(3)) = 0 Then
        MsgBox "Type Code Mandatory..Press on NEW button to generate such Code", vbCritical, App.Title
        cmdnew.SetFocus
        Exit Sub
    End If
     If Len(txtfields(1)) = 0 Then
        MsgBox "Type Title Mandatory...", vbCritical, App.Title
        txtfields(1).SetFocus
        Exit Sub
    End If
   
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Exam_Type_Info_Save"
    cmd(1) = "U"
    cmd(2) = Trim(Combo1.Text)
    cmd(3) = Trim(txtfields(3))
    cmd(4) = Trim(txtfields(1))
    cmd(5) = Trim(txtfields(2))
    cmd(6) = soft_user
    cmd(7) = Format(Date, "DD MMM YYYY")
    cmd.Execute
    MsgBox "Updated Successfully. ", vbInformation, App.Title
    
    For i = 2 To 3
     txtfields(i) = ""
    Next
    Call ShowFlexData

End Sub

Private Sub cmdnew_Click()

If Len(Combo1.Text) = 0 Then
   MsgBox " Please Select a term ..", vbInformation, App.Title
   Combo1.SetFocus
   Exit Sub
End If
   
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con

    Set rs = getdata("select max(cast(Exam_code as int))+ 1 from Exam_setup where Group_code='" & Trim(Combo1.Text) & "'")
    If Not rs.EOF Then
        txtfields(3) = IIf(IsNull(rs(0)) = True, "01", Format(rs(0), "00"))
    Else
        txtfields(3) = "01"
    End If
    txtfields(1) = ""
    txtfields(4) = ""
    txtfields(1).SetFocus



End Sub

Private Sub Combo1_Click()
  show_title
  ShowFlexData
End Sub

Private Sub Command1_Click()
 
End Sub

Private Sub Form_Load()
Call load_term
With MSFlexGrid1
    .Rows = 1
    .Cols = 3
    .Col = 0: .Text = " Term Code#"
    .Col = 1: .Text = "Exam Type code"
    .Col = 2: .Text = " Exam Type Title "
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 7000
   
End With
ShowFlexData
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys (Chr(9))
   End If
End Sub




Private Sub MSFlexGrid1_SelChange()
  MSFlexGrid1_Click
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
'        Case 2
'            txtfields(3).SetFocus
        Case 3
'            txtfields(4).SetFocus
'        Case 4
'            cmdsave.SetFocus
    End Select
End If
End Sub

Private Sub cmdSAVE_Click()
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con

    If Len(txtfields(3)) = 0 Then
        MsgBox "Type Code Mandatory..Press on NEW button to generate such Code", vbCritical, App.Title
        cmdnew.SetFocus
        Exit Sub
    End If
     If Len(txtfields(1)) = 0 Then
        MsgBox "Type Title Mandatory...", vbCritical, App.Title
        txtfields(1).SetFocus
        Exit Sub
    End If
   
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Exam_Type_Info_Save"
    cmd(1) = "S"
    cmd(2) = Trim(Combo1.Text)
    cmd(3) = Trim(txtfields(3))
    cmd(4) = Trim(txtfields(1))
    cmd(5) = Trim(txtfields(2))
    cmd(6) = soft_user
    cmd(7) = Format(Date, "DD MMM YYYY")
   
    cmd.Execute
    MsgBox "Saved Successfully.", vbInformation, "Student Management System"
    cmdnew.SetFocus
  
    Call ShowFlexData

End Sub
Private Sub ShowFlexData()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT   Group_code ,Exam_code ,Exam_title from Exam_SETUP where group_code='" & Trim(Combo1) & "'")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                
                .TextMatrix(i, 0) = "" & rs!Group_code
                .TextMatrix(i, 1) = "" & rs!Exam_code
                .TextMatrix(i, 2) = "" & rs!Exam_title
                

            rs.MoveNext
           i = i + 1
        Loop
    End With
 Else
     MSFlexGrid1.Rows = 1

 End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub MSFlexGrid1_Click()
On Error GoTo errdes
'Combo1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfields(3) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
txtfields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)

Exit Sub
errdes:
 MsgBox Err.Description, vbInformation, App.Title

End Sub


Private Sub load_term()
Dim rs As New ADODB.Recordset

Set rs = getdata("SELECT ETypeID from ExamTypeInfo")

Do Until rs.EOF
    Combo1.AddItem rs(0)
    rs.MoveNext
Loop
End Sub


Private Sub show_title()
Dim rs As New ADODB.Recordset

Set rs = getdata("SELECT ETypeName from ExamTypeInfo where ETypeID='" & Trim(Combo1.Text) & "'")

If Not rs.EOF Then
  txtfields(2).Text = rs(0)
End If
End Sub


