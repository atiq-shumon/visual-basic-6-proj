VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmExamTypeSetUp 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   -30
      TabIndex        =   9
      Top             =   4170
      Width           =   8025
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H8000000C&
         Caption         =   "Close"
         Height          =   375
         Left            =   6810
         TabIndex        =   13
         ToolTipText     =   "Click to Exit"
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H8000000C&
         Caption         =   "Delete"
         Height          =   375
         Left            =   5790
         TabIndex        =   12
         ToolTipText     =   "Click to Delete"
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H8000000C&
         Caption         =   "Save"
         Height          =   375
         Left            =   4770
         TabIndex        =   11
         ToolTipText     =   "Click to Save"
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton cmdnew 
         BackColor       =   &H8000000C&
         Caption         =   "New"
         Height          =   375
         Left            =   3750
         MaskColor       =   &H8000000C&
         TabIndex        =   10
         ToolTipText     =   "Click to insert new information"
         Top             =   210
         Width           =   975
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   3720
         Top             =   180
         Width           =   4095
      End
   End
   Begin VB.Frame Frame6 
      ForeColor       =   &H00C00000&
      Height          =   1185
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   7845
      Begin VB.TextBox txtfields 
         Height          =   465
         Index           =   4
         Left            =   840
         MaxLength       =   80
         TabIndex        =   1
         ToolTipText     =   "Insert Short Note"
         Top             =   630
         Width           =   6885
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   3
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   0
         ToolTipText     =   "Insert Exam Type Name"
         Top             =   240
         Width           =   4155
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   570
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term Title"
         Height          =   195
         Left            =   2700
         TabIndex        =   7
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term ID"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   570
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   7785
      TabIndex        =   2
      Top             =   0
      Width           =   7845
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   3
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Term Set Up"
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
         Left            =   2610
         TabIndex        =   14
         Top             =   180
         Width           =   2115
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   -30
         Picture         =   "frmExamTypeSetUp.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   7905
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2295
      Left            =   0
      TabIndex        =   15
      Top             =   1890
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   4048
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
Attribute VB_Name = "frmExamTypeSetUp"
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
If Len(txtfields(2)) = 0 Then
        MsgBox "Please Enter an valid Exam Term Code...or choose form list below", vbCritical, App.Title
        cmdnew.SetFocus
        Exit Sub
    End If
 Set rs = getdata("select term_code from SubjectMarksDistribution where term_code='" & Trim(txtfields(2)) & "'")
 If Not rs.EOF Then
    MsgBox "Already Used...You can't delete", vbInformation, cmp
    Exit Sub
  End If
    
   
If (MsgBox("Are You sure to delete ?", vbYesNo + vbCritical) = vbYes) Then
    cmd.CommandType = adCmdText
    cmd.CommandText = "Delete from ExamTypeInfo  where (EtypeID = '" & Trim(txtfields(2)) & "') "
    cmd.Execute
    MsgBox "Delete Successfully Exam Type Information.", vbInformation, App.Title
    
    For i = 2 To 4
     txtfields(i) = ""
    Next
    Call ShowFlexData
Else
    Exit Sub
End If
End Sub


Private Sub cmdnew_Click()
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con

    Set rs = getdata("select max(cast(EtypeID as int))+ 1 from ExamTypeInfo")
    If Not rs.EOF Then
        txtfields(2) = IIf(IsNull(rs(0)) = True, "01", Format(rs(0), "00"))
    Else
        txtfields(0) = "01"
    End If
    txtfields(3) = ""
    txtfields(4) = ""
    txtfields(3).SetFocus



End Sub


Private Sub Form_Load()
With MSFlexGrid1
    .Rows = 1
    .Cols = 3
    .Col = 0: .Text = " Examination ID   #"
    .Col = 1: .Text = "Examination Type Name   "
    .Col = 2: .Text = " Note  "
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 5000
    .ColWidth(2) = 5000
   
End With
ShowFlexData
End Sub

Private Sub MSFlexGrid1_SelChange()
  MSFlexGrid1_Click
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
        Case 2
            txtfields(3).SetFocus
        Case 3
            txtfields(4).SetFocus
        Case 4
            cmdsave.SetFocus
    End Select
End If
End Sub
Private Sub cmdSAVE_Click()
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con

    If Len(txtfields(2)) = 0 Then
        MsgBox "Please Enter Exam Term Code..Press on NEW button to generate such Code", vbCritical, App.Title
        cmdnew.SetFocus
        Exit Sub
    End If

    If Len(txtfields(3)) = 0 Then
        MsgBox "Please Enter Exam Type Name.", vbCritical, App.Title
        txtfields(3).SetFocus
        Exit Sub
    End If
   
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "ETypeInfo"
    cmd(1) = txtfields(2)
    cmd(2) = Trim(txtfields(3))
    cmd(3) = Trim(txtfields(4))
    cmd(4) = soft_user
    cmd(5) = Date
   
    cmd.Execute
    MsgBox "Save Successfully.", vbInformation, App.Title
    cmdnew.SetFocus
  
    Call ShowFlexData

End Sub

Private Sub ShowFlexData()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT   EtypeID ,ETypeName ,Note from ExamTypeInfo")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                
                .TextMatrix(i, 0) = "" & rs!EtypeID
                .TextMatrix(i, 1) = "" & rs!ETypeName
                .TextMatrix(i, 2) = "" & rs!Note
                

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
txtfields(2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfields(3) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
txtfields(4) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)

Exit Sub
errdes:
'MsgBox err.Description, vbInformation, App.Title

End Sub



