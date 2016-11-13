VERSION 5.00
Begin VB.Form Frmbookdistribution 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   435
      Left            =   5640
      TabIndex        =   12
      ToolTipText     =   "Click to save"
      Top             =   5820
      Width           =   945
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   435
      Left            =   6630
      TabIndex        =   11
      ToolTipText     =   "Click to Exit"
      Top             =   5820
      Width           =   945
   End
   Begin VB.Frame Frame2 
      Height          =   4035
      Left            =   0
      TabIndex        =   6
      Top             =   1770
      Width           =   7575
      Begin VB.ListBox List1 
         Height          =   3375
         Left            =   90
         TabIndex        =   8
         ToolTipText     =   "Select Student"
         Top             =   540
         Width           =   3495
      End
      Begin VB.ListBox List2 
         Height          =   3435
         Left            =   3750
         Style           =   1  'Checkbox
         TabIndex        =   7
         ToolTipText     =   "Select Subject"
         Top             =   540
         Width           =   3705
      End
      Begin VB.Line Line1 
         X1              =   3660
         X2              =   3660
         Y1              =   120
         Y2              =   4050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject List"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3930
         TabIndex        =   10
         Top             =   210
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student List"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   180
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   810
      Width           =   7575
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Select Class"
         Top             =   540
         Width           =   2745
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Frmbookdisandreturn.frx":0000
         Left            =   1200
         List            =   "Frmbookdisandreturn.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select Shift"
         Top             =   180
         Width           =   2025
      End
      Begin VB.Label Label3 
         Caption         =   "Class Name"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   -30
      Width           =   7575
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Distribution Information"
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
         Left            =   1890
         TabIndex        =   13
         Top             =   240
         Width           =   3375
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -30
         Picture         =   "Frmbookdisandreturn.frx":0035
         Stretch         =   -1  'True
         Top             =   30
         Width           =   7665
      End
   End
End
Attribute VB_Name = "Frmbookdistribution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim List1Click As Boolean
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSAVE_Click()
Dim cmd As New adodb.Command
If Len(Combo1) = 0 And Len(Combo2) = 0 Then Exit Sub
If Len(Combo1.Text) = 0 Then
    MsgBox "Select Shift Name.", vbInformation, App.Title
    Combo1.SetFocus
    Exit Sub
End If
If Len(Combo2.Text) = 0 Then
    MsgBox "Select Class Name.", vbInformation, App.Title
    Combo2.SetFocus
    Exit Sub
End If
If Len(List1.Text) = 0 Then
    MsgBox "Select Student Name.", vbInformation, App.Title
    List1.SetFocus
    Exit Sub
End If
If List2.Selected(i) = False Then
    MsgBox "Select Subjectt Name.", vbInformation, App.Title
    List2.SetFocus
    Exit Sub
End If
If Len(Combo1) <> 0 And Len(Combo2) <> 0 Then
Dim con As New adodb.Connection
Dim rs As New adodb.Recordset
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "BookDisretInformation"
For i = 0 To List2.ListCount - 1
        cmd(1) = Mid((List1.Text), 1, 15)
        cmd(2) = Combo1.Text
        cmd(3) = Mid(Combo2.Text, 1, 5)
        cmd(4) = Mid((List2.List(i)), 1, 5)
    
        If List2.Selected(i) = True Then
            cmd(5) = "Y"
        Else
            cmd(5) = "N"
    
        End If
        cmd(6) = Date
        cmd(7) = "DSL"
    
        cmd.Execute
Next
    MsgBox "Save Successfully.", vbInformation, "Student Management System"
    Call GetBooklist
Else
    Exit Sub
End If
End Sub

Private Sub Combo1_Click()
List1.Clear
Dim rs As New adodb.Recordset
Set rs = getdata("SELECT     StudentEvaluation.StudentID, StudentInfo.StudentName " + _
                "FROM StudentEvaluation INNER JOIN StudentInfo ON StudentEvaluation.StudentID = StudentInfo.StudentID where StudentEvaluation.ClassId ='" & Mid(Combo2, 1, 5) & "' and StudentEvaluation.Active = 'Y' and StudentEvaluation.Activeclass = 'Y' and StudentEvaluation.Shift = '" & Combo1.Text & "' ")
    If Not rs.EOF Then
    Do Until rs.EOF
            List1.AddItem rs(0) + " - " + rs(1)
            rs.MoveNext
    Loop
        If List1.ListCount > 0 Then List1.ListIndex = 0
  
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
List1.Clear
Dim rs As New adodb.Recordset
Set rs = getdata("SELECT     StudentEvaluation.StudentID, StudentInfo.StudentName " + _
"FROM StudentEvaluation INNER JOIN StudentInfo ON StudentEvaluation.StudentID = StudentInfo.StudentID where StudentEvaluation.ClassId ='" & Mid(Combo2, 1, 5) & "' and StudentEvaluation.Active = 'Y' and StudentEvaluation.Activeclass = 'Y' and StudentEvaluation.Shift = '" & Combo1.Text & "' ")
    If Not rs.EOF Then
    Do Until rs.EOF
            List1.AddItem rs(0) + " - " + rs(1)
            rs.MoveNext
        Loop
        If List1.ListCount > 0 Then List1.ListIndex = 0
   
    End If
End If
End Sub

Private Sub Combo2_Click()
List2.Clear
Dim rs1 As New adodb.Recordset
Set rs1 = getdata("select subjectId,SubjectDsc from SubjectInfo where classId ='" & Mid((Combo2.Text), 1, 5) & "'")
If Not rs1.EOF Then
Do Until rs1.EOF
        List2.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop
    If List2.ListCount > 0 Then List2.ListIndex = 0
End If
Dim rs As New adodb.Recordset
List1.Clear
Set rs = getdata("SELECT     StudentEvaluation.StudentID, StudentInfo.StudentName " + _
"FROM StudentEvaluation INNER JOIN StudentInfo ON StudentEvaluation.StudentID = StudentInfo.StudentID where StudentEvaluation.ClassId ='" & Mid(Combo2, 1, 5) & "' and StudentEvaluation.Active = 'Y' and StudentEvaluation.Activeclass = 'Y' and StudentEvaluation.shift = '" & Combo1.Text & "'")
If Not rs.EOF Then
Do Until rs.EOF
        List1.AddItem rs(0) + " - " + rs(1)
        rs.MoveNext
    Loop
    If List1.ListCount > 0 Then List1.ListIndex = 0

End If
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    List1.Clear
    Dim rs As New adodb.Recordset
    Set rs = getdata("SELECT     StudentEvaluation.StudentID, StudentInfo.StudentName " + _
    "FROM StudentEvaluation INNER JOIN StudentInfo ON StudentEvaluation.StudentID = StudentInfo.StudentID where StudentEvaluation.ClassId ='" & Mid(Combo2, 1, 5) & "' and StudentEvaluation.Active = 'Y' and StudentEvaluation.Activeclass = 'Y' and StudentEvaluation.Shift = '" & Combo1.Text & "' ")
        If Not rs.EOF Then
            Do Until rs.EOF
                    List1.AddItem rs(0) + " - " + rs(1)
                    rs.MoveNext
            Loop
                If List1.ListCount > 0 Then List1.ListIndex = 0
        
        End If
End If
End Sub

Private Sub Form_Load()
Dim rs1 As New adodb.Recordset
Set rs1 = getdata("select ClassId,ClassName from ClassInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
        Combo2.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop
'    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0

End If

End Sub
Public Function getdata(SQLString As String) As adodb.Recordset
Dim cmd As New adodb.Command
Dim con As New adodb.Connection
Dim rs As New adodb.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
cmd.CommandType = adCmdText
cmd.CommandText = SQLString
Set rs = cmd.Execute
Set getdata = rs
End Function
Public Function GetBooklist()
Dim rs1 As New adodb.Recordset
Dim rs2 As New adodb.Recordset
Set rs1 = getdata("select subjectId from BookDistributionandReturnInfo  where classid='" & Mid((Combo2.Text), 1, 5) & "'and shift='" & (Combo1.Text) & "' and  StudentId='" & Mid((List1.Text), 1, 15) & "' and BookRecievedback<>'N'")
If Not rs1.EOF Then
    Do Until rs1.EOF
        For i = 0 To List2.ListCount - 1
            If Mid((List2.List(i)), 1, 5) = rs1!SubjectID Then
                 List2.Selected(i) = True
        
            End If
        Next
        rs1.MoveNext
    Loop
        
End If

End Function

Private Sub List1_Click()
For i = 0 To List2.ListCount - 1
    List2.Selected(i) = False
Next

GetBooklist

End Sub



Private Sub List2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim rs As New adodb.Recordset
For i = 0 To List2.ListCount - 1
    If List2.Selected(i) = False Then
        Set rs = getdata("select DeliveryApproved, BookRecievedback    from BookDistributionandReturnInfo where classid='" & Mid((Combo2.Text), 1, 5) & "'and shift='" & (Combo1.Text) & "'and studentid ='" & Mid((List1.Text), 1, 15) & "' and subjectid='" & Mid((List2.List(i)), 1, 5) & "'")
            If Not rs.EOF Then
                If rs!DeliveryApproved = "Y" And rs!BookRecievedback <> "N" Then
                       MsgBox " '" & List2.Text & "'Cannot be removed .It was approved.", vbInformation, App.Title
                       List2.Selected(i) = True
                       Exit Sub
                 Else
                        List2.Selected(i) = False
                 End If
            End If
    End If
Next
End Sub
