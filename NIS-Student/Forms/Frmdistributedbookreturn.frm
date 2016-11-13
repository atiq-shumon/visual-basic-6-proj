VERSION 5.00
Begin VB.Form Frmdistributedbookreturn 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   30
      TabIndex        =   12
      Top             =   5370
      Width           =   7305
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H8000000C&
         Caption         =   "Save"
         Height          =   375
         Left            =   5160
         TabIndex        =   14
         ToolTipText     =   "Click to Save"
         Top             =   210
         Width           =   945
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H8000000C&
         Caption         =   "Close"
         Height          =   375
         Left            =   6150
         TabIndex        =   13
         ToolTipText     =   "Click to Exit"
         Top             =   210
         Width           =   945
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   5130
         Top             =   180
         Width           =   1995
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3585
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   7335
      Begin VB.ListBox List2 
         Height          =   2985
         Left            =   3570
         Style           =   1  'Checkbox
         TabIndex        =   10
         ToolTipText     =   "Select Received book"
         Top             =   510
         Width           =   3375
      End
      Begin VB.ListBox List1 
         Height          =   2985
         Left            =   90
         TabIndex        =   8
         ToolTipText     =   "Select Student"
         Top             =   510
         Width           =   3165
      End
      Begin VB.Line Line1 
         X1              =   3390
         X2              =   3390
         Y1              =   120
         Y2              =   3570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student List #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   150
         TabIndex        =   11
         Top             =   210
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recieved Book List #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   3600
         TabIndex        =   9
         Top             =   210
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   0
      TabIndex        =   1
      Top             =   810
      Width           =   7335
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Text            =   "Combo2"
         ToolTipText     =   "Select Class"
         Top             =   540
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Frmdistributedbookreturn.frx":0000
         Left            =   1080
         List            =   "Frmdistributedbookreturn.frx":000A
         TabIndex        =   4
         ToolTipText     =   "Select Shift"
         Top             =   180
         Width           =   1995
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Name "
         Height          =   195
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
      Height          =   795
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   7305
      TabIndex        =   0
      Top             =   0
      Width           =   7365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distributed Book Returning Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   1170
         TabIndex        =   2
         Top             =   180
         Width           =   4500
      End
   End
End
Attribute VB_Name = "Frmdistributedbookreturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSAVE_Click()
'Dim i
'If Len(Combo1) = 0 And Len(Combo2) = 0 Then Exit Sub
'
'If Len(List1.Selected(i)) = 0 Then
'    MsgBox "Please Select Student Name", vbCritical, "School Management System"
'    List1.SetFocus
'    Exit Sub
'End If
'
'If Len(List2.Selected(i)) = 0 Then
'    MsgBox "Please Select Subject Name", vbCritical, "School Management System"
'    List2.SetFocus
'    Exit Sub
'End If

Dim cmd As New adodb.Command
Dim con As New adodb.Connection
Dim rs As New adodb.Recordset
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "BookDisretInformation2"
For i = 0 To List2.ListCount - 1
        cmd(1) = Mid((List1.Text), 1, 15)
        cmd(2) = Combo1.Text
        cmd(3) = Mid(Combo2.Text, 1, 5)
        cmd(4) = Mid((List2.List(i)), 1, 5)
        If List2.Selected(i) = True Then
            cmd(5) = "R"
        Else
            cmd(5) = "Y"
        End If
        cmd(6) = Date
        cmd(7) = "DSL"
        cmd.Execute
Next
MsgBox "Save Successfully.", vbInformation, "Student Management System"
Call GetBooklist
End Sub

Private Sub Combo1_Click()
List1.Clear
Dim rs As New adodb.Recordset
Set rs = getdata("SELECT DISTINCT BookDistributionandReturnInfo.StudentId, StudentInfo.StudentName, BookDistributionandReturnInfo.DeliveryApproved " + _
"FROM  BookDistributionandReturnInfo INNER JOIN StudentInfo ON BookDistributionandReturnInfo.StudentId = StudentInfo.StudentID " + _
"WHERE  BookDistributionandReturnInfo.Shift =  '" & Combo1.Text & "' AND BookDistributionandReturnInfo.ClassId =  '" & Mid(Combo2.Text, 1, 5) & "'  AND BookDistributionandReturnInfo.DeliveryApproved = 'Y' ")
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
    Set rs = getdata("SELECT DISTINCT BookDistributionandReturnInfo.StudentId, StudentInfo.StudentName, BookDistributionandReturnInfo.DeliveryApproved " + _
    "FROM  BookDistributionandReturnInfo INNER JOIN StudentInfo ON BookDistributionandReturnInfo.StudentId = StudentInfo.StudentID " + _
    "WHERE  BookDistributionandReturnInfo.Shift =  '" & Combo1.Text & "' AND BookDistributionandReturnInfo.ClassId =  '" & Mid(Combo2.Text, 1, 5) & "'  AND BookDistributionandReturnInfo.DeliveryApproved = 'Y' ")
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
List1.Clear
Dim rs As New adodb.Recordset
Set rs = getdata("SELECT DISTINCT BookDistributionandReturnInfo.StudentId, StudentInfo.StudentName, BookDistributionandReturnInfo.DeliveryApproved " + _
"FROM  BookDistributionandReturnInfo INNER JOIN StudentInfo ON BookDistributionandReturnInfo.StudentId = StudentInfo.StudentID " + _
"WHERE  BookDistributionandReturnInfo.Shift =  '" & Combo1.Text & "' AND BookDistributionandReturnInfo.ClassId =  '" & Mid(Combo2.Text, 1, 5) & "'  AND BookDistributionandReturnInfo.DeliveryApproved = 'Y' ")
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
    Set rs = getdata("SELECT DISTINCT BookDistributionandReturnInfo.StudentId, StudentInfo.StudentName, BookDistributionandReturnInfo.DeliveryApproved " + _
    "FROM  BookDistributionandReturnInfo INNER JOIN StudentInfo ON BookDistributionandReturnInfo.StudentId = StudentInfo.StudentID " + _
    "WHERE  BookDistributionandReturnInfo.Shift =  '" & Combo1.Text & "' AND BookDistributionandReturnInfo.ClassId =  '" & Mid(Combo2.Text, 1, 5) & "'  AND BookDistributionandReturnInfo.DeliveryApproved = 'Y' ")
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
Set rs1 = getdata("SELECT     BookDistributionandReturnInfo.SubjectId, SubjectInfo.SubjectDsc " + _
    "FROM BookDistributionandReturnInfo INNER JOIN SubjectInfo ON BookDistributionandReturnInfo.SubjectId = SubjectInfo.SubjectID " + _
    "WHERE (BookDistributionandReturnInfo.StudentId = '" & Mid((List1.Text), 1, 15) & "') AND (SubjectInfo.ClassID = '" & Mid((Combo2.Text), 1, 5) & "' AND DeliveryApproved='Y')")
If Not rs1.EOF Then
    List2.Clear
    Do Until rs1.EOF
         List2.AddItem rs1!SubjectID + " - " + rs1!SubjectDsc
         rs1.MoveNext
    Loop
End If
For i = 0 To List2.ListCount - 1
    Set rs1 = getdata("select SubjectId from BookDistributionandReturnInfo where classid='" & Mid((Combo2.Text), 1, 5) & "'and shift='" & (Combo1.Text) & "'and StudentId='" & Mid((List1.Text), 1, 15) & "'and  BookRecievedback='R' and subjectId='" & Mid((List2.List(i)), 1, 5) & "'")
    If Not rs1.EOF Then
        Do Until rs1.EOF
            If Mid((List2.List(i)), 1, 5) = rs1!SubjectID Then
                List2.Selected(i) = True
            End If
           rs1.MoveNext
        Loop
    End If
Next
End Function
Private Sub List1_Click()
List2.Clear
Call GetBooklist
End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim rs As New adodb.Recordset
For i = 0 To List2.ListCount - 1
    If List2.Selected(i) = False Then
        Set rs = getdata("select ReturnApproved, BookRecievedback    from BookDistributionandReturnInfo where studentid='" & Mid((List1.Text), 1, 15) & "' and subjectid='" & Mid((List2.List(i)), 1, 5) & "'and classid='" & Mid((Combo2.Text), 1, 5) & "'and shift='" & (Combo1.Text) & "'")
            If Not rs.EOF Then
                If rs!ReturnApproved = "Y" And rs!BookRecievedback = "R" Then
                   MsgBox " '" & List2.Text & "'Cannot be removed .It is Return Approved.", vbInformation, App.Title
                   List2.Selected(i) = True
                   Exit Sub
                 Else
                    List2.Selected(i) = False
                 End If
            End If
    End If
Next
End Sub
