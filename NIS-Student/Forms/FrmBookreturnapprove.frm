VERSION 5.00
Begin VB.Form FrmBookreturnapprove 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdApproved 
      BackColor       =   &H8000000C&
      Caption         =   "Approved"
      Height          =   435
      Left            =   4590
      TabIndex        =   14
      ToolTipText     =   "Click to Approve"
      Top             =   5640
      Width           =   945
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   435
      Left            =   5580
      TabIndex        =   13
      ToolTipText     =   "Click to Exit"
      Top             =   5640
      Width           =   945
   End
   Begin VB.Frame Frame3 
      Height          =   945
      Left            =   0
      TabIndex        =   7
      Top             =   780
      Width           =   6525
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Select Class"
         Top             =   510
         Width           =   2745
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmBookreturnapprove.frx":0000
         Left            =   1230
         List            =   "FrmBookreturnapprove.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Select Shift"
         Top             =   150
         Width           =   1995
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Name"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   540
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3945
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   6525
      Begin VB.ListBox List2 
         Height          =   3210
         Left            =   3390
         Style           =   1  'Checkbox
         TabIndex        =   4
         ToolTipText     =   "Select Book for Approve"
         Top             =   540
         Width           =   2985
      End
      Begin VB.ListBox List1 
         Height          =   3180
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Select Student"
         Top             =   540
         Width           =   2985
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recieved Book Approval #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   3390
         TabIndex        =   6
         Top             =   180
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student List #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   1125
      End
      Begin VB.Line Line1 
         X1              =   3240
         X2              =   3240
         Y1              =   90
         Y2              =   3900
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   795
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   6465
      TabIndex        =   0
      Top             =   0
      Width           =   6525
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   1
         Top             =   930
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Returned Book (By Student) Approval Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   450
         TabIndex        =   10
         Top             =   180
         Width           =   5265
      End
   End
End
Attribute VB_Name = "FrmBookreturnapprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApproved_Click()
If Len(Combo1) = 0 And Len(Combo2) = 0 Then Exit Sub
If Len(Combo1) = 0 Then
    MsgBox "Please Enter Shift Name.", vbCritical, "Shool Management System"
    Combo1.SetFocus
    Exit Sub
End If
If Len(Combo2) = 0 Then
    MsgBox "Please Enter Class Name.", vbCritical, "Shool Management System"
    Combo2.SetFocus
    Exit Sub
End If
If Len(List1) = 0 Then
    MsgBox "Please Enter Student Name.", vbCritical, "Shool Management System"
    List1.SetFocus
    Exit Sub
End If
If List2.Selected(i) = False Then
    MsgBox "Please Enter Subject name", vbCritical, "School Managemnet System"
    List2.SetFocus
    Exit Sub
End If
Dim cmd As New adodb.Command
Dim con As New adodb.Connection
Dim rs As New adodb.Recordset
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "BookDisretInformation3"
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
    cmd(7) = Date
    cmd(6) = "DSL"
    cmd.Execute
Next
    MsgBox "Save Successfully.", vbInformation, "Student Management System"
Call GetBooklist
End Sub


Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Combo1_Click()
List1.Clear
Dim rs As New adodb.Recordset
Set rs = getdata(" SELECT DISTINCT BookDistributionandReturnInfo.StudentId, StudentInfo.StudentName " + _
"FROM         BookDistributionandReturnInfo INNER JOIN StudentInfo ON BookDistributionandReturnInfo.StudentId = StudentInfo.StudentID " + _
"WHERE     (BookDistributionandReturnInfo.Shift = '" & Combo1.Text & "') AND (BookDistributionandReturnInfo.ClassId =  '" & Mid(Combo2.Text, 1, 5) & "' ) AND " + _
"(BookDistributionandReturnInfo.BookRecievedBack = 'R')")
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
    Set rs = getdata(" SELECT DISTINCT BookDistributionandReturnInfo.StudentId, StudentInfo.StudentName " + _
    "FROM         BookDistributionandReturnInfo INNER JOIN StudentInfo ON BookDistributionandReturnInfo.StudentId = StudentInfo.StudentID " + _
    "WHERE     (BookDistributionandReturnInfo.Shift = '" & Combo1.Text & "') AND (BookDistributionandReturnInfo.ClassId =  '" & Mid(Combo2.Text, 1, 5) & "' ) AND " + _
    "(BookDistributionandReturnInfo.BookRecievedBack = 'R')")
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
Set rs = getdata(" SELECT DISTINCT BookDistributionandReturnInfo.StudentId, StudentInfo.StudentName " + _
"FROM         BookDistributionandReturnInfo INNER JOIN StudentInfo ON BookDistributionandReturnInfo.StudentId = StudentInfo.StudentID " + _
"WHERE     (BookDistributionandReturnInfo.Shift = '" & Combo1.Text & "') AND (BookDistributionandReturnInfo.ClassId =  '" & Mid(Combo2.Text, 1, 5) & "' ) AND " + _
"(BookDistributionandReturnInfo.BookRecievedBack = 'R')")
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
    Set rs = getdata(" SELECT DISTINCT BookDistributionandReturnInfo.StudentId, StudentInfo.StudentName " + _
    "FROM         BookDistributionandReturnInfo INNER JOIN StudentInfo ON BookDistributionandReturnInfo.StudentId = StudentInfo.StudentID " + _
    "WHERE     (BookDistributionandReturnInfo.Shift = '" & Combo1.Text & "') AND (BookDistributionandReturnInfo.ClassId =  '" & Mid(Combo2.Text, 1, 5) & "' ) AND " + _
    "(BookDistributionandReturnInfo.BookRecievedBack = 'R')")
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
    Combo2.AddItem (" ")
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
Private Sub List1_Click()
List2.Clear

Call GetBooklist
End Sub

Public Function GetBooklist()

Dim rs1 As New adodb.Recordset
Set rs1 = getdata("SELECT  BookDistributionandReturnInfo.SubjectId, SubjectInfo.SubjectDsc " + _
"FROM BookDistributionandReturnInfo INNER JOIN SubjectInfo ON BookDistributionandReturnInfo.ClassId = SubjectInfo.ClassID AND " + _
"BookDistributionandReturnInfo.SubjectID = SubjectInfo.SubjectID WHERE     (BookDistributionandReturnInfo.Shift = '" & Combo1.Text & "') AND (BookDistributionandReturnInfo.ClassId = '" & Mid(Combo2.Text, 1, 5) & "') AND  " + _
"BookDistributionandReturnInfo.BookRecievedBack = 'R' AND BookDistributionandReturnInfo.StudentId = '" & Mid(Trim(List1.Text), 1, 16) & "'")
If Not rs1.EOF Then
    List2.Clear
    Do Until rs1.EOF
        List2.AddItem rs1!SubjectID + " - " + rs1!SubjectDsc
        rs1.MoveNext
    Loop
End If
For i = 0 To List2.ListCount - 1
Set rs1 = getdata("select SubjectId from BookDistributionandReturnInfo where classid='" & Mid((Combo2.Text), 1, 5) & "' and Shift='" & Combo1.Text & "' and StudentId='" & Mid((List1.Text), 1, 16) & "'and  BookRecievedback='R' and subjectId='" & Mid((List2.List(i)), 1, 5) & "'and ReturnApproved      ='y'")
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

