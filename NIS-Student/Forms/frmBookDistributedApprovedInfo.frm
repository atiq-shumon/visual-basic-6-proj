VERSION 5.00
Begin VB.Form frmBookdistributedApprovedInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   435
      Left            =   6780
      TabIndex        =   10
      ToolTipText     =   "Click to Exit"
      Top             =   5220
      Width           =   945
   End
   Begin VB.CommandButton cmdApproved 
      BackColor       =   &H8000000C&
      Caption         =   "Approved"
      Height          =   435
      Left            =   5790
      TabIndex        =   9
      ToolTipText     =   "Click to Approve"
      Top             =   5220
      Width           =   945
   End
   Begin VB.Frame Frame3 
      Height          =   3255
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   7725
      Begin VB.ListBox List2 
         Height          =   2760
         Left            =   3960
         Style           =   1  'Checkbox
         TabIndex        =   11
         ToolTipText     =   "Select Books for approve"
         Top             =   420
         Width           =   3705
      End
      Begin VB.ListBox List1 
         Height          =   2790
         Left            =   30
         TabIndex        =   8
         ToolTipText     =   "Select Student"
         Top             =   420
         Width           =   3615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distributed Book List"
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
         Left            =   3960
         TabIndex        =   13
         Top             =   150
         Width           =   1785
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student List"
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
         Left            =   90
         TabIndex        =   12
         Top             =   150
         Width           =   1035
      End
      Begin VB.Line Line1 
         X1              =   3780
         X2              =   3780
         Y1              =   120
         Y2              =   3240
      End
   End
   Begin VB.Frame Frame2 
      Height          =   945
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   7725
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Select Class"
         Top             =   510
         Width           =   2595
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmBookDistributedApprovedInfo.frx":0000
         Left            =   1080
         List            =   "frmBookDistributedApprovedInfo.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Select Shift"
         Top             =   180
         Width           =   1875
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Name"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   570
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   7665
      TabIndex        =   0
      Top             =   0
      Width           =   7725
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   1
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Distributed Approved"
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
         Left            =   1710
         TabIndex        =   14
         Top             =   240
         Width           =   3120
      End
      Begin VB.Image Image1 
         Height          =   1230
         Left            =   -60
         Picture         =   "frmBookDistributedApprovedInfo.frx":0035
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   7785
      End
   End
End
Attribute VB_Name = "frmBookdistributedApprovedInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApproved_Click()
If Len(Combo1.Text) = 0 And Len(Combo2) = 0 Then Exit Sub
If Len(List1) = 0 Then
MsgBox "Please Enter student Name.", vbCritical, "School Management System"
List1.SetFocus
Exit Sub
End If
If List2.Selected(i) = False Then
MsgBox "Please Select subject Name.", vbCritical, "School Management System"
List2.SetFocus
Exit Sub
End If

    Dim cmd As New adodb.Command
    Dim con As New adodb.Connection
    con.Open GConnString
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BookDisretInformation1"
    For i = 0 To List2.ListCount - 1
            cmd(1) = Trim(Mid((List1.Text), 1, 15))
            cmd(2) = Trim(Combo1.Text)
            cmd(3) = Trim(Mid((Combo2.Text), 1, 5))
            cmd(4) = Trim(Mid((List2.List(i)), 1, 5))
            If List2.Selected(i) = True Then
                cmd(5) = "Y"
            Else
                cmd(5) = "N"
            End If
            cmd(7) = "DSL"
            cmd(6) = Date
            cmd.Execute
    Next
    MsgBox "Save Successfully.", vbInformation, "Student Management System"
    Exit Sub
 End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Public Function GetBooklist()
Dim rs1 As New adodb.Recordset
Dim rs2 As New adodb.Recordset
Set rs1 = getdata("SELECT     BookDistributionandReturnInfo.SubjectId, SubjectInfo.SubjectDsc " + _
    "FROM BookDistributionandReturnInfo INNER JOIN SubjectInfo ON BookDistributionandReturnInfo.SubjectId = SubjectInfo.SubjectID " + _
    "WHERE (BookDistributionandReturnInfo.StudentId = '" & Mid((List1.Text), 1, 15) & "') AND (BookDistributionandReturnInfo.BookRecievedBack <> 'N') AND " + _
    "(SubjectInfo.ClassID = '" & Mid((Combo2.Text), 1, 5) & "') AND (BookDistributionandReturnInfo.ClassId = '" & Mid((Combo2.Text), 1, 5) & "')")
If Not rs1.EOF Then
List2.Clear
Do Until rs1.EOF
     List2.AddItem rs1!SubjectID + " - " + rs1!SubjectDsc
      rs1.MoveNext
Loop
End If
For i = 0 To List2.ListCount - 1
Set rs1 = getdata("select SubjectId from BookDistributionandReturnInfo where ClassId='" & Mid((Combo2.Text), 1, 5) & "'and shift='" & (Combo1.Text) & "'and StudentId='" & Mid((List1.Text), 1, 15) & "' and subjectId='" & Mid((List2.List(i)), 1, 5) & "'and DeliveryApproved='Y'and BookRecievedback <> 'N'")
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
Private Sub Combo1_Click()
List1.Clear
Dim rs As New adodb.Recordset
Set rs = getdata("SELECT DISTINCT BookDistributionandReturnInfo.StudentId, StudentInfo.StudentName " + _
"FROM BookDistributionandReturnInfo INNER JOIN StudentInfo ON BookDistributionandReturnInfo.StudentId = StudentInfo.StudentID " + _
"WHERE     BookDistributionandReturnInfo.Shift = '" & Combo1.Text & "' AND BookDistributionandReturnInfo.ClassId = '" & Mid(Combo2.Text, 1, 5) & "' ")
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
Set rs = getdata("SELECT DISTINCT BookDistributionandReturnInfo.StudentId, StudentInfo.StudentName " + _
"FROM BookDistributionandReturnInfo INNER JOIN StudentInfo ON BookDistributionandReturnInfo.StudentId = StudentInfo.StudentID " + _
"WHERE     BookDistributionandReturnInfo.Shift = '" & Combo1.Text & "' AND BookDistributionandReturnInfo.ClassId = '" & Mid(Combo2.Text, 1, 5) & "' ")
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
Set rs = getdata("SELECT DISTINCT BookDistributionandReturnInfo.StudentId, StudentInfo.StudentName " + _
"FROM BookDistributionandReturnInfo INNER JOIN StudentInfo ON BookDistributionandReturnInfo.StudentId = StudentInfo.StudentID " + _
"WHERE     BookDistributionandReturnInfo.Shift = '" & Combo1.Text & "' AND BookDistributionandReturnInfo.ClassId = '" & Mid(Combo2.Text, 1, 5) & "' ")
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
Set rs = getdata("SELECT DISTINCT BookDistributionandReturnInfo.StudentId, StudentInfo.StudentName " + _
"FROM BookDistributionandReturnInfo INNER JOIN StudentInfo ON BookDistributionandReturnInfo.StudentId = StudentInfo.StudentID " + _
"WHERE     BookDistributionandReturnInfo.Shift = '" & Combo1.Text & "' AND BookDistributionandReturnInfo.ClassId = '" & Mid(Combo2.Text, 1, 5) & "' ")
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

Private Sub List1_Click()
List2.Clear
Call GetBooklist
End Sub
