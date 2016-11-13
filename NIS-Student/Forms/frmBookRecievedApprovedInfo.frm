VERSION 5.00
Begin VB.Form frmBookdistributedApprovedInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Exit"
      Height          =   435
      Left            =   6780
      TabIndex        =   11
      Top             =   5070
      Width           =   945
   End
   Begin VB.CommandButton cmdApproved 
      BackColor       =   &H8000000C&
      Caption         =   "Approved"
      Height          =   435
      Left            =   5850
      TabIndex        =   10
      Top             =   5070
      Width           =   945
   End
   Begin VB.Frame Frame3 
      Height          =   3135
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   7725
      Begin VB.ListBox List2 
         Height          =   2760
         Left            =   4020
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   180
         Width           =   3705
      End
      Begin VB.ListBox List1 
         Height          =   2790
         Left            =   90
         TabIndex        =   9
         Top             =   180
         Width           =   3615
      End
      Begin VB.Line Line1 
         X1              =   3780
         X2              =   3780
         Y1              =   120
         Y2              =   3150
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
         TabIndex        =   5
         Text            =   "Combo2"
         Top             =   510
         Width           =   2595
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmBookRecievedApprovedInfo.frx":0000
         Left            =   1080
         List            =   "frmBookRecievedApprovedInfo.frx":000A
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Name"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   570
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Left            =   120
         TabIndex        =   6
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
         Caption         =   "Book Recieved Approved"
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
         Left            =   2310
         TabIndex        =   3
         Top             =   240
         Width           =   2925
      End
   End
End
Attribute VB_Name = "frmBookdistributedApprovedInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApproved_Click()
Dim cmd As New Command

Dim con As New Connection
con.Open GConnString
    

    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BookDisretInformation1"
    For i = 0 To List2.ListCount - 1
        
            cmd(1) = Mid((List1.Text), 1, 15)

            cmd(2) = Mid((List2.List(i)), 1, 5)
            If List2.Selected(i) = True Then
                cmd(3) = "Y"
            Else
                cmd(3) = "N"
            End If
            cmd(4) = "Lia"
            cmd(5) = Date
        
            cmd.Execute
        
    Next
    MsgBox "Save Successfully.", vbInformation, "Student Management System"
    


End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub
Public Function GetBooklist()

Dim rs1 As New Recordset
Dim rs2 As New Recordset

Set rs1 = GetData("SELECT     BookDistributionandReturnInfo.SubjectId, SubjectInfo.SubjectDsc " + _
    "FROM BookDistributionandReturnInfo INNER JOIN SubjectInfo ON BookDistributionandReturnInfo.SubjectId = SubjectInfo.SubjectID " + _
    "WHERE (BookDistributionandReturnInfo.StudentId = '" & Mid((List1.Text), 1, 15) & "') AND (SubjectInfo.ClassID = '" & Mid((Combo2.Text), 1, 5) & "' AND BookRecievedback='Y')")
If Not rs1.EOF Then
Do Until rs1.EOF
 
     List2.AddItem rs1!SubjectID + " - " + rs1!SubjectDsc
     
     rs1.MoveNext
Loop

End If
For i = 0 To List2.ListCount - 1
Set rs1 = GetData("select SubjectId from BookDistributionandReturnInfo where StudentId='" & Mid((List1.Text), 1, 15) & "'and  BookRecievedback='Y' and subjectId='" & Mid((List2.List(i)), 1, 5) & "'and DeliveryApproved='y'")

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

Private Sub Combo2_GotFocus()
Dim rs1 As New Recordset
Set rs1 = GetData("select ClassId,ClassName from ClassInfo")
If Not rs1.EOF Then
Do Until rs1.EOF
       Combo2.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop
    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0

End If
End Sub

Private Sub Form_Load()
Dim rs As New Recordset
Set rs = GetData("select StudentId,StudentName from StudentInfo")
If Not rs.EOF Then
Do Until rs.EOF
        List1.AddItem rs(0) + " - " + rs(1)
        rs.MoveNext
    Loop
    If List1.ListCount > 0 Then List1.ListIndex = 0

End If

Dim rs1 As New Recordset
Set rs1 = GetData("select ClassId,ClassName from ClassInfo")
If Not rs1.EOF Then
Do Until rs1.EOF
       Combo2.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop
    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0

End If
End Sub
Public Function GetData(SQLString As String) As Recordset
Dim cmd As New Command
Dim con As New Connection
Dim rs As New Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = SQLString
 Set rs = cmd.Execute
Set GetData = rs
End Function

Private Sub List1_Click()
List2.Clear

Call GetBooklist
End Sub
