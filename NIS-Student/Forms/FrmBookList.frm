VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBookList 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   645
      Left            =   0
      TabIndex        =   22
      Top             =   5610
      Width           =   9975
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H8000000C&
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   26
         ToolTipText     =   "Click to Save"
         Top             =   180
         Width           =   1005
      End
      Begin VB.CommandButton cmdDelete 
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
         Height          =   375
         Left            =   7740
         TabIndex        =   25
         ToolTipText     =   "Click to Save"
         Top             =   180
         Width           =   1005
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
         Height          =   375
         Left            =   4650
         TabIndex        =   8
         ToolTipText     =   "Click to insert new information"
         Top             =   180
         Width           =   1005
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
         Height          =   375
         Left            =   5700
         TabIndex        =   7
         ToolTipText     =   "Click to Save"
         Top             =   180
         Width           =   1005
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
         Height          =   375
         Left            =   8790
         TabIndex        =   9
         ToolTipText     =   "Click to exit"
         Top             =   180
         Width           =   1005
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   4590
         Top             =   150
         Width           =   5265
      End
   End
   Begin VB.Frame Frame5 
      Height          =   3975
      Left            =   3420
      TabIndex        =   17
      Top             =   1620
      Width           =   6555
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   2
         Left            =   1290
         MaxLength       =   100
         TabIndex        =   5
         ToolTipText     =   "Insert Writer Name"
         Top             =   1140
         Width           =   5175
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2385
         Left            =   120
         TabIndex        =   21
         Top             =   1530
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4207
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   15397576
         ForeColorSel    =   12582912
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   0
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   3
         ToolTipText     =   "Insert Book Name"
         Top             =   450
         Width           =   5175
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   1
         Left            =   1290
         MaxLength       =   100
         TabIndex        =   4
         ToolTipText     =   "Insert Writer Name"
         Top             =   780
         Width           =   5175
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Publisher"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   23
         Top             =   1110
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Name"
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Writter"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   19
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book List #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   150
         TabIndex        =   18
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3975
      Left            =   0
      TabIndex        =   15
      Top             =   1620
      Width           =   3405
      Begin VB.ListBox List1 
         Height          =   3375
         Left            =   90
         TabIndex        =   2
         ToolTipText     =   "Select Subject"
         Top             =   510
         Width           =   3255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject List #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   180
         Width           =   1110
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   9915
      TabIndex        =   13
      Top             =   0
      Width           =   9975
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   14
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Book  List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEF0F7&
         Height          =   330
         Left            =   3870
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   1125
         Left            =   -150
         Picture         =   "FrmBookList.frx":0000
         Stretch         =   -1  'True
         Top             =   -60
         Width           =   10095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   9975
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmBookList.frx":CEA5
         Left            =   450
         List            =   "FrmBookList.frx":CEA7
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select Class"
         Top             =   240
         Width           =   2985
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   7050
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select Year"
         Top             =   210
         Width           =   2565
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   30
         Left            =   3450
         TabIndex        =   10
         Top             =   660
         Width           =   6315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class "
         Height          =   195
         Left            =   30
         TabIndex        =   12
         Top             =   270
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Education  Year"
         Height          =   195
         Left            =   5400
         TabIndex        =   11
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "FrmBookList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
  Dim cmd As New ADODB.Command
If Len(Combo1) = 0 And Len(Combo2) = 0 Then Exit Sub
If Len(Combo1.Text) = 0 Then
    MsgBox "Select Class Name.", vbInformation, App.Title
    Combo1.SetFocus
    Exit Sub
End If
If Len(Combo2.Text) = 0 Then
    MsgBox "Select Education Year.", vbInformation, App.Title
    Combo2.SetFocus
    Exit Sub
End If
If Len(List1.Text) = 0 Then
    MsgBox "Select Subject.", vbInformation, App.Title
    List1.SetFocus
    Exit Sub
End If
If Len(txtFields(0).Text) = 0 Then
    MsgBox "Enter Book Name.", vbInformation, App.Title
    txtFields(0).SetFocus
    Exit Sub
End If
If Len(txtFields(0).Text) = 0 Then
    MsgBox "Enter Writter Name.", vbInformation, App.Title
    txtFields(1).SetFocus
    Exit Sub
End If
If MsgBox("Are you sure to Delete ?", vbQuestion + vbYesNo + vbDefaultButton1, cmp) = vbYes Then
     Dim con As New ADODB.connection
     Dim rs As New ADODB.Recordset
     con.Open GConnString
     cmd.ActiveConnection = con
     cmd.CommandType = adCmdStoredProc
     cmd.CommandText = "BookList1"
     cmd(1) = 3
     cmd(2) = Mid((Combo1.Text), 1, 5)
     cmd(3) = Trim(Combo2.Text)
     cmd(4) = Mid((List1.Text), 1, 5)
     cmd(5) = Trim(txtFields(0))
     cmd(6) = Trim(txtFields(1))
     cmd(7) = Trim(txtFields(2))
     cmd(8) = soft_user
     cmd(9) = Date
     cmd.Execute
     MsgBox "Deleted Successfully.", vbInformation, "Student Management System"
     ShowFlexData
End If


End Sub

Private Sub CmdEdit_Click()
  Dim cmd As New ADODB.Command
If Len(Combo1) = 0 And Len(Combo2) = 0 Then Exit Sub
If Len(Combo1.Text) = 0 Then
    MsgBox "Select Class Name.", vbInformation, App.Title
    Combo1.SetFocus
    Exit Sub
End If
If Len(Combo2.Text) = 0 Then
    MsgBox "Select Education Year.", vbInformation, App.Title
    Combo2.SetFocus
    Exit Sub
End If
If Len(List1.Text) = 0 Then
    MsgBox "Select Subject.", vbInformation, App.Title
    List1.SetFocus
    Exit Sub
End If
If Len(txtFields(0).Text) = 0 Then
    MsgBox "Enter Book Name.", vbInformation, App.Title
    txtFields(0).SetFocus
    Exit Sub
End If
If Len(txtFields(0).Text) = 0 Then
    MsgBox "Enter Writter Name.", vbInformation, App.Title
    txtFields(1).SetFocus
    Exit Sub
End If

Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "BookList1"
cmd(1) = 2
cmd(2) = Mid((Combo1.Text), 1, 5)
cmd(3) = Combo2.Text
cmd(4) = Mid((List1.Text), 1, 5)
cmd(5) = Trim(txtFields(0))
cmd(6) = Trim(txtFields(1))
cmd(7) = Trim(txtFields(2))
cmd(8) = Trim(soft_user)
cmd(9) = Date

cmd.Execute
MsgBox "Updated Successfully.", vbInformation, "Student Management System"
ShowFlexData


End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
txtFields(0) = ""
txtFields(1) = ""
txtFields(0).SetFocus
End Sub

Private Sub cmdSAVE_Click()
Dim cmd As New ADODB.Command
If Len(Combo1) = 0 And Len(Combo2) = 0 Then Exit Sub
If Len(Combo1.Text) = 0 Then
    MsgBox "Select Class Name.", vbInformation, App.Title
    Combo1.SetFocus
    Exit Sub
End If
If Len(Combo2.Text) = 0 Then
    MsgBox "Select Education Year.", vbInformation, App.Title
    Combo2.SetFocus
    Exit Sub
End If
If Len(List1.Text) = 0 Then
    MsgBox "Select Subject.", vbInformation, App.Title
    List1.SetFocus
    Exit Sub
End If
If Len(txtFields(0).Text) = 0 Then
    MsgBox "Enter Book Name.", vbInformation, App.Title
    txtFields(0).SetFocus
    Exit Sub
End If
If Len(txtFields(0).Text) = 0 Then
    MsgBox "Enter Writter Name.", vbInformation, App.Title
    txtFields(1).SetFocus
    Exit Sub
End If

Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "BookList1"
cmd(1) = 1
cmd(2) = Mid((Combo1.Text), 1, 5)
cmd(3) = Combo2.Text
cmd(4) = Mid((List1.Text), 1, 5)
cmd(5) = Trim(txtFields(0))
cmd(6) = Trim(txtFields(1))
cmd(7) = Trim(txtFields(2))
cmd(8) = Trim(soft_user)
cmd(9) = Date

cmd.Execute
MsgBox "Save Successfully.", vbInformation, "Student Management System"
ShowFlexData


End Sub


Private Sub Combo1_Click()
List1.Clear
txtFields(0) = ""
txtFields(1) = ""

Set rs1 = getdata("select Sub_code,Sub_title from Subject_Info_sub where class_code ='" & Mid((Combo1.Text), 1, 5) & "'")
If Not rs1.EOF Then
    Do Until rs1.EOF
        List1.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop
'    If List1.ListCount > 0 Then List1.ListIndex = 0

End If

Combo2.SetFocus
ShowFlexData
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    List1.Clear
    Set rs1 = getdata("select subjectId,SubjectDsc from SubjectInfo where classId ='" & Mid((Combo1.Text), 1, 5) & "'")
        If Not rs1.EOF Then
            Do Until rs1.EOF
                    List1.AddItem rs1(0) + " - " + rs1(1)
                    rs1.MoveNext
             Loop
            '    If List1.ListCount > 0 Then List1.ListIndex = 0
        
        End If
    Combo2.SetFocus
End If
End Sub


Private Sub Combo2_Click()
  ShowFlexData
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    List1.SetFocus
End If
ShowFlexData
End Sub

Private Sub Command2_Click()
  
End Sub

Private Sub Form_Load()
For i = 2000 To 2050
Combo2.AddItem i
Next

Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("select ClassId,ClassName from ClassInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
        Combo1.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop
'    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0

End If
With MSFlexGrid1
    .Rows = 1
    .Cols = 3
    .Col = 0: .Text = "               Book Name #"
    .Col = 1: .Text = " By "
    .Col = 2: .Text = " Publisher "
    .ColWidth(0) = 5000
    .ColWidth(1) = 4000
    .ColWidth(2) = 4000
   
    
End With

End Sub

Private Sub List1_Click()
txtFields(0) = ""
txtFields(1) = ""
cmdnew.SetFocus
ShowFlexData
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If keyasii = 13 Then
    cmdnew.SetFocus
End If
End Sub


Private Sub MSFlexGrid1_SelChange()
  MSFlexGrid1_Click
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
        Case 0
            txtFields(1).SetFocus
        Case 1
            txtFields(2).SetFocus
        Case 2
            cmdSave.SetFocus

    End Select
End If
End Sub
Private Sub ShowFlexData()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT Book,Writter,Publisher from Booklist where ClassID='" & Mid((Combo1.Text), 1, 5) & "'and Eyear='" & Mid((Combo2.Text), 1, 5) & "'and subjectid='" & Mid((List1.Text), 1, 5) & "'")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
               
                .TextMatrix(i, 0) = "" & rs!Book
                .TextMatrix(i, 1) = "" & rs!Writter
                .TextMatrix(i, 2) = "" & rs!Publisher
                
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
If MSFlexGrid1.Rows > 1 Then
txtFields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtFields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
txtFields(2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)

Exit Sub
End If
errdes:
'MsgBox Err.Description, vbInformation, App.Title

End Sub
