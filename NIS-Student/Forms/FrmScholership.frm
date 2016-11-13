VERSION 5.00
Begin VB.Form FrmScholership 
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3990
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   435
      Left            =   6930
      TabIndex        =   22
      ToolTipText     =   "Click to Exit"
      Top             =   3510
      Width           =   945
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H8000000C&
      Caption         =   "Delete"
      Height          =   435
      Left            =   5970
      TabIndex        =   21
      ToolTipText     =   "Click to Delete"
      Top             =   3510
      Width           =   945
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   435
      Left            =   5010
      TabIndex        =   20
      ToolTipText     =   "Click to Save"
      Top             =   3510
      Width           =   945
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H8000000C&
      Caption         =   "New"
      Height          =   435
      Left            =   4050
      MaskColor       =   &H8000000C&
      TabIndex        =   19
      ToolTipText     =   "Click to insert new information"
      Top             =   3510
      Width           =   945
   End
   Begin VB.Frame Frame3 
      Caption         =   "ScholerShip Information"
      ForeColor       =   &H00FF0000&
      Height          =   1425
      Left            =   0
      TabIndex        =   12
      Top             =   2040
      Width           =   7875
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   6330
         TabIndex        =   18
         ToolTipText     =   "Insert Scholarship Amount"
         Top             =   990
         Width           =   1455
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   4740
         TabIndex        =   17
         ToolTipText     =   "Insert Valid Years"
         Top             =   990
         Width           =   765
      End
      Begin VB.ComboBox ComboGrade 
         Height          =   315
         Left            =   960
         TabIndex        =   16
         ToolTipText     =   "Select Grade"
         Top             =   990
         Width           =   2265
      End
      Begin VB.ComboBox ComboScholerType 
         Height          =   315
         ItemData        =   "FrmScholership.frx":0000
         Left            =   960
         List            =   "FrmScholership.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Select Scholership type"
         Top             =   600
         Width           =   6825
      End
      Begin VB.ComboBox ComScholerClass 
         Height          =   315
         Left            =   5490
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Select Schlarship Gaining Class "
         Top             =   210
         Width           =   2295
      End
      Begin VB.ComboBox ComboYr 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Select Year"
         Top             =   210
         Width           =   2175
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   5640
         TabIndex        =   28
         Top             =   1020
         Width           =   540
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Of Valid Year"
         Height          =   195
         Left            =   3420
         TabIndex        =   27
         Top             =   1020
         Width           =   1185
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SholarShip Gaining Class"
         Height          =   195
         Left            =   3420
         TabIndex        =   26
         Top             =   240
         Width           =   1770
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grade"
         Height          =   195
         Left            =   90
         TabIndex        =   25
         Top             =   990
         Width           =   435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Left            =   90
         TabIndex        =   24
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         Height          =   195
         Left            =   90
         TabIndex        =   23
         Top             =   240
         Width           =   330
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7815
      TabIndex        =   10
      Top             =   0
      Width           =   7875
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   11
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ScholarShip Information"
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
         Left            =   2340
         TabIndex        =   29
         Top             =   120
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   810
         Left            =   -90
         Picture         =   "FrmScholership.frx":0004
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   7935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Caption         =   "Student's Personal Information"
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   690
      Width           =   7875
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   900
         TabIndex        =   4
         Top             =   930
         Width           =   2295
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3930
         TabIndex        =   3
         Top             =   570
         Width           =   3885
      End
      Begin VB.ComboBox ComStuId 
         Height          =   315
         Left            =   900
         TabIndex        =   2
         ToolTipText     =   "Select Student"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   900
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student ID"
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   3180
         TabIndex        =   8
         Top             =   240
         Width           =   4635
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class "
         Height          =   195
         Left            =   3360
         TabIndex        =   6
         Top             =   630
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         Height          =   195
         Left            =   60
         TabIndex        =   5
         Top             =   960
         Width           =   540
      End
   End
End
Attribute VB_Name = "FrmScholership"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
ComStuId.Text = ""
'ComboYr.Text = " "
ComboGrade.Text = " "
ComboScholerType.Text = " "
ComScholerClass.Text = " "
txtFields(0).Text = ""
txtFields(1).Text = ""
txtFields(2).Text = ""
txtFields(3).Text = ""
txtFields(4).Text = ""
Label3.Caption = ""
ComStuId.SetFocus
End Sub

Private Sub ComStuId_Change()
If Len(ComStuId.Text) = 0 Then
    Label3.Caption = ""
End If
End Sub

Private Sub ComStuId_Click()
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT StudentName From StudentInfo WHERE StudentID = '" & Trim(ComStuId) & "'")
If Not rs.EOF Then
    If Not rs.EOF Then
        Label3.Caption = rs.Fields(0)
    Else
        MsgBox "Invalid Stdudent !", vbCritical, "Daffodil Software Ltd"
            
    End If

End If

End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT StudentInfo.StudentID, StudentEvaluation.Active FROM StudentEvaluation INNER JOIN " + _
"StudentInfo ON StudentEvaluation.StudentID = StudentInfo.StudentID WHERE     (StudentEvaluation.Active = 'Y')")
If Not rs.EOF Then
    Do Until rs.EOF
        ComStuId.AddItem rs!StudentID
        rs.MoveNext
    Loop

End If

Dim yr As Integer
For yr = 2000 To 2020
ComboYr.AddItem yr
Next


End Sub
