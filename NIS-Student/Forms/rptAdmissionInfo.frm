VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptAdmissionInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report : Admission Information"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Height          =   405
      Index           =   1
      Left            =   4785
      Picture         =   "rptAdmissionInfo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exit"
      Top             =   3750
      Width           =   585
   End
   Begin VB.CommandButton cmd 
      Height          =   405
      Index           =   0
      Left            =   4170
      Picture         =   "rptAdmissionInfo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Report Veiwer"
      Top             =   3750
      Width           =   555
   End
   Begin VB.Frame Frame2 
      Height          =   3045
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   5415
      Begin VB.Frame Frame4 
         Caption         =   "Select Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   90
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   5205
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   345
            Index           =   0
            Left            =   990
            TabIndex        =   15
            Top             =   630
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   609
            _Version        =   393216
            Format          =   50855937
            CurrentDate     =   38750
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   345
            Index           =   1
            Left            =   3390
            TabIndex        =   16
            Top             =   630
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   609
            _Version        =   393216
            Format          =   50855937
            CurrentDate     =   38750
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   2940
            TabIndex        =   11
            Top             =   660
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   270
            TabIndex        =   10
            Top             =   660
            Width           =   660
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Select Class"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   90
         TabIndex        =   5
         Top             =   1000
         Visible         =   0   'False
         Width           =   5205
         Begin VB.TextBox txtClassName 
            Enabled         =   0   'False
            Height          =   345
            Left            =   2010
            TabIndex        =   14
            Top             =   945
            Width           =   2475
         End
         Begin VB.ComboBox CboClassID 
            Height          =   315
            Left            =   2010
            TabIndex        =   12
            Top             =   390
            Width           =   2475
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Class"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   13
            Top             =   450
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   9
            Top             =   1020
            Width           =   1035
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Date to Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   1
         Top             =   480
         Width           =   1425
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By Class"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1140
         TabIndex        =   0
         Top             =   480
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   585
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   5325
      TabIndex        =   2
      Top             =   0
      Width           =   5385
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   3
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report : Student Admission Information"
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
         Left            =   540
         TabIndex        =   17
         Top             =   120
         Width           =   5100
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   -90
         Picture         =   "rptAdmissionInfo.frx":1194
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   5475
      End
   End
   Begin VB.Shape Shape1 
      Height          =   525
      Left            =   4110
      Top             =   3690
      Width           =   1305
   End
End
Attribute VB_Name = "rptAdmissionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CboClassID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmd(0).SetFocus
End If
End Sub

Private Sub CboClassID_Click()
Set rs = getdata("SELECT ClassName FROM Classinfo " + _
            "WHERE (ClassID = '" & CboClassID & "')")
                
    If Not rs.EOF Then
       txtClassName.Text = rs!classname
    Else
       txtClassName = ""
   End If

End Sub

Private Sub cmd_Click(Index As Integer)
Dim rs As New ADODB.Recordset

Select Case Index
    Case 0
       rptMode = 0
       Screen.MousePointer = vbHourglass
       frmViewer.Show 1
    
    Case 1
        Unload Me
End Select
End Sub



Private Sub DTPicker1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case 0
        If KeyCode = 13 Then
            DTPicker1(1).SetFocus
        End If
    Case 1
        If KeyCode = 13 Then
            cmd(0).SetFocus
        End If

End Select
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
 Set rs = getdata("Select classid from classinfo")
 
If rs.EOF = False Then
    Do Until rs.EOF
        CboClassID.AddItem rs(0)
        rs.MoveNext
    Loop
End If

Option1.Value = True

'MaskDate(0).Text = Format(Date, "dd/mm/yy")
'MaskDate(1).Text = Format(Date, "dd/mm/yy")

End Sub

Private Sub Option1_Click()
Frame3.Visible = True
Frame3.Top = 1000
Frame3.Left = 120

Frame4.Visible = False
End Sub

'Private Sub MaskDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Select Case Index
'Case 0
' If KeyCode = 13 Then
'    If MaskDate(0) <> "__/__/__" Then
'            If Check_ValidDate(MaskDate(0)) = False Then
'                MaskDate(1).SetFocus
'                Exit Sub
'            End If
'    Else
'        MaskDate(1).SetFocus
'    End If
' End If
'Case 1
' If KeyCode = 13 Then
'    If MaskDate(1) <> "__/__/__" Then
'            If Check_ValidDate(MaskDate(1)) = False Then
'                cmd(0).SetFocus
'                Exit Sub
'            End If
'    Else
'        cmd(0).SetFocus
'    End If
' End If
'End Select
' End Sub
'
'Private Sub Option1_Click()
'Frame3.Visible = True
'Frame3.Top = 1000
'Frame3.Left = 120
'
'Frame4.Visible = False
'End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CboClassID.SetFocus
End If
End Sub

Private Sub Option2_Click()
Frame4.Visible = True
Frame4.Top = 1000
Frame4.Left = 120

Frame3.Visible = False
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    DTPicker1(0).SetFocus
End If
End Sub
