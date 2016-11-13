VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptBirthalert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report : List birth dates"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
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
      Left            =   0
      TabIndex        =   6
      Top             =   540
      Width           =   5445
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Index           =   0
         Left            =   990
         TabIndex        =   7
         Top             =   630
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         _Version        =   393216
         Format          =   48824321
         CurrentDate     =   38750
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Index           =   1
         Left            =   3390
         TabIndex        =   8
         Top             =   630
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         _Version        =   393216
         Format          =   48824321
         CurrentDate     =   38750
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
         TabIndex        =   9
         Top             =   660
         Width           =   405
      End
   End
   Begin VB.CommandButton cmd 
      Height          =   405
      Index           =   1
      Left            =   4785
      Picture         =   "rptBirthdayaltert.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exit"
      Top             =   1890
      Width           =   585
   End
   Begin VB.CommandButton cmd 
      Height          =   405
      Index           =   0
      Left            =   4200
      Picture         =   "rptBirthdayaltert.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Report Veiwer"
      Top             =   1890
      Width           =   555
   End
   Begin VB.Frame Frame2 
      Height          =   1125
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   5385
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   585
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   5325
      TabIndex        =   0
      Top             =   0
      Width           =   5385
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   1
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report : Student Birthday Alert"
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
         Left            =   900
         TabIndex        =   2
         Top             =   120
         Width           =   3510
      End
   End
   Begin VB.Shape Shape1 
      Height          =   525
      Left            =   4110
      Top             =   1830
      Width           =   1305
   End
End
Attribute VB_Name = "rptBirthalert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub CboClassID_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'    cmd(0).SetFocus
'End If
'End Sub
'
'Private Sub CboClassID_Click()
'Set rs = GetData("SELECT ClassName FROM Classinfo " + _
'            "WHERE (ClassID = '" & CboClassID & "')")
'
'    If Not rs.EOF Then
'       txtClassName.Text = rs!ClassName
'    Else
'       txtClassName = ""
'   End If
'
'End Sub
'
Private Sub cmd_Click(Index As Integer)
Dim rs As New ADODB.Recordset

Select Case Index
    Case 0
       rptMode = 3
       Screen.MousePointer = vbHourglass
       frmViewer.Show 1

    Case 1
        Unload Me
End Select
End Sub

'
'
'Private Sub DTPicker1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Select Case Index
'    Case 0
'        If KeyCode = 13 Then
'            DTPicker1(1).SetFocus
'        End If
'    Case 1
'        If KeyCode = 13 Then
'            cmd(0).SetFocus
'        End If
'
'End Select
'End Sub
'
'Private Sub Form_Load()
'Dim rs As New adodb.Recordset
' Set rs = GetData("Select classid from classinfo")
'
'If rs.EOF = False Then
'    Do Until rs.EOF
'        CboClassID.AddItem rs(0)
'        rs.MoveNext
'    Loop
'End If
'
'Option1.Value = True
'
''MaskDate(0).Text = Format(Date, "dd/mm/yy")
''MaskDate(1).Text = Format(Date, "dd/mm/yy")
'
'End Sub
'
'Private Sub Option1_Click()
'Frame3.Visible = True
'Frame3.Top = 1000
'Frame3.Left = 120
'
'Frame4.Visible = False
'End Sub
'
''Private Sub MaskDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
''Select Case Index
''Case 0
'' If KeyCode = 13 Then
''    If MaskDate(0) <> "__/__/__" Then
''            If Check_ValidDate(MaskDate(0)) = False Then
''                MaskDate(1).SetFocus
''                Exit Sub
''            End If
''    Else
''        MaskDate(1).SetFocus
''    End If
'' End If
''Case 1
'' If KeyCode = 13 Then
''    If MaskDate(1) <> "__/__/__" Then
''            If Check_ValidDate(MaskDate(1)) = False Then
''                cmd(0).SetFocus
''                Exit Sub
''            End If
''    Else
''        cmd(0).SetFocus
''    End If
'' End If
''End Select
'' End Sub
''
''Private Sub Option1_Click()
''Frame3.Visible = True
''Frame3.Top = 1000
''Frame3.Left = 120
''
''Frame4.Visible = False
''End Sub
'
'Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'    CboClassID.SetFocus
'End If
'End Sub
'
'Private Sub Option2_Click()
'Frame4.Visible = True
'Frame4.Top = 1000
'Frame4.Left = 120
'
'Frame3.Visible = False
'End Sub
'
'Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'    DTPicker1(0).SetFocus
'End If
'End Sub

Private Sub Form_Load()

End Sub
