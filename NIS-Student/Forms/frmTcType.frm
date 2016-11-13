VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTcType 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   0
      TabIndex        =   10
      Top             =   3750
      Width           =   7815
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
         Left            =   3570
         MaskColor       =   &H8000000C&
         TabIndex        =   14
         ToolTipText     =   "Click to insert new information"
         Top             =   270
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
         Left            =   4590
         TabIndex        =   13
         ToolTipText     =   "Click to Save"
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton cmddelete 
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
         Left            =   5640
         TabIndex        =   12
         ToolTipText     =   "Click to Delete"
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton cmdClose 
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
         Left            =   6690
         TabIndex        =   11
         ToolTipText     =   "Click to Exit"
         Top             =   270
         Width           =   1005
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   3510
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   9915
      TabIndex        =   7
      Top             =   -210
      Width           =   9975
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   8
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer Certificate Type Set Up"
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
         Left            =   1380
         TabIndex        =   15
         Top             =   360
         Width           =   3630
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   -90
         Picture         =   "frmTcType.frx":0000
         Stretch         =   -1  'True
         Top             =   150
         Width           =   7905
      End
   End
   Begin VB.Frame Frame6 
      ForeColor       =   &H00C00000&
      Height          =   1485
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   7815
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   660
         TabIndex        =   3
         Text            =   "Auto"
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   3
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   0
         ToolTipText     =   "Insert Name of the TC"
         Top             =   300
         Width           =   4965
      End
      Begin VB.TextBox txtfields 
         Height          =   525
         Index           =   4
         Left            =   660
         MaxLength       =   80
         TabIndex        =   1
         ToolTipText     =   "Insert Short Note"
         Top             =   750
         Width           =   7065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TC ID"
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TC Name"
         Height          =   195
         Left            =   1950
         TabIndex        =   5
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
         Height          =   255
         Left            =   90
         TabIndex        =   4
         Top             =   780
         Width           =   345
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1605
      Left            =   0
      TabIndex        =   9
      Top             =   2190
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2831
      _Version        =   393216
      FixedCols       =   0
   End
End
Attribute VB_Name = "frmTcType"
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
    
If MsgBox("are You Sure To Delete?", vbYesNo + vbCritical) = vbYes Then
    cmd.CommandType = adCmdText
    cmd.CommandText = "Delete from TCTypeSetUp  where (TCID = '" & Trim(txtfields(2)) & "') "
    cmd.Execute
    MsgBox "Delete Successfully TC Information.", vbInformation, App.Title
    
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

    Set rs = getdata("select max((substring(TcId,4,5)))+ 1 from TCTypeSetUp")
    If Not rs.EOF Then
        txtfields(2) = IIf(IsNull(rs(0)) = True, "Tc" + "-" + "01", "Tc" + "-" + Format(rs(0), "00"))
    Else
        txtfields(0) = "Tc-01"
    End If


   
    txtfields(3) = ""
    txtfields(4) = ""
   
'    txtfields(3).SetFocus



End Sub


Private Sub Form_Load()
With MSFlexGrid1
    .Rows = 1
    .Cols = 3
    .Col = 0: .Text = " Transfer Certificate ID   #"
    .Col = 1: .Text = "Transfer Certificate Name   "
    .Col = 2: .Text = " Note  "
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 5000
    .ColWidth(2) = 5000
   
End With
ShowFlexData
cmdnew_Click
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
        Case 2
            txtfields(3).SetFocus
        Case 3
            txtfields(4).SetFocus
        Case 4
            cmdSave.SetFocus
    End Select
End If
End Sub
Private Sub cmdSAVE_Click()
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con

    If Len(txtfields(3)) = 0 Then
        MsgBox "Please Enter TC Type Name.", vbCritical, App.Title
        txtfields(3).SetFocus
        Exit Sub
    End If
   
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "TCType"
    cmd(1) = txtfields(2)
    cmd(2) = Trim(txtfields(3))
    cmd(3) = Trim(txtfields(4))
    cmd(4) = "DSL"
    cmd(5) = Date
   
    cmd.Execute
    MsgBox "Save Successfully.", vbInformation, "Student Management System"
    cmdnew.SetFocus
  
    Call ShowFlexData

End Sub

Private Sub ShowFlexData()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT   TCID, TCName ,Note from TCTypeSetUp")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                
                .TextMatrix(i, 0) = "" & rs!TCID
                .TextMatrix(i, 1) = "" & rs!TcName
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


