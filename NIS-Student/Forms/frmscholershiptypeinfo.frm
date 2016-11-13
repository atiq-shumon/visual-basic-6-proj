VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmscholershiptypeinfo 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   -30
      TabIndex        =   10
      Top             =   3750
      Width           =   7845
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
         Left            =   6750
         TabIndex        =   14
         ToolTipText     =   "Click to Exit"
         Top             =   240
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
         Left            =   5700
         TabIndex        =   13
         ToolTipText     =   "Click to Delete"
         Top             =   240
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
         Left            =   4650
         TabIndex        =   12
         ToolTipText     =   "Click to save"
         Top             =   240
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
         Left            =   3600
         MaskColor       =   &H8000000C&
         TabIndex        =   11
         ToolTipText     =   "Click to insert new information"
         Top             =   240
         Width           =   1005
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   3570
         Top             =   210
         Width           =   4215
      End
   End
   Begin VB.Frame Frame6 
      ForeColor       =   &H00C00000&
      Height          =   1155
      Left            =   0
      TabIndex        =   4
      Top             =   750
      Width           =   7815
      Begin VB.TextBox txtfields 
         Height          =   435
         Index           =   2
         Left            =   660
         MaxLength       =   80
         TabIndex        =   1
         ToolTipText     =   "Insert Short Note"
         Top             =   600
         Width           =   7065
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   1
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   0
         ToolTipText     =   "Insert Scholership type name"
         Top             =   240
         Width           =   4965
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   660
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   570
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Left            =   2070
         TabIndex        =   7
         Top             =   270
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   270
         Width           =   165
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   765
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7755
      TabIndex        =   2
      Top             =   -30
      Width           =   7815
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   3
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scholarship Type Set Up"
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
         Left            =   2250
         TabIndex        =   15
         Top             =   180
         Width           =   2835
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   -120
         Picture         =   "frmscholershiptypeinfo.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   7935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1875
      Left            =   0
      TabIndex        =   9
      Top             =   1890
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3307
      _Version        =   393216
      FixedCols       =   0
   End
End
Attribute VB_Name = "frmscholershiptypeinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim cmd As New adodb.Command
Dim con As New adodb.Connection
Dim rs As New adodb.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
    
If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical) = vbYes Then
    cmd.CommandType = adCmdText
    cmd.CommandText = "Delete from ScholershipType  where (SchTypeId = '" & Trim(txtfields(0)) & "') "
    cmd.Execute
    MsgBox "Delete Successfully Scholership Information.", vbInformation, App.Title
    
    For i = 0 To 2
        txtfields(i) = ""
    Next
    Call ShowFlexData
Else
    Exit Sub
End If
End Sub

Private Sub cmdnew_Click()
Dim cmd As New adodb.Command
Dim con As New adodb.Connection
con.Open GConnString
cmd.ActiveConnection = con
Dim rs As New adodb.Recordset
Set rs = getdata("select max((substring(SchTypeId,4,5)))+ 1 from ScholershipType")
If Not rs.EOF Then
    txtfields(0) = IIf(IsNull(rs(0)) = True, "ST" + "-" + "01", "ST" + "-" + Format(rs(0), "00"))
Else
    txtfields(0) = "ST-01"
End If
    txtfields(1) = ""
    txtfields(2).Text = ""
    txtfields(1).SetFocus

End Sub

Private Sub cmdSAVE_Click()
Dim cmd As New adodb.Command
Dim con As New adodb.Connection
con.Open GConnString
cmd.ActiveConnection = con

If Len(txtfields(1)) = 0 Then
    MsgBox "Please Enter Scholership Type Name.", vbCritical, App.Title
    txtfields(1).SetFocus
    Exit Sub
End If
   
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "ScholerShipTypeInformation"
    cmd(1) = txtfields(0)
    cmd(2) = Trim(txtfields(1))
    cmd(3) = Trim(txtfields(2))
    cmd(4) = "DSL"
    cmd(5) = Date
   
    cmd.Execute
    MsgBox "Save Successfully.", vbInformation, "Student Management System"
    cmdnew.SetFocus
  
    Call ShowFlexData


End Sub

Private Sub Form_Load()
With MSFlexGrid1
    .Rows = 1
    .Cols = 3
    .Col = 0: .Text = " Scholership Type ID   #"
    .Col = 1: .Text = "Scholership Type Name   "
    .Col = 2: .Text = " Note  "
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 4000
    .ColWidth(2) = 16000
   
End With
Dim rs As New adodb.Recordset
Set rs = getdata("select max((substring(SchTypeId,4,5)))+ 1 from ScholershipType")
If Not rs.EOF Then
    txtfields(0) = IIf(IsNull(rs(0)) = True, "ST" + "-" + "01", "ST" + "-" + Format(rs(0), "00"))
Else
    txtfields(0) = "ST-01"
End If
ShowFlexData
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
        Case 1
            txtfields(2).SetFocus
        Case 2
            cmdsave.SetFocus
    End Select
End If
End Sub

Private Sub ShowFlexData()
On Error GoTo errdes
Dim rs As New adodb.Recordset
Set rs = getdata("SELECT   SchTypeId, ScTypeName ,Notes from ScholershipType")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .TextMatrix(i, 0) = "" & rs!SchTypeId
                .TextMatrix(i, 1) = "" & rs!ScTypeName
                .TextMatrix(i, 2) = "" & rs!Notes
              
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
txtfields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
txtfields(2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)

Exit Sub
errdes:

End Sub



