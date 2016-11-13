VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmSupplierInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   -60
      TabIndex        =   15
      Top             =   5130
      Width           =   7485
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
         Height          =   405
         Left            =   6180
         TabIndex        =   7
         ToolTipText     =   "Click to Exit"
         Top             =   210
         Width           =   945
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
         Height          =   405
         Left            =   5190
         TabIndex        =   6
         ToolTipText     =   "Click to Delete"
         Top             =   210
         Width           =   945
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
         Height          =   405
         Left            =   4200
         TabIndex        =   4
         ToolTipText     =   "Click to Save"
         Top             =   210
         Width           =   945
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
         Height          =   405
         Left            =   3210
         TabIndex        =   5
         ToolTipText     =   "Click to insert new information"
         Top             =   210
         Width           =   945
      End
      Begin VB.Shape Shape1 
         Height          =   465
         Left            =   3180
         Top             =   180
         Width           =   3975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2025
      Left            =   -30
      TabIndex        =   14
      Top             =   3150
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   3572
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   765
      Left            =   -30
      ScaleHeight     =   705
      ScaleWidth      =   7215
      TabIndex        =   9
      Top             =   -30
      Width           =   7275
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Information"
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
         Left            =   2190
         TabIndex        =   16
         Top             =   150
         Width           =   2325
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   -30
         Picture         =   "FrmSupplierInfo.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   7245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   7275
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   3
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   3
         ToolTipText     =   "Insert Contact No"
         Top             =   1890
         Width           =   2355
      End
      Begin VB.TextBox txtfields 
         Height          =   615
         Index           =   2
         Left            =   1050
         TabIndex        =   2
         ToolTipText     =   "Insert Supplier Address"
         Top             =   1140
         Width           =   5925
      End
      Begin VB.TextBox txtfields 
         Height          =   315
         Index           =   1
         Left            =   1050
         TabIndex        =   1
         ToolTipText     =   "Insert Supplier Name"
         Top             =   720
         Width           =   4365
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1050
         TabIndex        =   0
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1170
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   330
         Width           =   165
      End
   End
End
Attribute VB_Name = "FrmSupplierInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDelete_Click()
Dim cmd As New adodb.Command
Dim con As New adodb.Connection
Dim rs As New adodb.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
If (MsgBox("Are You sure to delete ?", vbYesNo + vbCritical) = vbYes) Then
    cmd.CommandType = adCmdText
    cmd.CommandText = "Delete from SupplierInfo  where (SuppID = '" & Trim(txtfields(0)) & "') "
    cmd.Execute
    MsgBox "Delete Successfully Supplier Information.", vbInformation, App.Title
    For i = 0 To 3
     txtfields(i) = ""
    Next
    Call ShowFlexData
Else
     Exit Sub
        
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
Dim cmd As New adodb.Command
Dim con As New adodb.Connection
con.Open GConnString
cmd.ActiveConnection = con
For i = 0 To 3
   txtfields(i) = ""
Next
txtfields(1).SetFocus
'Call ShowFlexData
End Sub

Private Sub cmdSAVE_Click()
Dim cmd As New adodb.Command
Dim con As New adodb.Connection
con.Open GConnString
cmd.ActiveConnection = con

    If Len(txtfields(1)) = 0 Then
        MsgBox "Please Enter Supplier Name.", vbInformation, App.Title
        txtfields(1).SetFocus
        Exit Sub
    End If
    If Len(txtfields(2)) = 0 Then
       MsgBox "Please Enter Address.", vbInformation, App.Title
       txtfields(2).SetFocus
       Exit Sub
    End If
    
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SupplierInformation"
    cmd(1) = txtfields(0)
    cmd(2) = Trim(txtfields(1))
    cmd(3) = Trim(txtfields(2))
    cmd(4) = Trim(txtfields(3))
    cmd(5) = Date
    cmd(6) = "DSL"

    cmd.Execute
    MsgBox "Save Successfully.", vbInformation, "Student Management System"
    cmdnew.SetFocus
    get_Maximum
    Call ShowFlexData

End Sub

Private Sub Form_Load()

With MSFlexGrid1
    .Rows = 1
    .Cols = 4
    .Col = 0: .Text = "                  ID#"
    .Col = 1: .Text = "Name"
    .Col = 2: .Text = "Address"
    .Col = 3: .Text = "Phone No"
    
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 4000
    .ColWidth(2) = 5500
    .ColWidth(3) = 2000
End With
Call ShowFlexData

Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo errdes
txtfields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
txtfields(2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
txtfields(3) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
Exit Sub
errdes:
'MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
        Case 1
            txtfields(2).SetFocus
        Case 2
            txtfields(3).SetFocus
        Case 3
            cmdsave.SetFocus
    End Select
End If
End Sub
Private Sub get_Maximum()
Dim rs As New adodb.Recordset
Set rs = getdata("select max(SuppID) from SupplierInfo")
If Not rs.EOF Then
        txtfields(0) = rs.Fields(0)
Else
    txtfields(0) = "Supp-00001"
End If
End Sub
Private Sub ShowFlexData()
On Error GoTo errdes
Dim rs As New adodb.Recordset
Set rs = getdata("SELECT SuppID,SuppName,SuppAddr,Phone From SupplierInfo ")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .TextMatrix(i, 0) = "" & rs!SuppID
                .TextMatrix(i, 1) = "" & rs!SuppName
                .TextMatrix(i, 2) = "" & rs!SuppAddr
                .TextMatrix(i, 3) = "" & rs!Phone
                
                i = i + 1
            rs.MoveNext
        Loop
    End With
Else
    MSFlexGrid1.Rows = 1
 End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title
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

