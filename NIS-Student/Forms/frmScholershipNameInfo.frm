VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmScholershipNameInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   -30
      TabIndex        =   20
      Top             =   5280
      Width           =   7875
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
         Height          =   405
         Left            =   6780
         TabIndex        =   8
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
         Height          =   405
         Left            =   5730
         TabIndex        =   7
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
         Height          =   405
         Left            =   4680
         TabIndex        =   5
         ToolTipText     =   "Click to Save"
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
         Height          =   405
         Left            =   3630
         MaskColor       =   &H8000000C&
         TabIndex        =   6
         ToolTipText     =   "Click to insert new information"
         Top             =   240
         Width           =   1005
      End
      Begin VB.Shape Shape1 
         Height          =   465
         Left            =   3570
         Top             =   210
         Width           =   4275
      End
   End
   Begin VB.Frame Frame6 
      ForeColor       =   &H00C00000&
      Height          =   2715
      Left            =   0
      TabIndex        =   11
      Top             =   780
      Width           =   7815
      Begin RichTextLib.RichTextBox RichDescription 
         Height          =   1005
         Left            =   1140
         TabIndex        =   4
         ToolTipText     =   "Write Short Notes about the  Scholarship "
         Top             =   1530
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1773
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmScholershipNameInfo.frx":0000
      End
      Begin VB.TextBox txtfields 
         Height          =   315
         Index           =   3
         Left            =   1140
         MaxLength       =   80
         TabIndex        =   3
         ToolTipText     =   "Insert Address"
         Top             =   1200
         Width           =   6615
      End
      Begin VB.ComboBox ComboType 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select Scholarship Type"
         Top             =   180
         Width           =   3555
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1140
         TabIndex        =   12
         ToolTipText     =   "Insert ID"
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   1
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Insert Scholarship  Name"
         Top             =   540
         Width           =   6615
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   2
         Left            =   1140
         MaxLength       =   80
         TabIndex        =   2
         ToolTipText     =   "Insert Name of Scholarship Declared By"
         Top             =   870
         Width           =   6615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   1260
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Left            =   3630
         TabIndex        =   17
         Top             =   210
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   240
         Width           =   165
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   570
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Declared By"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   900
         Width           =   870
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   765
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7755
      TabIndex        =   9
      Top             =   0
      Width           =   7815
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   10
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scholarship Name Set Up"
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
         Left            =   1890
         TabIndex        =   21
         Top             =   180
         Width           =   2925
      End
      Begin VB.Image Image1 
         Height          =   930
         Left            =   -60
         Picture         =   "frmScholershipNameInfo.frx":0082
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   7845
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1755
      Left            =   0
      TabIndex        =   16
      Top             =   3480
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3096
      _Version        =   393216
      FixedCols       =   0
   End
End
Attribute VB_Name = "frmScholershipNameInfo"
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
    cmd.CommandText = "Delete from scholerShipNameSetup  where (ScholerShipNameId = '" & Trim(txtfields(0)) & "') "
    cmd.Execute
    MsgBox "Delete Successfully Scholership Information.", vbInformation, App.Title
    
    For i = 0 To 3
        txtfields(i) = ""
    Next
    ComboType = " "
    RichDescription = ""
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
Set rs = getdata("select max((substring(ScholerShipNameId,4,5)))+ 1 from scholerShipNameSetup")
If Not rs.EOF Then
    txtfields(0) = IIf(IsNull(rs(0)) = True, "SN" + "-" + "01", "SN" + "-" + Format(rs(0), "00"))
Else
    txtfields(0) = "SN-01"
End If
    txtfields(1) = ""
    txtfields(2).Text = ""
    txtfields(3) = ""
    RichDescription = ""
    ComboType = " "
    ComboType.SetFocus

End Sub
Private Sub cmdSAVE_Click()
Dim cmd As New adodb.Command
Dim con As New adodb.Connection
con.Open GConnString
cmd.ActiveConnection = con
Dim rs As New adodb.Recordset
Dim SType As String
Set rs = getdata("select SchTypeId from ScholershipType where ScTypeName='" & ComboType & "'")
If Not rs.EOF Then
    SType = rs!SchTypeId
End If

 If Len(txtfields(1)) = 0 Then
     MsgBox "Please Enter Scholership  Name.", vbCritical, App.Title
     txtfields(1).SetFocus
     Exit Sub
 End If
 
If Len(txtfields(2)) = 0 Then
     MsgBox "Please Enter Declared By.", vbCritical, App.Title
     txtfields(2).SetFocus
     Exit Sub
 End If
 
If Len(txtfields(3)) = 0 Then
     MsgBox "Please Enter Address.", vbCritical, App.Title
     txtfields(3).SetFocus
     Exit Sub
 End If
    
If Len(RichDescription) = 0 Then
    MsgBox "Please Enter Description.", vbCritical, App.Title
    RichDescription.SetFocus
    Exit Sub
End If
    
If Len(ComboType) = 0 Then
   MsgBox "Please Enter Type of Scholership.", vbCritical, App.Title
   ComboType.SetFocus
   Exit Sub
End If

 cmd.CommandType = adCmdStoredProc
 cmd.CommandText = "ScholerShipNameSetupInformation"
 cmd(1) = txtfields(0)
 cmd(2) = SType
 cmd(3) = Trim(txtfields(1))
 cmd(4) = Trim(txtfields(2))
 cmd(5) = Trim(txtfields(3))
 cmd(6) = RichDescription.TextRTF
 cmd(7) = "DSL"
 cmd(8) = Date

 cmd.Execute
 MsgBox "Save Successfully.", vbInformation, "Student Management System"
 cmdnew.SetFocus

 Call ShowFlexData


End Sub



Private Sub ComboType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtfields(1).SetFocus
End If
End Sub

Private Sub Form_Load()
With MSFlexGrid1
    .Rows = 1
    .Cols = 4
    .Col = 0: .Text = " ID   #"
    .Col = 1: .Text = "Scholership Name   "
    .Col = 3: .Text = " Declared By  "
    .Col = 2: .Text = " Type of Scholership  "
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 6000
    .ColWidth(2) = 4000
    .ColWidth(3) = 8000
   
End With
Dim rs As New adodb.Recordset
Set rs = getdata("select max((substring(ScholerShipNameId,4,5)))+ 1 from scholerShipNameSetup")
If Not rs.EOF Then
    txtfields(0) = IIf(IsNull(rs(0)) = True, "SN" + "-" + "01", "SN" + "-" + Format(rs(0), "00"))
Else
    txtfields(0) = "SN-01"
End If

Set rs = getdata("select ScTypeName from ScholershipType")
If Not rs.EOF Then
Do Until rs.EOF
ComboType.AddItem rs!ScTypeName
rs.MoveNext
Loop
'ComboType = " "
End If
ShowFlexData
End Sub
Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
Select Case Index
Case 1
    txtfields(2).SetFocus
Case 2
    txtfields(3).SetFocus
Case 3
    RichDescription.SetFocus
    
End Select
End If
End Sub
Private Sub ShowFlexData()
On Error GoTo errdes
Dim rs As New adodb.Recordset
Set rs = getdata("SELECT scholerShipNameSetup.ScholerShipNameId, scholerShipNameSetup.ScName, ScholershipType.ScTypeName, scholerShipNameSetup.SchBy FROM  ScholershipType INNER JOIN scholerShipNameSetup ON ScholershipType.SchTypeId = scholerShipNameSetup.SchType")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                
                .TextMatrix(i, 0) = "" & rs!ScholerShipNameId
                .TextMatrix(i, 1) = "" & rs!ScName
                .TextMatrix(i, 2) = "" & rs!ScTypeName
                .TextMatrix(i, 3) = "" & rs!SchBy
            
                
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
txtfields(2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
ComboType = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
Dim rs As New adodb.Recordset
Set rs = getdata("SELECT Address, DesOfSch FROM scholerShipNameSetup where ScholerShipNameId='" & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) & "' ")
If Not rs.EOF Then
    txtfields(3) = rs!Address
    RichDescription = rs!DesOfSch
End If

Exit Sub
errdes:


End Sub





