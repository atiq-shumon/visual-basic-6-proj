VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVaccineSetUp 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   -30
      TabIndex        =   9
      Top             =   4200
      Width           =   8025
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H8000000C&
         Caption         =   "Edit"
         Height          =   375
         Left            =   4770
         TabIndex        =   14
         ToolTipText     =   "Click to Save"
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H8000000C&
         Caption         =   "Close"
         Height          =   375
         Left            =   6810
         TabIndex        =   11
         ToolTipText     =   "Click to Exit"
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H8000000C&
         Caption         =   "Delete"
         Height          =   375
         Left            =   5790
         TabIndex        =   10
         ToolTipText     =   "Click to Delete"
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H8000000C&
         Caption         =   "Save"
         Height          =   375
         Left            =   3750
         TabIndex        =   2
         ToolTipText     =   "Click to Save"
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton cmdnew 
         BackColor       =   &H8000000C&
         Caption         =   "New"
         Height          =   375
         Left            =   2730
         MaskColor       =   &H8000000C&
         TabIndex        =   3
         ToolTipText     =   "Click to insert new information"
         Top             =   210
         Width           =   975
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   2700
         Top             =   180
         Width           =   5115
      End
   End
   Begin VB.Frame Frame6 
      ForeColor       =   &H00C00000&
      Height          =   705
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   7845
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   1
         Left            =   3810
         MaxLength       =   75
         TabIndex        =   0
         Top             =   240
         Width           =   3915
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H80000018&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   1020
         MaxLength       =   3
         TabIndex        =   1
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vaccine Name"
         Height          =   195
         Left            =   2700
         TabIndex        =   8
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vaccine ID"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   7785
      TabIndex        =   4
      Top             =   0
      Width           =   7845
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   5
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vaccine Set Up"
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
         Left            =   2610
         TabIndex        =   12
         Top             =   180
         Width           =   1740
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   -30
         Picture         =   "frmVaccineSetUp.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   7905
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2745
      Left            =   0
      TabIndex        =   13
      Top             =   1440
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   4842
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   15005934
      BackColorSel    =   -2147483624
      ForeColorSel    =   16711680
      BackColorBkg    =   15724265
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmVaccineSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
On Error GoTo ErrorDes

Unload Me

Exit Sub
ErrorDes:    MsgBox Err.Description, vbCritical, App.Title

End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorDes

Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con

If Len(txtfields(0)) = 0 Then
        MsgBox "Please Enter an Vaccine ID...or choose form list below", vbCritical, App.Title
        cmdnew.SetFocus
        Exit Sub
    End If
    
  
   
If (MsgBox("Are You sure to delete ?", vbYesNo + vbCritical) = vbYes) Then
    cmd.CommandType = adCmdText
    cmd.CommandText = "Delete from VaccineInfo  where (VaccineID = '" & Trim(txtfields(0)) & "') "
    cmd.Execute
    
    MsgBox "Delete Successfully Vaccine Information.", vbInformation, App.Title
    
    For i = 0 To 1
     txtfields(i) = ""
    Next
    
    Call ShowFlexData
Else
    Exit Sub
End If

Exit Sub
ErrorDes:    MsgBox Err.Description, vbCritical, App.Title

End Sub


Private Sub cmdEdit_Click()
On Error GoTo ErrorDes

    Call cmdSAVE_Click

Exit Sub
ErrorDes:    MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdnew_Click()
On Error GoTo ErrorDes

Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString

cmd.ActiveConnection = con

    Set rs = getdata("select max(cast(VaccineID as int))+ 1 from VaccineInfo")
    If Not rs.EOF Then
        txtfields(0) = IIf(IsNull(rs(0)) = True, "01", Format(rs(0), "00"))
    Else
        txtfields(1) = "01"
    End If
    
    txtfields(1) = ""
    txtfields(1).SetFocus

Exit Sub
ErrorDes:    MsgBox Err.Description, vbCritical, App.Title
End Sub


Private Sub Form_Load()
On Error GoTo ErrorDes

With MSFlexGrid1
    .Rows = 1
    .Cols = 2
    .Col = 0: .Text = "Vaccine ID #"
    .Col = 1: .Text = "Vaccine Name"
    
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 6250
    
    .ColAlignment(0) = 4
    .ColAlignment(1) = 1
   
End With
ShowFlexData

Exit Sub
ErrorDes:    MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Sub MSFlexGrid1_SelChange()
On Error GoTo ErrorDes

  MSFlexGrid1_Click

Exit Sub
ErrorDes:    MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Sub txtfields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorDes

If KeyCode = 13 Then

    Select Case Index
        Case 0
            txtfields(1).SetFocus
        Case 1
            cmdsave.SetFocus
    End Select

End If

Exit Sub
ErrorDes:    MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ErrorDes
If KeyAscii = 13 Then
    Select Case Index
        Case 2
            txtfields(3).SetFocus
        Case 3
            txtfields(4).SetFocus
        Case 4
            cmdsave.SetFocus
    End Select
End If

Exit Sub
ErrorDes:    MsgBox Err.Description, vbCritical, App.Title

End Sub
Private Sub cmdSAVE_Click()
On Error GoTo ErrorDes

Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString

cmd.ActiveConnection = con
    
    If Len(txtfields(0)) = 0 Then
        Set rs = getdata("select max(cast(VaccineID as int))+ 1 from VaccineInfo")
    If Not rs.EOF Then
        txtfields(0) = IIf(IsNull(rs(0)) = True, "01", Format(rs(0), "00"))
    Else
        txtfields(1) = "01"
    End If
    End If
    
    If Len(txtfields(1)) = 0 Then
        MsgBox "Please Enter Vaccine Name.", vbCritical, App.Title
        txtfields(1).SetFocus
        Exit Sub
    End If
   
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Save_VaccineInfo"
    cmd(1) = txtfields(0)
    cmd(2) = Trim(txtfields(1))
    cmd(3) = soft_user
    cmd(4) = Date
   
    cmd.Execute
    MsgBox "Save Successfully.", vbInformation, App.Title
    cmdnew.SetFocus
  
    Call ShowFlexData

Exit Sub
ErrorDes:    MsgBox Err.Description, vbCritical, App.Title

End Sub

Private Sub ShowFlexData()
On Error GoTo errdes

Dim rs As New ADODB.Recordset

Set rs = getdata("SELECT VaccineID, VaccineName  FROM VaccineInfo")

If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                
                .TextMatrix(i, 0) = "" & rs!VaccineID
                .TextMatrix(i, 1) = "" & rs!VaccineName

            rs.MoveNext
           i = i + 1
        Loop
    End With
 Else
     MSFlexGrid1.Rows = 1

 End If

Exit Sub
errdes: MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo errdes
txtfields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)


Exit Sub
errdes: MsgBox Err.Description, vbInformation, App.Title

End Sub



