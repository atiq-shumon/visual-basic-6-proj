VERSION 5.00
Begin VB.Form frmMachineReg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Machine Registration"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Height          =   495
      Index           =   1
      Left            =   795
      Picture         =   "frmMachineReg.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1725
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Height          =   495
      Index           =   0
      Left            =   75
      Picture         =   "frmMachineReg.frx":081A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1725
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4440
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   2
         Left            =   1725
         TabIndex        =   6
         Top             =   975
         Width           =   2490
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   1
         Left            =   1725
         TabIndex        =   4
         Top             =   600
         Width           =   2490
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   0
         Left            =   1725
         TabIndex        =   2
         Top             =   225
         Width           =   2490
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   5
         Top             =   1020
         Width           =   1260
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   645
         Width           =   690
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Log on Name"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   270
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmMachineReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Click(Index As Integer)
'On Error GoTo Err_Des

Dim objcom As New DSLComFram.CRijndael
Select Case Index
    Case 0 ' save
        If Len(Trim(txtFields(0))) = 0 Then
            MsgBox "Please Insert Log on Name.", vbInformation + vbOKOnly, strmsgtitle
            txtFields(0).SetFocus
            Exit Sub
        ElseIf Len(Trim(txtFields(1))) = 0 Then
            MsgBox "Please Insert Password.", vbInformation + vbOKOnly, strmsgtitle
            txtFields(1).SetFocus
            Exit Sub
        ElseIf Len(Trim(txtFields(2))) = 0 Then
            MsgBox "Please Insert Confirm Password.", vbInformation + vbOKOnly, strmsgtitle
            txtFields(2).SetFocus
            Exit Sub
        End If
        
        If Trim(txtFields(1)) <> Trim(txtFields(2)) Then
            MsgBox "Passwrod and Confirm Password is not Same.", vbInformation + vbOKOnly, strmsgtitle
            txtFields(1) = ""
            txtFields(2) = ""
            txtFields(1).SetFocus
            Exit Sub
        End If
    
        Dim objcmd As New ADODB.Command
        Set objcmd.ActiveConnection = objmyCon
        objcmd.CommandType = adCmdStoredProc
        objcmd.CommandText = "AddSQLBranchUser"
        objcmd(1) = Trim(txtFields(0))
        objcmd(2) = Trim(txtFields(1))
        objcmd(3) = strDatabaseName
        objcmd.Execute
        
        Dim strpass As String
        
               
        Dim bytInpass() As Byte
        Dim bytOutpass() As Byte
        
        bytPass() = "DSL"
        
        bytInpass() = Trim(txtFields(1))
        
        bytOutpass = objcom.EncryptData(bytInpass(), bytPass())
        
        SaveSetting strAppName, "Settings", "DBUser", Trim(txtFields(0))
        SaveSetting strAppName, "Settings", "DUPass", bytOutpass
        MsgBox "System Needs to Restart.", vbInformation + vbOKOnly, strmsgtitle
        
        End
    Case 1
        Unload Me
End Select

Exit Sub
Err_Des:
MsgBox Err.Description, vbCritical + vbOKOnly, strmsgtitle

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmMachineReg = Nothing
End Sub

Private Sub txtfields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'On Error GoTo Err_Des

If KeyCode = vbKeyReturn Then
    Select Case Index
        Case 0
            txtFields(1).SetFocus
        Case 1
            txtFields(2).SetFocus
        Case 2
            cmd(0).SetFocus
    End Select
End If

Exit Sub
Err_Des:
MsgBox Err.Description, vbCritical + vbOKOnly, strmsgtitle

End Sub
