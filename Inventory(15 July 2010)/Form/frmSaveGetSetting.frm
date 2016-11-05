VERSION 5.00
Begin VB.Form frmSaveGetSetting 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Machine Registry"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmSaveGetSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "System Configuration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   9
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   7
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Database Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1500
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Backup Server Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Local Server Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1740
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Central Server Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1875
      End
   End
   Begin VB.CommandButton cmd 
      Height          =   495
      Index           =   1
      Left            =   3960
      Picture         =   "frmSaveGetSetting.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Height          =   495
      Index           =   0
      Left            =   3240
      Picture         =   "frmSaveGetSetting.frx":0B24
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   615
   End
End
Attribute VB_Name = "frmSaveGetSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Get_SystemConfigurationInfo()
On Err GoTo Err_Des
    strCentralServerName = GetSetting(strAppName, "Settings", "CentralServerName")
    strLocalServerName = GetSetting(strAppName, "Settings", "LocalServerName")
    strBackupServerName = GetSetting(strAppName, "Settings", "BackupServerName")
    strDatabaseName = GetSetting(strAppName, "Settings", "DatabaseName")
    
    txtFields(0) = strCentralServerName
    txtFields(1) = strLocalServerName
    txtFields(2) = strBackupServerName
    txtFields(3) = strDatabaseName
    
Exit Sub
Err_Des:
MsgBox Err.Description, vbCritical + vbOKOnly, strmsgtitle

End Sub

Private Sub Cmd_Click(Index As Integer)
On Err GoTo Err_Des
Select Case Index
    Case 0 'Save
        If Len(Trim(txtFields(0))) = 0 Then
            MsgBox "Please Insert Central Server Name.", vbInformation + vbOKOnly, strAppName
            txtFields(0).SetFocus
            Exit Sub
        ElseIf Len(Trim(txtFields(1))) = 0 Then
            MsgBox "Please Insert Local Server Name.", vbInformation + vbOKOnly, strAppName
            txtFields(1).SetFocus
            Exit Sub
        ElseIf Len(Trim(txtFields(3))) = 0 Then
            MsgBox "Please Insert Database Name.", vbInformation + vbOKOnly, strAppName
            txtFields(3).SetFocus
            Exit Sub
        End If
    
        SaveSetting strAppName, "Settings", "CentralServerName", Trim(txtFields(0))
        SaveSetting strAppName, "Settings", "LocalServerName", Trim(txtFields(1))
        SaveSetting strAppName, "Settings", "BackupServerName", Trim(txtFields(1))
        SaveSetting strAppName, "Settings", "DatabaseName", Trim(txtFields(3))
        
        Unload Me
    
    
    Case 1 ' Exit
        Unload Me
End Select

Exit Sub
Err_Des:
MsgBox Err.Description, vbCritical + vbOKOnly, strmsgtitle

End Sub

Private Sub Form_Load()
On Err GoTo Err_Des
    Get_SystemConfigurationInfo
Exit Sub
Err_Des:
MsgBox Err.Description, vbCritical + vbOKOnly, strmsgtitle
    
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
txtFields(Index).SelStart = 0
txtFields(Index).SelLength = Len(txtFields(Index))
End Sub

Private Sub txtfields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Select Case Index
        Case 0
            txtFields(1).SetFocus
        Case 1
            txtFields(2).SetFocus
        Case 2
            txtFields(3).SetFocus
        Case 3
            cmd(0).SetFocus
    End Select
End If
End Sub
