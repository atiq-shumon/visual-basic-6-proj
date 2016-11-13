VERSION 5.00
Begin VB.Form frmmain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Backup & Restore"
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   8775
   Icon            =   "frmmain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   -30
      TabIndex        =   0
      Top             =   2370
      Width           =   9285
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Backup and Restore Utility"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   345
         Left            =   2100
         TabIndex        =   1
         Top             =   270
         Width           =   5235
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup"
         Shortcut        =   ^B
      End
      Begin VB.Menu gfdsg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
         Shortcut        =   ^R
      End
      Begin VB.Menu gfdsgdsfgf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SFillSysObj As New Scripting.FileSystemObject

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
         If SFillSysObj.FileExists("C:\\EXPORT11AM.bat") Then
          SFillSysObj.DeleteFile ("C:\\EXPORT11AM.bat")
        End If
        If SFillSysObj.FileExists("C:\\tmpsql.sql") Then
            SFillSysObj.DeleteFile ("C:\\tmpsql.sql")
        End If
   End If
  Unload Me
  End Sub

Private Sub Form_Unload(Cancel As Integer)
  If SFillSysObj.FileExists("C:\\EXPORT11AM.bat") Then
     SFillSysObj.DeleteFile ("C:\\EXPORT11AM.bat")
   End If
   If SFillSysObj.FileExists("C:\\tmpsql.sql") Then
       SFillSysObj.DeleteFile ("C:\\tmpsql.sql")
  End If
End Sub

Private Sub mnuClose_Click()
      Unload Me
End Sub

Private Sub mnuRestore_Click()
  FrmRestore1.Show 1
End Sub
