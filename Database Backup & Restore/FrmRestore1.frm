VERSION 5.00
Begin VB.Form FrmRestore1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4050
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   -30
      TabIndex        =   4
      Top             =   -60
      Width           =   4125
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1590
         TabIndex        =   5
         Top             =   150
         Width           =   780
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   315
      Index           =   1
      Left            =   2820
      TabIndex        =   3
      ToolTipText     =   "Press to Close"
      Top             =   1800
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Index           =   0
      Left            =   1710
      TabIndex        =   2
      ToolTipText     =   "Press  to OK"
      Top             =   1800
      Width           =   1065
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Immediate Restore"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Index           =   1
      Left            =   1020
      TabIndex        =   1
      Top             =   1110
      Value           =   -1  'True
      Width           =   2775
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Initial Setup"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Index           =   0
      Left            =   1020
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Index           =   1
      Left            =   -30
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Index           =   0
      Left            =   -120
      Top             =   570
      Width           =   4185
   End
   Begin VB.Shape Shape1 
      Height          =   435
      Left            =   1650
      Top             =   1740
      Width           =   2295
   End
End
Attribute VB_Name = "FrmRestore1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SFillSysObj As New Scripting.FileSystemObject
Dim tx As TextStream
Dim txsql As TextStream
Private Sub Command1_Click(Index As Integer)
   Select Case Index
          Case 0
            If Option1(0).Value = True Then
              If Not SFillSysObj.FileExists("C:\\EXPORT11AM.bat") Then
                     With SFillSysObj
                           .CreateTextFile ("C:\\EXPORT11AM.bat")
                            Set tx = .OpenTextFile("C:\\EXPORT11AM.bat", ForWriting)
                       End With
                               
                       tx.WriteLine ("@echo off")
                       tx.WriteLine ("CD C:\")
                       tx.WriteLine ("SQLPLUS  system/hansaworld @C:\\tmpSQL.SQL")
                       tx.WriteLine ("Exit")
                       tx.Close
              End If ''''end of export11AM.bat
              
              If Not SFillSysObj.FolderExists("D:\databases") Then
                     MkDir ("D:\databases")
                 
                     
                   On Error Resume Next
                   If Not SFillSysObj.FileExists("C:\\tmpSQL.sql") Then
                       With SFillSysObj
                           .CreateTextFile ("C:\\tmpSQL.sql")
                            Set txsql = .OpenTextFile("C:\\tmpsql.sql", ForWriting)
                       End With
                               
                       txsql.WriteLine ("conn system/hansaworld")
                       txsql.WriteLine ("CREATE TABLESPACE NAT_PMIS DATAFILE 'D:\databases\NAT_PMIS'")
                       txsql.WriteLine ("SIZE 512M")
                       txsql.WriteLine ("AUTOEXTEND ON MAXSIZE UNLIMITED;")
                       txsql.WriteLine ("CREATE TABLESPACE NAT_ACCT_06_07 DATAFILE 'D:\databases\ NAT_ACCT_06_07'")
                       txsql.WriteLine ("SIZE 512M")
                       txsql.WriteLine ("AUTOEXTEND ON  MAXSIZE UNLIMITED;")
                       txsql.WriteLine ("CREATE TABLESPACE NAT_BILL DATAFILE 'D:\databases\NAT_BILL'")
                       txsql.WriteLine ("SIZE 512M")
                       txsql.WriteLine ("AUTOEXTEND ON MAXSIZE UNLIMITED;")
                 End If
                 txsql.Close
                End If
               
               Shell ("C:\\EXPORT11AM.bat")
               
               If SFillSysObj.FileExists("D:\databases\NAT_PMIS") And SFillSysObj.FileExists("D:\databases\ NAT_ACCT_06_07") And SFillSysObj.FileExists("D:\databases\NAT_BILL") Then
'                 MsgBox "Initial Setup Successfull", vbInformation, "Congratulations...."
                 Exit Sub
               End If
                
               
            End If
            
          
            
      If Option1(1).Value = True Then  ''''Immediate Restore]
'         If SFillSysObj.FolderExists("D:\databases") Then
                    frmRestore2.Show 1
'         Else
'            MsgBox "Please Complete Initial Setup first", vbInformation, cmp
'            Option1(0).SetFocus
'       End If
     End If
                
    Case 1
                   
          Unload Me
          
   End Select
End Sub

Private Sub Form_Load()
  If SFillSysObj.FolderExists("F:\databases") Then
    Option1(0).Enabled = False
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If SFillSysObj.FileExists("C:\\EXPORT11AM.bat") Then
          SFillSysObj.DeleteFile ("C:\\EXPORT11AM.bat")
        End If
        If SFillSysObj.FileExists("C:\\tmpsql.sql") Then
            SFillSysObj.DeleteFile ("C:\\tmpsql.sql")
        End If
End Sub

