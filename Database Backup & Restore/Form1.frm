VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5610
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   ":::"
      Height          =   345
      Left            =   3360
      TabIndex        =   3
      Top             =   2520
      Width           =   675
   End
   Begin VB.TextBox txtDataFileName 
      Height          =   345
      Left            =   750
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2610
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Step1"
      Height          =   495
      Left            =   1950
      TabIndex        =   1
      Top             =   1110
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   405
      Left            =   1800
      TabIndex        =   0
      Top             =   1770
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SFillSysObj As New Scripting.FileSystemObject
Dim tx As TextStream
Private Sub Command1_Click()
On Error Resume Next
Kill ("D:\test_databases\tmp.sql")
If Not SFillSysObj.FileExists("D:\test_databases\tmp.sql") Then
       With SFillSysObj
            .CreateTextFile ("D:\test_databases\tmp.sql")
           Set tx = .OpenTextFile("D:\test_databases\tmp.sql", ForWriting)
            
       End With
       
   tx.WriteLine ("conn system/hansaworld")
   tx.WriteLine ("drop user acct cascade;")
   tx.WriteLine ("create user acct identified by dn_acct")
   tx.WriteLine ("default tablespace acct_tabspace;")
   tx.WriteLine ("grant connect,resource,dba to acct;")
   tx.WriteLine ("conn acct/dn_acct@dcc ;")
   tx.Write ("$imp acct/dn_acct@dcc file=")
   tx.Write (txtDataFileName.Text)
   tx.Write (" fromuser=acct touser=acct;")
End If
'Open ("D:\test_databases\tmp.txt") For Random As 2
'Set tx = SFillSysObj.OpenTextFile("D:\test_databases\tmp.txt", ForWriting)

'tmp.WriteLine (Text1)
End Sub

Private Sub Command2_Click()
On Error Resume Next
 RmDir ("D:\test_databases")
 MkDir ("D:\test_databases")
End Sub


Private Sub Command3_Click()
'  On Error GoTo ErrHandler:
    
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "All Files (*.*)|*.*|Backup Files (*.dmp)|*.dmp"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.InitDir = gBkupRstrFilePath
    CommonDialog1.DefaultExt = "dmp"
    CommonDialog1.DialogTitle = "Browse:File Location"
    CommonDialog1.Action = 1
    txtDataFileName.Text = CommonDialog1.FileName
    
    If Len(Trim(txtDataFileName.Text)) > 0 Then
        Command1.Enabled = True
    End If
   
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbInformation, "Backup_Restore"
End Sub

Private Sub Command4_Click()
        With SFillSysObj
            .CreateTextFile ("D:\tmp.bat")
           Set tx = .OpenTextFile("D:\tmp.bat", ForWriting)
            
       End With
   tx.Write ("Batch")
End Sub
