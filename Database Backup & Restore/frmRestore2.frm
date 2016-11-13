VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRestore2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   ":::"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   14
      ToolTipText     =   "Press to Browse DMP file"
      Top             =   2910
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   ":::"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3930
      TabIndex        =   13
      ToolTipText     =   "Press to Browse DMP file"
      Top             =   2400
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Diagnostic Management System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Index           =   4
      Left            =   60
      TabIndex        =   12
      Top             =   2850
      Width           =   3855
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Store and Inventory Management System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Index           =   3
      Left            =   60
      TabIndex        =   11
      Top             =   2340
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   30
      Top             =   3540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   ":::"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3930
      TabIndex        =   10
      ToolTipText     =   "Press to Browse DMP file"
      Top             =   1890
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   ":::"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3930
      TabIndex        =   9
      ToolTipText     =   "Press to Browse DMP file"
      Top             =   1350
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   ":::"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3930
      TabIndex        =   8
      ToolTipText     =   "Press to Browse DMP file"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtDataFileName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   30
      TabIndex        =   7
      Top             =   3330
      Width           =   5325
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   -30
      TabIndex        =   5
      Top             =   -60
      Width           =   5595
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2"
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
         Left            =   2130
         TabIndex        =   6
         Top             =   150
         Width           =   780
      End
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Financial Accounting Management System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Index           =   2
      Left            =   60
      TabIndex        =   4
      Top             =   1830
      Width           =   3855
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Personnel Management System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   810
      Value           =   -1  'True
      Width           =   3855
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Billing Management System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   1320
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Index           =   0
      Left            =   3120
      TabIndex        =   1
      ToolTipText     =   "Presss to OK"
      Top             =   3720
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   315
      Index           =   1
      Left            =   4230
      TabIndex        =   0
      ToolTipText     =   "Press to Close"
      Top             =   3720
      Width           =   1065
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Index           =   2
      Left            =   -60
      Top             =   1800
      Width           =   5655
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Index           =   0
      Left            =   -120
      Top             =   780
      Width           =   6225
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Index           =   1
      Left            =   -30
      Top             =   1290
      Width           =   5685
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   3090
      Top             =   3690
      Width           =   2265
   End
End
Attribute VB_Name = "frmRestore2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SFillSysObj As New Scripting.FileSystemObject
Dim tx As TextStream
Dim txBatch As TextStream

Private Sub Command1_Click(Index As Integer)
  On Error Resume Next
          Select Case Index
   
          Case 0
               
               If Len(txtDataFileName) = 0 Then
                  MsgBox "Please select a valid DMP file name", vbInformation, cmp
                  Command2(1).SetFocus
                  Exit Sub
               End If
               If SFillSysObj.FileExists("C:\\tmpSQl.sql") Then
                  SFillSysObj.DeleteFile ("C:\\tmpSQL.sql")
              End If
              If Option1(0).Value = True Then '''payroll
                   On Error Resume Next
                   
                   If Not SFillSysObj.FileExists("C:\\tmpSQL.sql") Then
                         With SFillSysObj
                              .CreateTextFile ("C:\\tmpSQL.sql")
                               Set tx = .OpenTextFile("C:\\tmpSQL.sql", ForWriting)
                                    
                         End With
                               
                           tx.WriteLine ("conn system/hansaworld")
                           tx.WriteLine ("drop user payroll cascade;")
                           tx.WriteLine ("create user payroll identified by payroll")
                           tx.WriteLine ("default tablespace NAT_PMIS;")
                           tx.WriteLine ("grant connect,resource,dba to payroll;")
                           tx.WriteLine ("conn payroll/payroll ;")
                           tx.Write ("$imp payroll/payroll@bank file=")
                           tx.Write (txtDataFileName.Text)
                           tx.Write (" fromuser=payroll touser=payroll;")
                         tx.Close
                           
                     '''for batch file
                         If Not SFillSysObj.FileExists("C:\\EXPORT11AM.bat") Then
                               With SFillSysObj
                                    .CreateTextFile ("C:\\EXPORT11AM.bat")
                                   Set txBatch = .OpenTextFile("C:\\EXPORT11AM.bat", ForWriting)
                                    
                               End With
                        
                           txBatch.WriteLine ("@echo off")
                           txBatch.WriteLine ("CD C:\")
                           txBatch.WriteLine ("SQLPLUS system/hansaworld @C:\\tmpSQL.sql")
                           txBatch.Write ("EXIT")
                           txBatch.Close
                  End If
                  Dim i As Integer
                  Shell ("C:\\EXPORT11AM.bat")
                  '''MsgBox i
                  'MsgBox "Imported Successfully", vbInformation, "Congratulation..."
                End If
                
               End If
                
                 If Option1(1).Value = True Then '''Billing
                   On Error Resume Next
                   If Not SFillSysObj.FileExists("C:\\tmpSQl.sql") Then
                               With SFillSysObj
                                    .CreateTextFile ("C:\\tmpSql.sql")
                                   Set tx = .OpenTextFile("C:\\tmpSQL.sql", ForWriting)
                                    
                               End With
                               
                           tx.WriteLine ("conn system/hansaworld")
                           tx.WriteLine ("drop user Hospital_billing cascade;")
                           tx.WriteLine ("create user hospital_billing identified by dn_medical_hospital")
                           tx.WriteLine ("default tablespace NAT_BILL;")
                           tx.WriteLine ("grant connect,resource,dba to hospital_billing;")
                           tx.WriteLine ("conn hospital_billing/dn_medical_hospital ;")
                           tx.Write ("$imp hospital_billing/dn_medical_hospital@bank file=")
                           tx.Write (txtDataFileName.Text)
                           tx.Write (" fromuser=hospital_billing touser=hospital_billing;")
                           tx.Close
'
                         '''for batch file
                        
                         If Not SFillSysObj.FileExists("C:\\EXPORT11AM.bat") Then
                               With SFillSysObj
                                    .CreateTextFile ("C:\\EXPORT11AM.bat")
                                   Set txBatch = .OpenTextFile("C:\\EXPORT11AM.bat", ForWriting)
                                    
                               End With
                        
                           txBatch.WriteLine ("@echo off")
                           txBatch.WriteLine ("CD C:\")
                           txBatch.WriteLine ("SQLPLUS system/hansaworld @C:\\tmpSQL.sql")
                           txBatch.Write ("EXIT")
                           txBatch.Close
                    End If
  
                     Shell ("C:\\EXPORT11AM.bat")
                    ''' MsgBox "Imported Successfully", vbInformation, "Congratulation..."
                End If
                
             End If
                
                If Option1(2).Value = True Then '''accounting
                   On Error Resume Next
                  If Not SFillSysObj.FileExists("C:\\tmpSQl.sql") Then
                               With SFillSysObj
                                    .CreateTextFile ("C:\\tmpSql.sql")
                                   Set tx = .OpenTextFile("C:\\tmpSQL.sql", ForWriting)
                                    
                               End With
                 
                               
                           tx.WriteLine ("conn system/hansaworld")
                           tx.WriteLine ("drop user acct_06_07 cascade;")
                           tx.WriteLine ("create user acct_06_07 identified by dn_acct")
                           tx.WriteLine ("default tablespace NAT_ACCT_06_07;")
                           tx.WriteLine ("grant connect,resource,dba to acct_06_07;")
                           tx.WriteLine ("conn acct_06_07/dn_acct;")
                           tx.Write ("$imp acct_06_07/dn_acct@bank file=")
                           tx.Write (txtDataFileName.Text)
                           tx.Write (" fromuser=acct_06_07 touser=acct_06_07;")
                           tx.Close
                           '''for batch file
                         If Not SFillSysObj.FileExists("C:\\EXPORT11AM.bat") Then
                               With SFillSysObj
                                    .CreateTextFile ("C:\\EXPORT11AM.bat")
                                   Set txBatch = .OpenTextFile("C:\\EXPORT11AM.bat", ForWriting)
                                    
                               End With
                        
                           txBatch.WriteLine ("@echo off")
                           txBatch.WriteLine ("CD C:\")
                           txBatch.WriteLine ("SQLPLUS system/hansaworld @C:\\tmpSQL.sql")
                           txBatch.Write ("EXIT")
                           txBatch.Close
                  End If
     
                 
                     Shell ("C:\\EXPORT11AM.bat")
                    ''' MsgBox "Imported Successfully", vbInformation, "Congratulation..."
                End If
           End If
           
               If Option1(3).Value = True Then '''Inventory
                   On Error Resume Next
                  If Not SFillSysObj.FileExists("C:\\tmpSQl.sql") Then
                               With SFillSysObj
                                    .CreateTextFile ("C:\\tmpSql.sql")
                                   Set tx = .OpenTextFile("C:\\tmpSQL.sql", ForWriting)
                                    
                               End With
                 
                               
                           tx.WriteLine ("conn system/hansaworld")
                           tx.WriteLine ("drop user national_inventory cascade;")
                           tx.WriteLine ("create user national_inventory identified by dn_inventory")
                           tx.WriteLine ("default tablespace NATIONAL_inventory;")
                           tx.WriteLine ("grant connect,resource,dba to national_inventory;")
                           tx.WriteLine ("conn national_inventory/dn_inventory;")
                           tx.Write ("$imp national_inventory/dn_inventory@bank file=")
                           tx.Write (txtDataFileName.Text)
                           tx.Write (" fromuser=national_inventory touser=national_inventory;")
                           tx.Close
                           '''for batch file
                         If Not SFillSysObj.FileExists("C:\\EXPORT11AM.bat") Then
                               With SFillSysObj
                                    .CreateTextFile ("C:\\EXPORT11AM.bat")
                                   Set txBatch = .OpenTextFile("C:\\EXPORT11AM.bat", ForWriting)
                                    
                               End With
                        
                           txBatch.WriteLine ("@echo off")
                           txBatch.WriteLine ("CD C:\")
                           txBatch.WriteLine ("SQLPLUS system/hansaworld @C:\\tmpSQL.sql")
                           txBatch.Write ("EXIT")
                           txBatch.Close
                  End If
     
                 
                     Shell ("C:\\EXPORT11AM.bat")
                    ''' MsgBox "Imported Successfully", vbInformation, "Congratulation..."
                End If
           End If

             If Option1(4).Value = True Then '''Pathology
                   On Error Resume Next
                  If Not SFillSysObj.FileExists("C:\\tmpSQl.sql") Then
                               With SFillSysObj
                                    .CreateTextFile ("C:\\tmpSql.sql")
                                   Set tx = .OpenTextFile("C:\\tmpSQL.sql", ForWriting)
                                    
                               End With
                 
                               
                           tx.WriteLine ("conn system/hansaworld")
                           tx.WriteLine ("drop user hospital_pathology cascade;")
                           tx.WriteLine ("create user hospital_pathology identified by dn_hospital_pathology")
                           tx.WriteLine ("default tablespace NATIONAL_pathology;")
                           tx.WriteLine ("grant connect,resource,dba to hospital_pathology;")
                           tx.WriteLine ("conn hospital_pathology/dn_hospital_pathology;")
                           tx.Write ("$imp hospital_pathology/dn_hospital_pathology@bank file=")
                           tx.Write (txtDataFileName.Text)
                           tx.WriteLine (" fromuser=hospital_pathology touser=hospital_pathology;")
                           tx.WriteLine ("conn Hospital_billing/dn_medical_hospital@bank")
                           tx.WriteLine ("grant select on test_info_main to hospital_pathology;")
                           tx.WriteLine ("grant select on test_info_sub to hospital_pathology;")
                           tx.WriteLine ("grant select on indoor_pat_bed_info to hospital_pathology;")
                           tx.WriteLine ("grant select on doctor_info to hospital_pathology;")
                           tx.WriteLine ("grant select on pat_info_main_out_door  to hospital_pathology;")
                           tx.WriteLine ("grant select on pat_info_sub1_out_door to hospital_pathology;")
                           
                          
                           tx.Close
                           '''for batch file
                         If Not SFillSysObj.FileExists("C:\\EXPORT11AM.bat") Then
                               With SFillSysObj
                                    .CreateTextFile ("C:\\EXPORT11AM.bat")
                                   Set txBatch = .OpenTextFile("C:\\EXPORT11AM.bat", ForWriting)
                                    
                               End With
                        
                           txBatch.WriteLine ("@echo off")
                           txBatch.WriteLine ("CD C:\")
                           txBatch.WriteLine ("SQLPLUS system/hansaworld @C:\\tmpSQL.sql")
                           txBatch.Write ("EXIT")
                           txBatch.Close
                  End If
     
                 
                     Shell ("C:\\EXPORT11AM.bat")
                    ''' MsgBox "Imported Successfully", vbInformation, "Congratulation..."
                End If
           End If
  
           
           
           
     Case 1
           If SFillSysObj.FileExists("C:\\tmpSql.sql") Then
                  SFillSysObj.DeleteFile ("C:\\tmpSql.sql")
           End If
            If SFillSysObj.FileExists("C:\\EXPORT11AM.bat") Then
                  SFillSysObj.DeleteFile ("C:\\EXPORT11AM.bat")
           End If
           Unload Me
               
   End Select
End Sub

Private Sub Command2_Click(Index As Integer)
On Error GoTo ErrHandler
  Select Case Index
         Case 0, 1, 2, 3, 4
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "All Files (*.*)|*.*|Backup Files (*.dmp)|*.dmp"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.InitDir = gBkupRstrFilePath
    CommonDialog1.DefaultExt = "dmp"
    CommonDialog1.DialogTitle = "Browse:File Location"
    CommonDialog1.Action = 1
    txtDataFileName.Text = CommonDialog1.FileName
    
    If Len(Trim(txtDataFileName.Text)) > 0 Then
        Command1(0).Enabled = True
    Else
       Command1(0).Enabled = False
    End If
End Select
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbInformation, "Backup_Restore"


  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Unload Me
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   SFillSysObj.DeleteFolder ("C:\tmp_databases")
End Sub

Private Sub Option1_GotFocus(Index As Integer)
            Select Case Index
                   Case 0, 1, 2, 3, 4
                     If SFillSysObj.FileExists("C:\\tmpSql.sql") Then
                  SFillSysObj.DeleteFile ("C:\\tmpSql.sql")
           End If
            If SFillSysObj.FileExists("C:\\EXPORT11AM.bat") Then
                  SFillSysObj.DeleteFile ("C:\\EXPORT11AM.bat")
           End If
                     Option1(Index).ForeColor = vbRed
                   
            End Select
End Sub

Private Sub Option1_LostFocus(Index As Integer)
   Select Case Index
           Case 0, 1, 2, 3, 4
                   Option1(Index).ForeColor = vbBlack
            End Select
End Sub
