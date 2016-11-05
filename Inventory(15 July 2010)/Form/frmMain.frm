VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H80000001&
   ClientHeight    =   10830
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   10275
      Left            =   0
      ScaleHeight     =   10215
      ScaleWidth      =   15180
      TabIndex        =   1
      Top             =   660
      Width           =   15240
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Real time stock management"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3390
         TabIndex        =   13
         Top             =   9480
         Width           =   3180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Store && Inventory MIS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   525
         Left            =   1020
         TabIndex        =   12
         Top             =   9000
         Width           =   5520
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IT Division, DNMIH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   270
         Left            =   10500
         TabIndex        =   11
         Top             =   630
         Width           =   2460
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Software Programmer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   9990
         TabIndex        =   10
         Top             =   210
         Width           =   2985
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Powered By:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7950
         TabIndex        =   9
         Top             =   210
         Width           =   1695
      End
      Begin VB.Label lblCategoryCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         Height          =   195
         Left            =   3450
         TabIndex        =   8
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2040
         TabIndex        =   7
         Top             =   960
         Width           =   870
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2040
         TabIndex        =   6
         Top             =   510
         Width           =   870
      End
      Begin VB.Label lblUserId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2040
         TabIndex        =   5
         Top             =   90
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Store Type :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   510
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Id       :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   2
         Top             =   90
         Width           =   1530
      End
      Begin VB.Image Image1 
         Height          =   11520
         Left            =   -330
         Picture         =   "frmMain.frx":014A
         Top             =   -360
         Width           =   15360
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   5370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D5EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E2C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E5E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EA34
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F30E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnusys 
      Caption         =   "[  System  ]"
      Begin VB.Menu smnusysMacReg 
         Caption         =   "Machine Registry"
      End
      Begin VB.Menu dfd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserInfo 
         Caption         =   "User Profile"
         Visible         =   0   'False
      End
      Begin VB.Menu rtyrtyr 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUserPermission 
         Caption         =   "User Permission"
      End
      Begin VB.Menu mmuItemOpening 
         Caption         =   "Product Opening"
         Visible         =   0   'False
      End
      Begin VB.Menu fsdf 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLogoff 
         Caption         =   "Log Off ..."
         Shortcut        =   ^L
      End
      Begin VB.Menu fgfg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEntry 
      Caption         =   "[  Entry  ]"
      Begin VB.Menu mnuIPE 
         Caption         =   "Item Purchase Entry"
         Shortcut        =   ^P
      End
      Begin VB.Menu dfsfgfds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIIE 
         Caption         =   "Item Issue Entry"
         Shortcut        =   ^I
      End
      Begin VB.Menu fgfdg 
         Caption         =   "-"
      End
      Begin VB.Menu mmuItemPurchaseReturn 
         Caption         =   "Item Purchase Return"
      End
      Begin VB.Menu safdsafdsa 
         Caption         =   "-"
      End
      Begin VB.Menu mmuItemIssueReturn 
         Caption         =   "Item Issue Return"
      End
      Begin VB.Menu dfasfdsa 
         Caption         =   "-"
      End
      Begin VB.Menu mmuItemAdjustment 
         Caption         =   "Item Adjustment"
      End
      Begin VB.Menu gfsdgfds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSIE 
         Caption         =   "Special Issue Entry"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuGS 
      Caption         =   "[  General Setup  ]"
      Begin VB.Menu mnuFYS 
         Caption         =   "Fiscal Year Setup"
      End
      Begin VB.Menu dxfdsafas 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAU 
         Caption         =   "Add User"
      End
      Begin VB.Menu mnuSpaceb 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSupplierInformation 
         Caption         =   "Supplier Information"
      End
      Begin VB.Menu sp24 
         Caption         =   "-"
      End
      Begin VB.Menu mnuITI 
         Caption         =   "Issue Type Information"
      End
      Begin VB.Menu sp21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIRI 
         Caption         =   "Item Related information"
         Begin VB.Menu mnuIUI 
            Caption         =   "Item Unit Information"
         End
         Begin VB.Menu mfdsafdsa 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCatagoryInformation 
            Caption         =   "Store Category Information"
         End
         Begin VB.Menu sP27 
            Caption         =   "-"
         End
         Begin VB.Menu mnuITII 
            Caption         =   "Store Item Type Information"
         End
         Begin VB.Menu fdsafdsafds 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGI 
            Caption         =   "Item Group Information"
         End
         Begin VB.Menu dfgdsg 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIIS 
            Caption         =   "Item Information Setup"
         End
      End
      Begin VB.Menu dswqss 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOIE 
         Caption         =   "Opening Balance Entry"
      End
      Begin VB.Menu fdsafdsaf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDB 
         Caption         =   "Data Backup"
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "[  Utility  ]"
      Begin VB.Menu mnuCP 
         Caption         =   "Change Password"
      End
      Begin VB.Menu fdsafsa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCStock 
         Caption         =   "Current Stock"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "[  Report  ]"
      Begin VB.Menu mnuIIR 
         Caption         =   "Statements"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "[  About  ]"
      Begin VB.Menu mnuAboutUs 
         Caption         =   "About Us"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SFillSysObj As New Scripting.FileSystemObject
Dim tx As TextStream

Private Sub MDIForm_Activate()
   Me.Label2.Caption = CategoryTitle
   Me.lblCategoryCode.Caption = CategoryCode
   lblUserId = userid
   lblUserName = userName
   If Len(lblUserId.Caption) = 0 Then
       Unload Me
    End If
 End Sub
Private Sub MDIForm_Load()
   If SFillSysObj.FileExists("C:\\bkp.bat") Then
     SFillSysObj.DeleteFile ("C:\\bkp.bat")
  End If
  If SFillSysObj.FileExists("C:\\tmpSQL.SQL") Then
     SFillSysObj.DeleteFile ("C:\\tmpSQL.SQL")
  End If
   'Label3.Visible = True
  ' Label4.Visible = True
'   Label5.Visible = True
'   Label6(0).Visible = True
'   Label6(1).Visible = True
'   Label6(2).Visible = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  If SFillSysObj.FileExists("C:\\bkp.bat") Then
     SFillSysObj.DeleteFile ("C:\\bkp.bat")
  End If
  If SFillSysObj.FileExists("C:\\tmpSQL.SQL") Then
     SFillSysObj.DeleteFile ("C:\\tmpSQL.SQL")
  End If
End Sub

Private Sub mmuItemAdjustment_Click()
Dim f As New frmAdjustment
f.Show 1
End Sub

Private Sub mmuItemIssueReturn_Click()
 frmIssueReturn.Show 1
End Sub
Private Sub mmuItemPurchaseReturn_Click()
Dim f As New frmPurchaseReturn
f.Show 1
End Sub

Private Sub mnuAboutUs_Click()
 frmAbout.Show 1
End Sub
Private Sub mnuAU_Click()
  Form22.Show 1
End Sub
Private Sub mnuCatagoryInformation_Click()
    frmItemCateInfo.Show 1
End Sub
Private Sub mnuCP_Click()
  Change_pass.Show 1
  End Sub

Private Sub mnuCStock_Click()
   frmCurrentStock.Show 1
End Sub

Private Sub mnuDB_Click()
 On Error GoTo err_desc
  If SFillSysObj.FileExists("C:\\bkp.bat") Then
     SFillSysObj.DeleteFile ("C:\\bkp.bat")
  End If
  If SFillSysObj.FileExists("C:\\tmpSQL.SQL") Then
     SFillSysObj.DeleteFile ("C:\\tmpSQL.SQL")
  End If
   If Not SFillSysObj.FileExists("C:\\tmpSQl.sql") Then
      With SFillSysObj
          .CreateTextFile ("C:\\tmpSQl.sql")
          Set tx = .OpenTextFile("C:\\tmpSQl.sql", ForWriting)
      End With
      tx.WriteLine ("conn system/hansaworld@bank")
      tx.WriteLine ("$EXP USERID=National_inventory/DN_inventory@BANK FILE='G:\BACKUP\Inventory_BACKUP'")
          
  End If
  tx.Close
  If Not SFillSysObj.FileExists("C:\\bkp.bat") Then
                     With SFillSysObj
                           .CreateTextFile ("C:\\bkp.bat")
                            Set tx = .OpenTextFile("C:\\bkp.bat", ForWriting)
                       End With
                               
          tx.WriteLine ("SQLPLUS  national_inventory/dn_inventory @C:\\tmpSQL.SQL")
          tx.WriteLine ("Exit")
          tx.Close
    End If ''''end of export11AM.bat
   
  
 
   Shell "c:\bkp.bat"
   
   
'  SFillSysObj.DeleteFile ("C:\\bkp.bat")
'  SFillSysObj.DeleteFile ("C:\\tmpSQL.SQL")
   Exit Sub
err_desc:
      MsgBox Err.Description, vbCritical, "IT Division, DNMIH"
 End Sub
Private Sub mnuExit_Click()
   If SFillSysObj.FileExists("C:\\bkp.bat") Then
     SFillSysObj.DeleteFile ("C:\\bkp.bat")
  End If
  If SFillSysObj.FileExists("C:\\tmpSQL.SQL") Then
     SFillSysObj.DeleteFile ("C:\\tmpSQL.SQL")
  End If
    Unload Me
    
End Sub
Private Sub mnuGI_Click()
  frmItemgroupInfo.Show 1
End Sub

Private Sub mnuIIE_Click()
  frmIssue.Show 1
End Sub

Private Sub mnuIIR_Click()
Dim f As New frmReport
f.Show 1
End Sub

Private Sub mnuIIS_Click()
  frmItemInfo.Show 1
End Sub

Private Sub mnuIOB_Click()
   rptmode = 2
  rptViewer.Show 1
End Sub

Private Sub mnuIPE_Click()
   frmPurchase.Show 1
End Sub

Private Sub mnuITI_Click()
 frmIssueType.Show 1
End Sub

Private Sub mnuITII_Click()
  frmItemTypeInfo.Show 1
End Sub

Private Sub mnuIUI_Click()
  frmItemUnitInfo.Show 1
End Sub

Private Sub mnuLogoff_Click()
   If SFillSysObj.FileExists("C:\\bkp.bat") Then
     SFillSysObj.DeleteFile ("C:\\bkp.bat")
  End If
  If SFillSysObj.FileExists("C:\\tmpSQL.SQL") Then
     SFillSysObj.DeleteFile ("C:\\tmpSQL.SQL")
  End If
  Unload Me
  frmUserLogOn.Show 1
End Sub

Private Sub mnuOIE_Click()
  frmOpninfo.Show 1
End Sub

Private Sub mnuSIE_Click()
  frmSpecialIssue.Show 1
End Sub

Private Sub mnuUserInfo_Click()
Dim f As New frmUserInfo
f.Show 1
End Sub

Private Sub smnusysMacReg_Click()
    Dim objfrmMacReg As New frmMachineReg
    Load objfrmMacReg
    objfrmMacReg.Show 1
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Index
        Case 1
              mnuIPE_Click
        Case 2
             mnuIIE_Click
        Case 3
             mnuIIR_Click
        Case 4
             mnuCStock_Click
       Case 5
           Unload Me
           
 End Select
End Sub

