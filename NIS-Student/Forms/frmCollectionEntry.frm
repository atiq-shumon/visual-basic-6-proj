VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCollection_info 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   180
      TabIndex        =   7
      Top             =   2640
      Width           =   2865
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   -30
      TabIndex        =   21
      Top             =   7080
      Width           =   11055
      Begin VB.TextBox txtfields 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   12
         Left            =   4410
         TabIndex        =   49
         Top             =   360
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H8000000C&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8070
         TabIndex        =   14
         ToolTipText     =   "Click to insert new information"
         Top             =   240
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
         Height          =   375
         Left            =   5880
         TabIndex        =   13
         ToolTipText     =   "Click to insert new information"
         Top             =   240
         Width           =   945
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H8000000C&
         Caption         =   "Edit Receipt"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6855
         TabIndex        =   12
         ToolTipText     =   "Click to Save"
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H8000000C&
         Caption         =   "Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9045
         TabIndex        =   15
         ToolTipText     =   "Click to Delete"
         Top             =   240
         Width           =   945
      End
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
         Height          =   375
         Left            =   10020
         TabIndex        =   16
         ToolTipText     =   "Click to Close"
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   270
         TabIndex        =   52
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acc_code"
         Height          =   195
         Index           =   13
         Left            =   3750
         TabIndex        =   50
         Top             =   270
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblFeeTitle 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   270
         TabIndex        =   40
         Top             =   300
         Width           =   2835
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   5850
         Top             =   210
         Width           =   5145
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   885
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   10965
      TabIndex        =   18
      Top             =   -60
      Width           =   11025
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Show Dues"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   9570
         TabIndex        =   54
         Top             =   450
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Collection Information Entry"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEF0F7&
         Height          =   270
         Left            =   3510
         TabIndex        =   51
         Top             =   240
         Width           =   3150
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   -90
         Picture         =   "frmCollectionEntry.frx":0000
         Stretch         =   -1  'True
         Top             =   30
         Width           =   11055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   6255
      Left            =   -30
      TabIndex        =   17
      Top             =   780
      Width           =   11055
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000C&
         Caption         =   "Generate MR No"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9000
         TabIndex        =   6
         ToolTipText     =   "Click to Save"
         Top             =   1140
         Width           =   1545
      End
      Begin VB.Frame Frame5 
         Height          =   1245
         Left            =   3240
         TabIndex        =   37
         Top             =   4800
         Width           =   7275
         Begin VB.TextBox txtfields 
            BackColor       =   &H00FFFFFF&
            Height          =   795
            Index           =   1
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   7035
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            Height          =   195
            Index           =   7
            Left            =   150
            TabIndex        =   39
            Top             =   120
            Width           =   630
         End
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   315
         ItemData        =   "frmCollectionEntry.frx":CEA5
         Left            =   4680
         List            =   "frmCollectionEntry.frx":CECD
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "Fees Head(Ctrl+H)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4485
         Left            =   120
         TabIndex        =   29
         Top             =   1560
         Width           =   3075
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   285
         Left            =   8790
         TabIndex        =   3
         Top             =   210
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   300
         Left            =   2610
         Picture         =   "frmCollectionEntry.frx":CF0D
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1140
         Width           =   420
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H80000018&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   25
         Top             =   690
         Width           =   5835
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   690
         Width           =   1545
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H80000018&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   23
         Top             =   1140
         Width           =   4305
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   0
         Left            =   1080
         MaxLength       =   80
         TabIndex        =   5
         ToolTipText     =   "Insert Student Id"
         Top             =   1140
         Width           =   1545
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   1080
         TabIndex        =   0
         Top             =   195
         Width           =   1545
      End
      Begin VB.Frame Frame4 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2955
         Left            =   3240
         TabIndex        =   30
         Top             =   1560
         Width           =   7785
         Begin VB.TextBox txtfields 
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   630
            TabIndex        =   45
            Text            =   "0"
            ToolTipText     =   "Insert Marks "
            Top             =   360
            Width           =   3225
         End
         Begin VB.TextBox txtfields 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   10
            Left            =   60
            TabIndex        =   43
            Top             =   360
            Width           =   555
         End
         Begin VB.TextBox txtfields 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   9
            Left            =   6990
            TabIndex        =   41
            Top             =   2460
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.CommandButton Command2 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7290
            TabIndex        =   11
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtfields 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   6540
            MaxLength       =   7
            TabIndex        =   35
            Text            =   "0"
            ToolTipText     =   "Insert Marks "
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   7
            Left            =   5730
            MaxLength       =   7
            TabIndex        =   10
            Text            =   "0"
            ToolTipText     =   "Insert Discount"
            Top             =   360
            Width           =   795
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   5
            Left            =   4860
            MaxLength       =   7
            TabIndex        =   9
            Text            =   "0"
            ToolTipText     =   "Insert Fine"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   3
            Left            =   3840
            MaxLength       =   7
            TabIndex        =   8
            Text            =   "0"
            ToolTipText     =   "Insert Marks "
            Top             =   360
            Width           =   1005
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   2265
            Left            =   30
            TabIndex        =   53
            Top             =   660
            Width           =   7665
            _ExtentX        =   13520
            _ExtentY        =   3995
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fee Title"
            Height          =   195
            Index           =   10
            Left            =   1560
            TabIndex        =   46
            Top             =   150
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Serial #"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   44
            Top             =   180
            Width           =   540
         End
         Begin VB.Label Label4 
            Caption         =   "seq"
            Height          =   15
            Index           =   0
            Left            =   7290
            TabIndex        =   42
            Top             =   1860
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total "
            Height          =   195
            Index           =   8
            Left            =   6870
            TabIndex        =   34
            Top             =   150
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Discount"
            Height          =   195
            Index           =   6
            Left            =   5790
            TabIndex        =   33
            Top             =   150
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fine "
            Height          =   195
            Index           =   3
            Left            =   5100
            TabIndex        =   32
            Top             =   150
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Actual Amount"
            Height          =   195
            Index           =   4
            Left            =   3810
            TabIndex        =   31
            Top             =   150
            Width           =   1035
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Index           =   12
         Left            =   9840
         TabIndex        =   48
         Top             =   4560
         Width           =   75
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   9210
         TabIndex        =   47
         Top             =   4530
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month Year"
         Height          =   195
         Index           =   9
         Left            =   3570
         TabIndex        =   36
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Collection Date"
         Height          =   195
         Index           =   1
         Left            =   7650
         TabIndex        =   28
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class  Name"
         Height          =   195
         Index           =   5
         Left            =   3570
         TabIndex        =   27
         Top             =   705
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Id"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   24
         Top             =   750
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name"
         Height          =   195
         Index           =   1
         Left            =   3570
         TabIndex        =   22
         Top             =   1170
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Id"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   20
         Top             =   1230
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt. No"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mnuDel 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmCollection_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
 Dim rs1 As New ADODB.Recordset
     If Len(txtfields(0)) = 0 Then
                MsgBox "Fee Code Mandatory", vbInformation, App.Title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(txtfields(1)) = 0 Then
                MsgBox "Fee Title Mandatory", vbInformation, App.Title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            Set rs1 = getdata("SELECT fee_code from fee_info WHERE (fee_code= '" & txtfields(0) & "')")
            
            If rs1.EOF Then
               MsgBox "No such Fee Code Exists..Please Verify.", vbInformation, cmp
               cmdnew.SetFocus
               Exit Sub
            End If
            
                     
             If MaskEdBox1 = "__/__/__" Then
                MsgBox "Please Put a Valid Date. ", vbInformation, cmp
                MaskEdBox1.SetFocus
                Exit Sub
            End If
        If Len(Trim(txtfields(1))) <> 0 Then
           If (MsgBox("Are You sure to delete ?", vbYesNo + vbCritical) = vbYes) Then
                Dim rs As New ADODB.Recordset
                Dim cmd As New ADODB.Command
                Dim con As New ADODB.connection
                con.Open GConnString
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "Fee_info_Save"
                cmd(1) = "d"
                cmd(2) = Format(Trim(txtfields(0)), "00")
                cmd(3) = Trim(txtfields(1))
                cmd(4) = soft_user
                cmd(5) = Format(Date, "DD MMM YYYY")
                cmd.Execute
                MsgBox "Deleted successfully.", vbInformation, cmp
                
                cmdnew.SetFocus
         End If
End If

End Sub
Private Sub cmdEdit_Click()
   Dim rs1 As New ADODB.Recordset
     If Len(txtfields(0)) = 0 Then
                MsgBox "Fee Code Mandatory", vbInformation, App.Title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(txtfields(1)) = 0 Then
                MsgBox "Fee Title Mandatory", vbInformation, App.Title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            Set rs1 = getdata("SELECT fee_code from fee_info WHERE (fee_code= '" & txtfields(0) & "')")
            
            If rs1.EOF Then
               MsgBox "No such Fee Code Exists..Please Verify.", vbInformation, cmp
               cmdnew.SetFocus
               Exit Sub
            End If
            
            
                Dim rs As New ADODB.Recordset
                Dim cmd As New ADODB.Command
                Dim con As New ADODB.connection
                con.Open GConnString
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "Fee_info_Save"
                cmd(1) = "u"
                cmd(2) = Format(Trim(txtfields(0)), "00")
                cmd(3) = Trim(txtfields(1))
                cmd(4) = soft_user
                cmd(5) = Format(Date, "DD MMM YYYY")
                cmd.Execute
                MsgBox "Edited successfully.", vbInformation, cmp
                
                cmdnew.SetFocus

End Sub
Private Sub cmdExit_Click()
   Unload Me
End Sub
Private Sub cmdnew_Click()
  Dim rs2 As New ADODB.Recordset
  Set rs2 = getdata("select max(C_srl) from Collec_master")
  If Not rs2.EOF Then
       txtfields(6).Text = rs2(0)
  End If
              
txtfields(3) = 0
txtfields(6) = ""
txtfields(0) = ""
txtfields(2) = ""
txtfields(4) = ""
txtfields(5) = ""
txtfields(9) = ""
txtfields(10) = ""
txtfields(11) = ""
'MaskEdBox1.Text = Format(Date, "dd/mm/yy")
Label3(12).Caption = 0

Combo1.Clear
get_Mr_info
load_Fee
load_class
Combo1.SetFocus

cmdsave.Enabled = True
Set rs2 = Nothing
End Sub

Private Sub cmdPrint_Click()
  Screen.MousePointer = vbHourglass
   rptMode = 2
   frmViewer.Show 1
End Sub
Private Sub cmdSAVE_Click()
            Dim rs As New ADODB.Recordset
            Dim cmd As New ADODB.Command
            Dim con As New ADODB.connection
              
            If Len(txtfields(0)) = 0 Then
                MsgBox "Student Id Mandatory ", vbInformation, App.Title
                Exit Sub
            End If
            
            
            If Len(Trim(Combo1.Text)) = 0 Then
                MsgBox "Please select a valid Class ID", vbInformation, App.Title
                Combo1.SetFocus
                Exit Sub
            End If
            
            If Len(txtfields(6)) = 0 Then
                MsgBox "Please put a valid Receipt No ", vbInformation, App.Title
                Exit Sub
            End If

             If MaskEdBox1 = "__/__/__" Then
                MsgBox "Please Put a Valid Date. ", vbInformation, cmp
                MaskEdBox1.SetFocus
                Exit Sub
            End If
            
           If Val(Mid(MaskEdBox1.Text, 7, 8)) <> Val(Mid(Combo3, 3, 4)) Then
               MsgBox "Academic Year Conflicts with the date year you have entered", vbInformation, cmp
               Combo3.SetFocus
              Exit Sub
          End If
          
         
          Set rs = getdata("select C_srl from Collec_master where c_srl='" & txtfields(6).Text & "'")
          If rs.EOF Then
            MsgBox "Invalid Money Receipt No...Please Verify", vbInformation, cmp
            txtfields(0).SetFocus
            Exit Sub
          End If
          Set rs = Nothing
            
                con.Open GConnString
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "Collec_master_Save"
                cmd(1) = "u"
                cmd(2) = Val(Trim(txtfields(6)))
                cmd(3) = 0
                cmd(4) = Trim(Combo2.Text)
                cmd(5) = Trim(Combo3.Text)
                cmd(6) = Format(MaskEdBox1, "dd mmm yyyy")
                cmd(7) = Trim(txtfields(1))
                cmd(8) = soft_user
                cmd(9) = Format(Date, "DD MMM YYYY")
                cmd(10) = txtfields(0).Text
                cmd(11) = Combo1.Text
              
                cmd.Execute
                MsgBox "Updated successfully.", vbInformation, "Student Management System"
                
                
                cmdnew.SetFocus
                
End Sub

Private Sub dtpic_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdsave.SetFocus
End If
End Sub

Private Sub cmdsearch_Click()
  Dim f As New frmFind
  Set f.OwnerForm = Me
    f.intInputsel = 0
    f.SQLString = "Select a.StudentId,(select  Studentname from studentinfo b where b.studentid=a.studentid) as StudentName from StudentAdmission a where a.approval='Y' and a.admissionCancel='N'and a.serial_no=(select max(serial_no) from studentadmission where studentid=a.studentid and classid='" & Trim(Combo1) & "' )"
  f.Show 1
    txtfields(0).SetFocus
End Sub

Private Sub Combo1_Click()
'   load_Fee
   load_Fee
   load_title
   
End Sub
Private Sub load_amt()
    Dim rs2 As New ADODB.Recordset
    Set rs2 = getdata("SELECT fee_amt from fee_setup WHERE Class_id= '" & Trim(Combo1.Text) & "' and Fee_Code='" & Mid(Trim(List1.Text), 1, 2) & "'")

            If Not rs2.EOF Then
               txtfields(3).Text = rs2!Fee_amt
            End If
End Sub
Private Sub Combo2_Click()
'  load_Fee
'  load_class
  load_Fee_title
  load_amt
End Sub

Private Sub Combo2_GotFocus()
  Label5.Caption = ""
End Sub
Private Sub Command1_Click()
   If Len(txtfields(0)) = 0 Then
      MsgBox "Student Id Mandatory ", vbInformation, App.Title
      Exit Sub
   End If
   If Len(Trim(Combo1.Text)) = 0 Then
      MsgBox "Please select a valid Class ID", vbInformation, App.Title
      Combo1.SetFocus
      Exit Sub
   End If
               
   If MaskEdBox1 = "__/__/__" Then
      MsgBox "Please Put a Valid Date. ", vbInformation, cmp
      MaskEdBox1.SetFocus
      Exit Sub
   End If
            
    Set rs = getdata("Select a.StudentId,(select  Studentname from studentinfo b where b.studentid=a.studentid) as StudentName from StudentAdmission a where a.approval='Y' and a.admissionCancel='N'and a.studentid='" & Trim(txtfields(0)) & "' and a.serial_no=(select max(serial_no) from studentadmission where studentid=a.studentid and classid='" & Trim(Combo1) & "')")
              If rs.EOF Then
                 MsgBox "This ID may be for another class please verify", vbInformation, cmp
                 txtfields(0).SetFocus
                 Exit Sub
   End If
         
         
   If Val(Mid(MaskEdBox1.Text, 7, 8)) <> Val(Mid(Combo3, 3, 4)) Then
       MsgBox "Academic Year Conflicts with the date year you have entered", vbInformation, cmp
       Combo3.SetFocus
       Exit Sub
   End If
                If con.State = 0 Then
                 con.Open GConnString
                End If
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "Collec_master_Save"
                cmd(1) = "s"
                cmd(2) = Val(Trim(txtfields(6)))
                cmd(3) = 0
                cmd(4) = Trim(Combo2.Text)
                cmd(5) = Trim(Combo3.Text)
                cmd(6) = Format(MaskEdBox1, "dd mmm yyyy")
                cmd(7) = Trim(txtfields(1))
                cmd(8) = soft_user
                cmd(9) = Format(Date, "DD MMM YYYY")
                cmd(10) = txtfields(0).Text
                cmd(11) = Combo1.Text
                cmd.Execute
'                MsgBox "Saved successfully.", vbInformation, "Student Management System"
                
                cmdsave.Enabled = False
               
                
                
                Dim rs2 As New ADODB.Recordset
                Set rs2 = getdata("select max(C_srl) from Collec_master")
                If Not rs2.EOF Then
                   txtfields(6).Text = rs2(0)
                End If
                Command1.Enabled = False
                List1.SetFocus

End Sub

Private Sub Command2_Click()
          Dim rs As New ADODB.Recordset
          Dim cmd As New ADODB.Command
          Dim con As New ADODB.connection
           
            If Len(txtfields(0)) = 0 Then
                MsgBox "Student Id Mandatory ", vbInformation, App.Title
                Exit Sub
            End If
            
            
            If Len(Trim(Combo1.Text)) = 0 Then
                MsgBox "Please select a valid Class ID", vbInformation, App.Title
                Combo1.SetFocus
                Exit Sub
            End If
            
             If Len(Trim(Combo3.Text)) = 0 Then
                MsgBox "Please select a Year", vbInformation, App.Title
                Combo3.SetFocus
                Exit Sub
            End If
            
              If Len(Trim(Combo2.Text)) = 0 Then
                MsgBox "Please select a Month", vbInformation, App.Title
                Combo2.SetFocus
                Exit Sub
            End If
            
            If Len(Trim(List1.Text)) = 0 Then
                MsgBox "Please select a valid Fee Code", vbInformation, App.Title
                List1.SetFocus
                Exit Sub
            End If
            
            If Len(Trim(txtfields(3))) = 0 Then
                MsgBox "Fee Amount Required ", vbInformation, App.Title
                txtfields(3).SetFocus
                Exit Sub
            End If
            
           Set rs = getdata("select C_srl from Collec_master where c_srl='" & txtfields(6).Text & "'")
          If rs.EOF Then
            MsgBox "Invalid Money Receipt No...Please Verify", vbInformation, cmp
            txtfields(0).SetFocus
            Exit Sub
          End If
          Set rs = Nothing
      
            
          Set rs = getdata("SELECT fee_code from collec_details WHERE C_Srl= '" & txtfields(6).Text & "' and fee_code= '" & Mid(List1.Text, 1, 2) & "'")

            If Not rs.EOF Then
               MsgBox "Same Fee Code already Exists..Please Verify.", vbInformation, cmp
                List1.SetFocus
               Exit Sub
            End If
            Set rs = Nothing

                con.Open GConnString
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "Collec_sub_save"
                cmd(1) = "S"
                cmd(2) = Val(Trim(txtfields(9)))
                cmd(3) = Val(Trim(txtfields(6)))
                cmd(4) = soft_user
                cmd(5) = Trim(txtfields(0).Text)
                cmd(6) = Mid(List1.Text, 1, 2)
                cmd(7) = txtfields(3)
                cmd(8) = txtfields(5)
                cmd(9) = txtfields(7)
                cmd.Execute
                
            If Len(txtfields(9).Text) = 0 Then
             Dim rs_set As New ADODB.Recordset
                Set rs_set = getdata("select max(serial_no) from collec_details where C_Srl='" & Trim(txtfields(6)) & "'")
            If Not rs_set.EOF Then
                 txtfields(9).Text = rs_set(0)
            End If
          End If
        Mr_flush
        
        show_sum_mr
        txtfields(11).Text = ""
        txtfields(3).Text = 0
        txtfields(5).Text = 0
        txtfields(7).Text = 0
        txtfields(8).Text = 0
        txtfields(10).Text = 0
        lblFeeTitle.Caption = ""
        List1.SetFocus
        
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  
   If KeyAscii = 13 Then
      SendKeys (Chr(9))
   ElseIf KeyAscii = 27 Then
      Unload Me
    ElseIf KeyAscii = 6 Then
      cmdsearch.SetFocus
   ElseIf KeyAscii = 8 Then
      List1.SetFocus
   ElseIf KeyAscii = 14 Or KeyAscii = 32 Then
      cmdnew_Click
   ElseIf KeyAscii = 19 Then
      cmdsave.SetFocus
    End If
  
End Sub
Private Sub load_Fee()
Dim rs As New ADODB.Recordset
List1.Clear
Set rs = getdata("Select a.fee_code,(select fee_title from fee_info b where a.fee_code=b.fee_code)from Fee_setup a where a.Class_id='" & Trim(Combo1.Text) & "'")
If Not rs.EOF Then
    Do Until rs.EOF
        List1.AddItem rs(0) + "-" + rs(1)
        rs.MoveNext
    Loop
    
End If
End Sub
Private Sub load_title()
Dim rs As New ADODB.Recordset
Set rs = getdata("Select Classname  from ClassInfo where classid='" & Trim(Combo1.Text) & "'")
If Not rs.EOF Then
  txtfields(4).Text = rs(0)
End If
End Sub
Private Sub load_Fee_title()
Dim rs As New ADODB.Recordset
Set rs = getdata("Select Fee_title  from Fee_Info where Fee_code='" & Trim(Combo2.Text) & "'")
If Not rs.EOF Then
  txtfields(5).Text = rs(0)
End If
End Sub

Private Sub Form_Load()
''Dim rs As New adodb.Recordset
''Set rs = GetData("select max (Fee_code+1)from fee_info")
''If Not rs.EOF Then
''    txtfields(0) = IIf(IsNull(rs(0)) = True, "01", Format(rs(0), "00"))
''Else
''    txtfields(0) = "01"
''End If
Label5.Caption = ""
With MSFlexGrid1
    .Rows = 1
    .Cols = 7
    .Col = 0: .Text = "Srl"
    .Col = 1: .Text = "Code"
    .Col = 2: .Text = "Title"
    .Col = 3: .Text = "Amount"
    .Col = 4: .Text = "Fine "
    .Col = 5: .Text = "Discount"
    .Col = 6: .Text = "Total"
    .ColWidth(0) = 300
    .ColWidth(1) = 300
    .ColWidth(2) = 3200
    .ColWidth(3) = 1000
    .ColWidth(4) = 900
    .ColWidth(5) = 750
    .ColWidth(6) = 720
'    .ColWidth(6) = 200

End With
'Call ShowFlexData
load_Fee
load_class
generate_yr
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title

End Sub
Private Sub generate_yr()
   Dim i As Integer
   Dim j As Integer
    j = 2000
   For i = 0 To 50
      Combo3.AddItem j
      
     j = j + 1
      
  Next i
 Combo3.Text = Format(Date, "YYYY")
 Combo2.Text = Format(Date, "MMM")
End Sub
Private Sub load_class()
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Combo1.Clear
Set rs1 = getdata("Select ClassId from ClassInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
        Combo1.AddItem rs1(0)
        rs1.MoveNext
    Loop
   
'    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
End If

End Sub

Private Sub Label6_Click()
  frmDuesinfo.Show 1
End Sub

Private Sub List1_Click()
  load_amt
  load_caption
End Sub
Private Sub load_caption()
   lblFeeTitle.Caption = Mid(List1.Text, 4, 60)
   txtfields(11).Text = Mid(List1.Text, 4, 60)
End Sub

Private Sub MaskEdBox1_GotFocus()
              MaskEdBox1.SelStart = 0
             MaskEdBox1.SelLength = Len(MaskEdBox1)
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then

    If MaskEdBoxDate1 <> "__/__/__" Then
            If Check_ValidDate(MaskEdBox1) = False Then
                MaskEdBox1.SetFocus
                Exit Sub
            End If
    End If
    
    
 End If
End Sub

Private Sub mnuDel_Click()
          If Len(txtfields(0)) = 0 Then
                MsgBox "Student Id Mandatory ", vbInformation, App.Title
                Exit Sub
            End If
            
            
            If Len(Trim(Combo1.Text)) = 0 Then
                MsgBox "Please select a valid Class ID", vbInformation, App.Title
                Combo1.SetFocus
                Exit Sub
            End If
            
'            If Len(Trim(List1.Text)) = 0 Then
'                MsgBox "Please select a valid Fee Code", vbInformation, App.Title
'                List1.SetFocus
'                Exit Sub
'            End If
            If Len(Trim(txtfields(3))) = 0 Then
                MsgBox "Fee Amount Required ", vbInformation, App.Title
                txtfields(3).SetFocus
                Exit Sub
            End If
            
            
'            Dim rs1 As New adodb.Recordset
'             Set rs1 = GetData("SELECT fee_code from fee_info WHERE (fee_code= '" & txtfields(0) & "')")
'
'            If Not rs1.EOF Then
'               MsgBox "Same Fee Code already Exists..Please Verify.", vbInformation, cmp
'               Exit Sub
'            End If
'
                Dim rs As New ADODB.Recordset
                Dim cmd As New ADODB.Command
                Dim con As New ADODB.connection
                con.Open GConnString
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "Collec_sub_save"
                cmd(1) = "P"
                cmd(2) = Val(Trim(txtfields(9)))
                cmd(3) = Val(Trim(txtfields(6)))
                cmd(4) = soft_user
                cmd(5) = Trim(txtfields(0).Text)
                cmd(6) = Mid(List1.Text, 1, 2)
                cmd(7) = txtfields(3)
                cmd(8) = txtfields(5)
                cmd(9) = txtfields(7)
                cmd.Execute
                Mr_flush
                show_sum_mr
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
       PopupMenu mnuDelete, 2
       
       show_sum_mr
    End If
  
End Sub

Private Sub txtfields_Change(Index As Integer)
            Select Case Index
                   Case 3
                        If Not IsNumeric(txtfields(3).Text) Then
                               txtfields(3) = 0
                            Exit Sub
                        End If
                         txtfields(8) = (Val(txtfields(3)) + Val(txtfields(5))) - Val(txtfields(7))
                   Case 5
                        If Not IsNumeric(txtfields(5).Text) Then
                               txtfields(5) = 0
                               Exit Sub
                        End If
                        
                        txtfields(8) = (Val(txtfields(3)) + Val(txtfields(5))) - Val(txtfields(7))
                   Case 6
                        If Len(txtfields(6)) > 0 Then
                          Command1.Enabled = False
                        Else
                          Command1.Enabled = True
                        End If
                   Case 7
                        If Not IsNumeric(txtfields(7).Text) Then
                               txtfields(7) = 0
                               Exit Sub
                        End If
                       txtfields(8) = (Val(txtfields(3)) + Val(txtfields(5))) - Val(txtfields(7))
            End Select
End Sub
Private Sub txtfields_GotFocus(Index As Integer)
  Select Case Index
         Case Index
              txtfields(Index).SelStart = 0
              txtfields(Index).SelLength = Len(txtfields(Index).Text)
  End Select
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
           Case 0
                If Len(txtfields(0).Text) = 0 Then
                  cmdsearch.SetFocus
                  Exit Sub
               Else
                  If Len(txtfields(0)) <> 0 Then
                       If Mid(txtfields(0), 1, 3) <> "STI" Then
                           txtfields(0).Text = Format(txtfields(0), "000000")
                           txtfields(0).Text = "STI-" + txtfields(0).Text
                   End If
              End If
              
              Set rs = getdata("Select a.StudentId from StudentAdmission a where a.approval='Y' and a.admissionCancel='N'and a.studentid='" & Trim(txtfields(0)) & "' and a.serial_no=(select max(serial_no) from studentadmission where studentid=a.studentid and classid='" & Trim(Combo1) & "')")
              If rs.EOF Then
                 MsgBox "This ID may be for another class..... please verify", vbInformation, cmp
                 txtfields(0) = ""
                 txtfields(0).SetFocus
                 Exit Sub
              End If

              
                Set rs = getdata("select StudentName from studentinfo where studentid='" & txtfields(0).Text & "'")
                  If Not rs.EOF Then
                    txtfields(2).Text = "" & rs!StudentName
                  End If
               
              End If
                            
          Case 6
                
                get_Mr_info
                show_sum_mr
                Set rs = getdata("select StudentName from studentinfo where studentid='" & txtfields(0).Text & "'")
                If Not rs.EOF Then
                     txtfields(2).Text = "" & rs!StudentName
                End If
          End Select
   End If
End Sub
Private Sub show_sum_mr()
On Error GoTo ERR_DESC
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT Sum(Act_amount),sum(Fine),sum(Discount) from collec_details where C_Srl='" & Trim(txtfields(6).Text) & "'")
If Not rs.EOF Then
   Label3(12).Caption = (rs(0) + rs(1)) - rs(2)
End If

Exit Sub
ERR_DESC:
   Label3(12).Caption = 0
End Sub
Private Sub get_Mr_info()
On Error Resume Next
    Dim rs As New ADODB.Recordset
    Set rs = getdata("select a.Std_id,a.Class_id,a.Mon,a.Yr,a.Remark,a.collec_date from  collec_master a where a.C_Srl='" & txtfields(6).Text & "'")
    If Not rs.EOF Then
      Label5.Caption = ""
      txtfields(0) = rs(0)
      Combo1.Text = rs(1)
      Combo2.Text = rs(2)
      Combo3.Text = rs(3)
      txtfields(1) = "" & rs(4)
      MaskEdBox1.Text = Format(rs(5), "dd/mm/yy")
      Mr_flush
    Else
'      MaskEdBox1.Text = "__/__/__"
      txtfields(0).Text = ""
      txtfields(2).Text = ""
      Mr_flush
      
        Label5.Caption = "N.B. No such MR no exists...Press NEW to generate"
                   
   End If
End Sub
Private Sub Mr_flush()
  On Error GoTo errdes
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT a.serial_no,a.Fee_code,(select Fee_title from  fee_info where fee_code=a.fee_code) as Fee_title ,a.Act_amount,a.Fine,a.Discount from collec_details a where a.C_Srl='" & txtfields(6).Text & "'")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = rs!serial_no
                .TextMatrix(i, 1) = rs!Fee_Code
                .TextMatrix(i, 2) = rs!fee_title
                .TextMatrix(i, 3) = rs!Act_amount
                .TextMatrix(i, 4) = rs!Fine
                .TextMatrix(i, 5) = rs!Discount
                .TextMatrix(i, 6) = (rs!Act_amount + rs!Fine) - rs!Discount
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

Public Function getdata(SQLString As String) As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = SQLString

 Set rs = cmd.Execute
Set getdata = rs
End Function


'Private Sub ShowFlexData()
'On Error GoTo errdes
'Dim rs As New adodb.Recordset
'Set rs = getdata("SELECT srl_no,Fee_code,Act_amount,Fine,Discount from collec_details where C_Srl='" & Trim(txtfields(6).Text) & "'")
'If Not rs.EOF Then
'    i = 1
'    With MSFlexGrid1
'        Do Until rs.EOF
'            MSFlexGrid1.Rows = i + 1
'                .Rows = i + 1
'                .TextMatrix(i, 0) = rs!srl_no
'                .TextMatrix(i, 1) = rs!Fee_code
'                .TextMatrix(i, 2) = "rs!fee_title"
'                .TextMatrix(i, 3) = rs!Act_amount
'                .TextMatrix(i, 4) = rs!Fine
'                .TextMatrix(i, 5) = rs!Discount
'                .TextMatrix(i, 6) = (rs!Act_amount + rs!Fine) - rs!Discount
'                i = i + 1
'            rs.MoveNext
'        Loop
'    End With
'Else
'    MSFlexGrid1.Rows = 1
' End If
'Exit Sub
'errdes:
'MsgBox Err.Description, vbInformation, App.Title
'End Sub
Private Sub MSFlexGrid1_Click()

On Error GoTo errdes
txtfields(10).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfields(9).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfields(12) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
txtfields(11) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
txtfields(3) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
txtfields(5) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
txtfields(7) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
txtfields(8) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
Exit Sub
errdes:
  MsgBox Err.Description, vbInformation, App.Title
End Sub

