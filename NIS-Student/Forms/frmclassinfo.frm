VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmclassinfo 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8820
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      TabIndex        =   37
      Top             =   6810
      Width           =   8865
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000C&
         Caption         =   "Edit"
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
         Left            =   5850
         MaskColor       =   &H8000000C&
         TabIndex        =   47
         ToolTipText     =   "Click to insert New information"
         Top             =   90
         Width           =   945
      End
      Begin VB.TextBox txtTrackid 
         Height          =   285
         Left            =   690
         TabIndex        =   46
         Top             =   150
         Visible         =   0   'False
         Width           =   1305
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
         Height          =   405
         Left            =   7800
         TabIndex        =   41
         ToolTipText     =   "Click to Exit"
         Top             =   90
         Width           =   945
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
         Left            =   6810
         TabIndex        =   40
         ToolTipText     =   "Click to Delete"
         Top             =   90
         Width           =   945
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
         Left            =   4890
         TabIndex        =   39
         ToolTipText     =   "Click to save "
         Top             =   90
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
         Height          =   405
         Left            =   3930
         MaskColor       =   &H8000000C&
         TabIndex        =   38
         ToolTipText     =   "Click to insert New information"
         Top             =   90
         Width           =   945
      End
      Begin VB.Shape Shape1 
         Height          =   465
         Left            =   3870
         Top             =   60
         Width           =   4905
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   855
      Left            =   30
      ScaleHeight     =   795
      ScaleWidth      =   8775
      TabIndex        =   4
      Top             =   -30
      Width           =   8835
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Class && Section Information"
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
         Height          =   555
         Left            =   600
         TabIndex        =   43
         Top             =   180
         Width           =   6915
      End
      Begin VB.Image Image1 
         Height          =   990
         Left            =   -60
         Picture         =   "frmclassinfo.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   9135
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5985
      Left            =   30
      TabIndex        =   5
      Top             =   840
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   10557
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "Class Setup"
      TabPicture(0)   =   "frmclassinfo.frx":CEA5
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MSFlexGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Section Setup"
      TabPicture(1)   =   "frmclassinfo.frx":CEC1
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "MSFlexGrid2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdsearch"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdsearch 
         Height          =   300
         Left            =   2610
         Picture         =   "frmclassinfo.frx":CEDD
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1530
         Width           =   420
      End
      Begin VB.Frame Frame7 
         Caption         =   "Shift Setup"
         ForeColor       =   &H00C00000&
         Height          =   1125
         Left            =   -74880
         TabIndex        =   31
         Top             =   1170
         Width           =   3795
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmclassinfo.frx":D1BF
            Left            =   180
            List            =   "frmclassinfo.frx":D1C9
            TabIndex        =   1
            Text            =   "Morning-Shift"
            ToolTipText     =   "Select shift"
            Top             =   420
            Width           =   3255
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H8000000B&
         Caption         =   "Class Monitor Information"
         ForeColor       =   &H00C00000&
         Height          =   975
         Left            =   60
         TabIndex        =   27
         Top             =   1350
         Width           =   8685
         Begin VB.CommandButton Command1 
            Height          =   300
            Left            =   2550
            Picture         =   "frmclassinfo.frx":D1E7
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   540
            Width           =   420
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   7
            Left            =   930
            TabIndex        =   20
            ToolTipText     =   "Press ENTER to select Class Monitor"
            Top             =   540
            Width           =   1575
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   6
            Left            =   930
            TabIndex        =   19
            ToolTipText     =   "Press ENTER to select Class Monitor"
            Top             =   210
            Width           =   1575
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Left            =   3120
            TabIndex        =   34
            Top             =   600
            Width           =   420
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Std ID"
            Height          =   195
            Left            =   90
            TabIndex        =   33
            Top             =   540
            Width           =   450
         End
         Begin VB.Label Label12 
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   3690
            TabIndex        =   32
            Top             =   600
            Width           =   4935
         End
         Begin VB.Label Label12 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   3690
            TabIndex        =   30
            Top             =   240
            Width           =   4935
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Left            =   3120
            TabIndex        =   29
            Top             =   270
            Width           =   420
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Std ID"
            Height          =   195
            Left            =   90
            TabIndex        =   28
            Top             =   240
            Width           =   450
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   3645
         Left            =   90
         TabIndex        =   24
         Top             =   2310
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   6429
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorSel    =   12640511
         ForeColorSel    =   16711680
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Frame Frame6 
         Caption         =   "Section Information"
         ForeColor       =   &H00C00000&
         Height          =   1035
         Left            =   60
         TabIndex        =   14
         Top             =   330
         Width           =   8685
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   315
            Left            =   7260
            TabIndex        =   44
            Top             =   270
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd-mmm-yy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cmdTeacherID 
            Height          =   315
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   630
            Width           =   975
         End
         Begin VB.TextBox txtfields 
            BackColor       =   &H00CEF0F7&
            Height          =   285
            Index           =   5
            Left            =   3000
            TabIndex        =   25
            ToolTipText     =   "Insert Class Teacher Name"
            Top             =   630
            Width           =   5595
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   4
            Left            =   5490
            TabIndex        =   17
            ToolTipText     =   "Insert Room No"
            Top             =   270
            Width           =   615
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   3
            Left            =   3000
            TabIndex        =   16
            ToolTipText     =   "Insert Section Name"
            Top             =   270
            Width           =   1875
         End
         Begin VB.TextBox txtfields 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   870
            TabIndex        =   15
            Top             =   270
            Width           =   945
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Effective Date "
            Height          =   195
            Left            =   6150
            TabIndex        =   45
            Top             =   300
            Width           =   1065
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teacher ID "
            Height          =   195
            Index           =   1
            Left            =   30
            TabIndex        =   42
            Top             =   660
            Width           =   855
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class Teacher"
            Height          =   195
            Index           =   0
            Left            =   1890
            TabIndex        =   26
            Top             =   660
            Width           =   1020
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Room#"
            Height          =   195
            Left            =   4920
            TabIndex        =   23
            Top             =   300
            Width           =   525
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Section Name"
            Height          =   195
            Left            =   1890
            TabIndex        =   22
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Section ID"
            Height          =   195
            Left            =   30
            TabIndex        =   21
            Top             =   300
            Width           =   840
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Time Setup"
         ForeColor       =   &H00C00000&
         Height          =   1155
         Left            =   -69960
         TabIndex        =   10
         Top             =   1170
         Width           =   3675
         Begin MSComCtl2.DTPicker dtpCloseTime 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "h:nn AM/PM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   4
            EndProperty
            Height          =   345
            Left            =   1530
            TabIndex        =   3
            ToolTipText     =   "Insert Closing Time"
            Top             =   690
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   609
            _Version        =   393216
            Format          =   50135042
            CurrentDate     =   38613
         End
         Begin MSComCtl2.DTPicker dtpStTime 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "h:nn AM/PM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   4
            EndProperty
            Height          =   345
            Left            =   1530
            TabIndex        =   2
            ToolTipText     =   "Insert Start Time"
            Top             =   210
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   609
            _Version        =   393216
            Format          =   50135042
            CurrentDate     =   38613
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Closing Time"
            Height          =   195
            Left            =   180
            TabIndex        =   12
            Top             =   690
            Width           =   900
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Starting Time"
            Height          =   195
            Left            =   150
            TabIndex        =   11
            Top             =   240
            Width           =   930
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000B&
         Height          =   795
         Left            =   -74940
         TabIndex        =   6
         Top             =   360
         Width           =   8685
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   1
            Left            =   3270
            MaxLength       =   80
            TabIndex        =   0
            ToolTipText     =   "Insert Class Name"
            Top             =   300
            Width           =   5205
         End
         Begin VB.TextBox txtfields 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   900
            TabIndex        =   7
            Top             =   270
            Width           =   1305
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class Name"
            Height          =   195
            Left            =   2310
            TabIndex        =   9
            Top             =   330
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class ID"
            Height          =   195
            Left            =   180
            TabIndex        =   8
            Top             =   300
            Width           =   585
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3585
         Left            =   -74910
         TabIndex        =   13
         Top             =   2340
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   6324
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   8421504
         BackColorFixed  =   -2147483626
         BackColorSel    =   12640511
         ForeColorSel    =   -2147483635
         BackColorBkg    =   -2147483636
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmclassinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
On Error GoTo errdes
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
If SSTab1.Tab = 0 Then
    If (MsgBox("Are You sure to delete ?", vbYesNo + vbCritical) = vbYes) Then
        Set rs = getdata("select * from SectionInfo where(ClassID = '" & Trim(txtfields(0)) & "') ")
            If rs.EOF Then
                cmd.CommandType = adCmdText
                cmd.CommandText = "Delete from ClassInfo  where (ClassID = '" & Trim(txtfields(0)) & "') "
                cmd.Execute
                MsgBox "Delete successfully Class Information.", vbInformation, App.Title
                txtfields(0) = ""
                txtfields(1) = ""
                Combo1.Text = " "
                dtpCloseTime.Value = "00:00:00"
                dtpStTime.Value = "00:00:00"
                Call ShowFlexData
            Else
                MsgBox "Section of this class has to remove 1st", vbCritical, App.Title
                Exit Sub
            End If
    Else
        Exit Sub
    End If
Else  '''Tab-1        i.e. Section Information
    If Len(txttrackid) = 0 Then
       MsgBox "Please Select an information by clicking on the grid", vbInformation, cmp
       Exit Sub
    End If
    If (MsgBox("Are You sure to delete ?", vbYesNo + vbCritical) = vbYes) Then
        cmd.CommandType = adCmdText
        cmd.CommandText = "Delete from SectionHistoryInfo  where (trackid = '" & Trim(txttrackid) & "')"
'        and sectionId= '" & Trim(txtFields(2)) & "'
        cmd.Execute
         If MSFlexGrid2.Row = 1 Then
            Set rs = getdata("select SectionID  from StudentAttendanceLeaveInfo where ClassID='" & Trim(txtfields(0)) & "' and  sectionId= '" & Trim(txtfields(2)) & "'")
            If Not rs.EOF Then
               MsgBox "This section is Already used...You can't delete", vbInformation, cmp
               Exit Sub
            End If
            
            cmd.CommandType = adCmdText
            cmd.CommandText = "Delete from SectionInfo  where (ClassID = '" & Trim(txtfields(0)) & "') and sectionId= '" & Trim(txtfields(2)) & "'"
            cmd.Execute
         End If
            
        MsgBox "Delete successfully Section Information.", vbInformation, App.Title
        txtfields(2) = ""
        txtfields(3) = ""
        txtfields(4) = ""
        txtfields(5) = ""
        txtfields(6) = ""
        txtfields(7) = ""
        Label12(6) = ""
        Label12(7) = ""
        Call ShowFlexData1
    Else
        Exit Sub
    End If
End If
errdes:
If Err Then MsgBox Err.Description
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
If SSTab1.Tab = 0 Then
    Set rs = getdata("select max(ClassID)+ 1 from ClassInfo")
    If Not rs.EOF Then
        txtfields(0) = IIf(IsNull(rs(0)) = True, "00001", Format(rs(0), "00000"))
    Else
        txtfields(0) = "00001"
    End If
    txtfields(1) = ""
    Combo1.Text = " "
    dtpCloseTime.Value = "00:00:00"
    dtpStTime.Value = "00:00:00"
    txtfields(1).SetFocus
End If
If SSTab1.Tab = 1 Then
    Set rs = getdata("select max(SectionID)+ 1 from SectionInfo where ClassID='" & Trim(txtfields(0)) & "' ")
    If Not rs.EOF Then
        txtfields(2) = IIf(IsNull(rs(0)) = True, "00001", Format(rs(0), "00000"))
    Else
        txtfields(2) = "00001"
    End If
    txtfields(3) = ""
    txtfields(4) = ""
    txtfields(5) = ""
    txtfields(6) = ""
    txtfields(7) = ""
    Label12(6) = ""
    Label12(7) = ""
    txtfields(3).SetFocus
End If

End Sub

Private Sub cmdSAVE_Click()
If Len(txtfields(1)) = 0 And Len(Combo1) = 0 Then Exit Sub
If Len(txtfields(0)) = 0 Then
    MsgBox "Please Enter Class Id.", vbInformation, App.Title
    cmdnew.SetFocus
    Exit Sub
End If
If Len(txtfields(1)) = 0 Then
    MsgBox "Please Enter Class Name.", vbInformation, App.Title
    txtfields(1).SetFocus
    Exit Sub
End If
'If Trim(Combo1.Text) <> Trim("Morning-Shift") Or Trim(Combo1.Text) <> Trim("Day-Shift") Then
'    MsgBox "Please Enter Valid Shift Name.", vbInformation, App.Title
'    Combo1.SetFocus
'    Exit Sub
'End If
If dtpCloseTime.Hour = 0 Then
    MsgBox "Please enter Ending Time", vbInformation, "School Management Sysytem"
    dtpCloseTime.SetFocus
    Exit Sub
End If
If dtpStTime.Hour = 0 Then
    MsgBox "Please enter Strating Time", vbInformation, "School Management Sysytem"
    dtpStTime.SetFocus
    Exit Sub
End If
If dtpCloseTime.Value = dtpStTime.Value Then
    MsgBox "Please enter valid Strating and Ending  Time", vbCritical, "School Management Sysytem"
    dtpStTime.SetFocus
    Exit Sub
End If
Dim rs As New ADODB.Recordset
Set rs = getdata("select classname from classinfo where classname ='" & txtfields(1) & "' and shiftname ='" & Combo1.Text & "' and ClassId <>'" & txtfields(0) & "'")
If Not rs.EOF Then
    MsgBox "Same Class Name for Same Shift can't be inserted Twice .", vbCritical, "School Management Sysytem"
    txtfields(1) = ""
    Exit Sub
End If
If SSTab1.Tab = 1 Then
    If Len(txtfields(3)) = 0 And Len(txtfields(6)) = 0 Then Exit Sub
    If Len(txtfields(3)) = 0 Then
        MsgBox "Please Enter Section Name.", vbInformation, App.Title
        txtfields(3).SetFocus
        Exit Sub
    End If
     If Len(txtfields(4)) = 0 Then
        MsgBox "Please Enter Room No.", vbInformation, App.Title
        txtfields(4).SetFocus
        Exit Sub
    End If
   If Len(cmdTeacherID) = 0 Then
     MsgBox "Please Put Valid teacher ID", vbInformation, cmp
     cmdTeacherID.SetFocus
     Exit Sub
    End If
  End If
 

'Dim rs As New adodb.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
If SSTab1.Tab = 0 Then
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "ClassInformation"
    cmd(1) = Format(Trim(txtfields(0)), "00000")
    cmd(2) = Trim(txtfields(1))
    cmd(3) = Combo1.Text
    cmd(4) = dtpStTime.Value
    cmd(5) = dtpCloseTime.Value
    cmd.Execute
    Call ShowFlexData
End If
If SSTab1.Tab = 1 Then
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SectionInformation"
    cmd(1) = 1
    cmd(2) = Format(Trim(txtfields(2)), "00000")
    cmd(3) = Trim(txtfields(3))
    cmd(4) = Trim(txtfields(0))
    cmd(5) = Trim(txtfields(4))
    cmd(6) = Trim(cmdTeacherID)
    cmd(7) = Trim(txtfields(6))
    cmd(8) = Trim(txtfields(7))
    cmd(9) = Date
    cmd(10) = soft_user
    cmd(11) = Trim(Format(MaskEdBox1.Text, "dd/mm/yyyy"))
    cmd.Execute
    Call ShowFlexData1
End If
MsgBox "Information Save successfully.", vbInformation, "Student Management System"
cmdnew.SetFocus
End Sub
Private Sub cmdsearch_Click()
Dim f As New frmFind
Set f.OwnerForm = Me
    f.intInputsel = 6
    f.SQLString = "Select a.StudentId,(select  Studentname from studentinfo b where b.studentid=a.studentid) as StudentName from StudentAdmission a where a.approval='Y' and a.admissionCancel='N'and a.serial_no=(select max(serial_no) from studentadmission where studentid=a.studentid and classid='" & txtfields(0) & "')"
    f.Show 1
    txtfields(6).SetFocus
End Sub
Private Sub cmdSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim f As New frmFind
Set f.OwnerForm = Me
    f.intInputsel = 6
    f.SQLString = "Select a.StudentId,(select  Studentname from studentinfo b where b.studentid=a.studentid) as StudentName from StudentAdmission a where a.approval='Y' and a.admissionCancel='N'and a.serial_no=(select max(serial_no) from studentadmission where studentid=a.studentid)"
    f.Show 1
End If
End Sub
Private Sub cmdTeacherID_Click()
   load_teacher_title
End Sub
Private Sub cmdTeacherID_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdsearch.SetFocus
   End If
End Sub
Private Sub Combo1_Change()
  Combo1.Text = "Morning-Shift"
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    dtpStTime.SetFocus
End If
End Sub
Private Sub Command1_Click()
Dim f As New frmFind

Set f.OwnerForm = Me
    f.intInputsel = 7
   f.SQLString = "Select a.StudentId,(select  Studentname from studentinfo b where b.studentid=a.studentid) as StudentName from StudentAdmission a where a.approval='Y' and a.admissionCancel='N'and a.serial_no=(select max(serial_no) from studentadmission where studentid=a.studentid and classid='" & txtfields(0) & "')"
   f.Show 1
'    txtfields(7).SetFocus

End Sub
Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim f As New frmFind
Set f.OwnerForm = Me
    f.intInputsel = 7
    f.SQLString = "Select a.StudentId,(select  Studentname from studentinfo b where b.studentid=a.studentid) as StudentName from StudentAdmission a where a.approval='Y' and a.admissionCancel='N'and a.serial_no=(select max(serial_no) from studentadmission where studentid=a.studentid)"
    f.Show 1
End If
End Sub
Private Sub Command2_Click()
      Dim rs As New ADODB.Recordset
      Dim cmd As New ADODB.Command
      Dim con As New ADODB.connection
      con.Open GConnString
      cmd.ActiveConnection = con

        cmd.CommandType = adCmdText
        cmd.CommandText = "Update  SectionHistoryInfo set Sectiondsc='" & Trim(txtfields(3).Text) & "' ,SecMonitor1='" & Trim(txtfields(6).Text) & "' ,secmonitor2='" & Trim(txtfields(7).Text) & "',ClassTeacher='" & Trim(cmdTeacherID.Text) & "',EffectiveDate='" & Format(MaskEdBox1, "dd mmm yyyy") & "' where (trackid = '" & Trim(txttrackid) & "')"
'        and sectionId= '" & Trim(txtFields(2)) & "'
        cmd.Execute
         If MSFlexGrid2.Row = 1 Then
            cmd.CommandType = adCmdText
             cmd.CommandText = "Update  SectionInfo set Sectiondsc='" & Trim(txtfields(3).Text) & "' ,SecMonitor1='" & Trim(txtfields(6).Text) & "' ,secmonitor2='" & Trim(txtfields(7).Text) & "',ClassTeacher='" & Trim(cmdTeacherID.Text) & "',EffectiveDate='" & Format(MaskEdBox1, "dd mmm yyyy") & "' where (ClassID = '" & Trim(txtfields(0).Text) & "') and SectionID='" & Trim(txtfields(2).Text) & "'"
            cmd.Execute
         End If
       MsgBox "Updated successfully", vbInformation, cmp
End Sub

Private Sub dtpCloseTime_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdsave.SetFocus
End If
End Sub
Private Sub dtpStTime_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    dtpCloseTime.SetFocus
End If
End Sub
Private Sub Form_Load()

'On Error GoTo errdes
SSTab1.Tab = 0
Dim rs As New ADODB.Recordset
Set rs = getdata("select max(ClassID)+ 1 from ClassInfo")
    If Not rs.EOF Then
        txtfields(0) = IIf(IsNull(rs(0)) = True, "00001", Format(rs(0), "00000"))
    Else
        txtfields(0) = "00001"
    End If
'    txtfields(1) = ""
'    txtfields(1).SetFocus
With MSFlexGrid1
    .Rows = 1
    .Cols = 5
    .Col = 0: .Text = "                     ID #"
    .Col = 1: .Text = "Class Name"
    .Col = 2: .Text = "Shift Name"
    .Col = 3: .Text = "Start Time"
    .Col = 4: .Text = "End Time"
    .ColWidth(0) = 1400
    .ColWidth(1) = 4740
    .ColWidth(2) = 4000
    .ColWidth(2) = 0
    .ColWidth(2) = 0
    
    
End With
With MSFlexGrid2
    .Rows = 1
    .Cols = 11
    .Col = 0: .Text = "       ID #"
    .Col = 1: .Text = "Section  Name"
    .Col = 2: .Text = "Room No"
    .Col = 3: .Text = " Class ID"
    .Col = 4: .Text = " Class Teacher Name"
    .Col = 5: .Text = " Monitor's(1st) ID"
    .Col = 6: .Text = " Monitor's(1st) Name"
    .Col = 7: .Text = " Monitor's(Alternative) ID"
    .Col = 8: .Text = " Monitor's(Alternative) Name"
    .ColAlignment(9) = 0
    .Col = 9: .Text = " Effective Date"
    .ColWidth(0) = 1200
    .ColWidth(1) = 3000
    .ColWidth(2) = 1500
    .ColWidth(3) = 0
    .ColWidth(4) = 2000
    .ColWidth(5) = 0
    .ColWidth(6) = 2000
    .ColWidth(7) = 0
    .ColWidth(8) = 2000
    .ColWidth(9) = 2200
    .ColWidth(10) = 0
 End With
Call ShowFlexData
load_teacher

Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title


End Sub
Private Sub load_teacher()
  cmdTeacherID.Clear
  Dim rs As New ADODB.Recordset
  Set rs = getdata("SELECT Emp_id  FROM  Emp_Per_Info")
  If Not rs.EOF Then
     Do Until rs.EOF
       cmdTeacherID.AddItem Trim(rs(0))
       rs.MoveNext
     Loop
   End If
     
End Sub
Private Sub load_teacher_title()
   Dim rs As New ADODB.Recordset
   Set rs = getdata("SELECT Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Name From Emp_Per_Info where Emp_Per_Info.Emp_id='" & Trim(cmdTeacherID) & "'")
   If Not rs.EOF Then
     txtfields(5).Text = "" & rs!name
   End If
   
End Sub
Private Function get_name(mode As Integer, tec_id As String) As String
   Dim name As String
   Dim rs As New ADODB.Recordset
  If mode = 1 Then
       Set rs = getdata("SELECT Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Name From Emp_Per_Info where Emp_Per_Info.Emp_id='" & Trim(tec_id) & "'")
       If Not rs.EOF Then
         name = "" & rs!name
       Else
         name = ""
       End If
 ElseIf mode = 2 Then
    Set rs = getdata("SELECT studentname as name From studentinfo where studentid='" & Trim(tec_id) & "'")
     If Not rs.EOF Then
        name = "" & rs!name
      Else
        name = ""
     End If
  
 End If
   get_name = name
 End Function

Private Sub MaskEdBox1_GotFocus()
     MaskEdBox1.SelStart = 0
     MaskEdBox1.SelLength = Len(MaskEdBox1.Text)
End Sub
Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If MaskEdBox1 <> "__/__/__" Then
            If Check_ValidDate(MaskEdBox1) = False Then
                MaskEdBox1.SetFocus
                Exit Sub
            End If
    End If
    cmdTeacherID.SetFocus
 End If
End Sub
Private Sub MSFlexGrid1_Click()
On Error GoTo errdes
txtfields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
Combo1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
dtpStTime.Value = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
dtpCloseTime = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
txtfields(2) = ""
txtfields(3) = ""
txtfields(4) = ""
txtfields(5) = ""
txtfields(6) = ""
txtfields(7) = ""
Call ShowFlexData1
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub MSFlexGrid1_SelChange()
  MSFlexGrid1_Click
End Sub

Private Sub MSFlexGrid2_Click()
On Error GoTo errdes
If MSFlexGrid2.Row > 0 Then
    txtfields(2) = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 0)
    txtfields(3) = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1)
    txtfields(4) = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2)
    cmdTeacherID = Trim(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 3))
    txtfields(6) = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 5)
    txtfields(7) = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 7)
    MaskEdBox1 = Format(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 9), "dd/mm/yy")
    txttrackid = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 10)


Dim rs2 As New ADODB.Recordset
  If Len(txtfields(6)) > 0 Then
  Set rs2 = getdata("select StudentName from StudentInfo where StudentID='" & Trim(txtfields(6)) & "'")
  If Not rs2.EOF Then
     Label12(6).Caption = "" & rs2!StudentName
   End If
 End If

 Dim rs1 As New ADODB.Recordset
   If Len(txtfields(7)) > 0 Then
   Set rs1 = getdata("select StudentName from StudentInfo where StudentID='" & Trim(txtfields(7)) & "'")
   If Not rs1.EOF Then
      Label12(7).Caption = "" & rs1!StudentName
      End If
   End If
End If '''''''''end of if msflexgrid2.Row>0
Exit Sub

errdes:
 MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub MSFlexGrid2_SelChange()
  MSFlexGrid2_Click
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then
    Dim rs As New ADODB.Recordset
    Set rs = getdata("select max(SectionID)+ 1 from SectionInfo where ClassID='" & Trim(txtfields(0)) & "' ")
    If Not rs.EOF Then
        txtfields(2) = IIf(IsNull(rs(0)) = True, "00001", Format(rs(0), "00000"))
    Else
        txtfields(2) = "00001"
            
    End If
       Call ShowFlexData1
       txtfields(3).SetFocus
End If
End Sub
Private Sub txtfields_GotFocus(Index As Integer)
   Select Case Index
          Case 6
              Dim rs2 As New ADODB.Recordset
               If Len(txtfields(6)) > 0 Then
                    Set rs2 = getdata("select StudentName from StudentInfo where StudentID='" & Trim(txtfields(6)) & "'")
                    If Not rs2.EOF Then
                        Label12(6).Caption = "" & rs2!StudentName
                    End If
               End If

           Case 7
              Dim rs1 As New ADODB.Recordset
              If Len(txtfields(7)) > 0 Then
                Set rs1 = getdata("select StudentName from StudentInfo where StudentID='" & Trim(txtfields(7)) & "'")
                If Not rs1.EOF Then
                    Label12(7).Caption = "" & rs1!StudentName
                End If
              End If
   End Select
End Sub
Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
Dim rs As New ADODB.Recordset
If KeyAscii = 13 Then
    Select Case Index
        Case 1
            Combo1.SetFocus
        Case 3
            txtfields(4).SetFocus
        Case 4
            MaskEdBox1.SetFocus
            
            
        Case 5
            txtfields(6).SetFocus
        Case 6
            txtfields(7).SetFocus
       Case 7
          cmdsave.SetFocus
'            If Len(Trim(txtfields(6))) = 0 Then
'                cmdsearch_Click
'            Else
'                Set rs = getdata("select StudentID from StudentInfo where StudentID='" & Trim(txtfields(6)) & "'")
'                If Not rs.EOF Then
'                    Set rs = getdata("select StudentName from StudentInfo where StudentID='" & Trim(txtfields(6)) & "'")
'                    If Not rs.EOF Then
'                        Label12(6).Caption = "" & rs!StudentName
'                    End If
'                    txtfields(7).SetFocus
'                Else
'                    cmdsearch_Click
'                End If
'            End If
'        Case 7
'            If Len(Trim(txtfields(7))) = 0 Then
'                Command1_Click
'            Else
'                Set rs = getdata("select StudentID from StudentInfo where StudentID='" & Trim(txtfields(7)) & "'")
'                If Not rs.EOF Then
'                    cmdSave.SetFocus
'                Else
'                    Command1_Click
'                End If
'            End If
   End Select
    
End If
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
Private Sub txtfields_LostFocus(Index As Integer)
Dim rs As New ADODB.Recordset
Select Case Index
    Case 0
        If Len(Trim(txtfields(0))) = 0 Then Exit Sub
      
            txtfields(0) = Format(txtfields(0), "00000")
          
            Set rs = getdata("SELECT ClassName from ClassInfo WHERE (classID = '" & Trim(txtfields(0)) & "')")
                 If Not rs.EOF Then
                        txtfields(1) = rs!ClassName
                        
            End If
'    Case 6
'            Dim rs2 As New ADODB.Recordset
'
'            Set rs2 = getdata("select StudentName from StudentInfo where StudentID='" & Trim(txtfields(6)) & "'")
'            If Not rs2.EOF Then
'                Label12(6).Caption = "" & rs2!StudentName
'            End If
'            rs2.Close
'    Case 7
'            Dim rs1 As New ADODB.Recordset
'
'            Set rs1 = getdata("select StudentName from StudentInfo where StudentID='" & Trim(txtfields(7)) & "'")
'            If Not rs1.EOF Then
'                Label12(7).Caption = "" & rs1!StudentName
'            End If
'            rs1.Close
       

End Select
End Sub
Private Sub ShowFlexData()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT classID,ClassName,Shiftname,StartTime,EndTime From ClassInfo")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = "" & rs!classId
                .TextMatrix(i, 1) = "" & rs!ClassName
                .TextMatrix(i, 2) = "" & rs!Shiftname
                .TextMatrix(i, 3) = "" & rs!Starttime
                .TextMatrix(i, 4) = "" & rs!Endtime
                
                i = i + 1
            rs.MoveNext
        Loop
    End With

 End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub ShowFlexData1()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT SectionID,SectionDsc,SectionRoomNo,classteacher,SecMonitor1,SecMonitor2,Effectivedate,trackid from SectionHistoryInfo where ClassID='" & Trim(txtfields(0)) & "' order by trackId desc")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid2
        Do Until rs.EOF
            MSFlexGrid2.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = "" & rs!SectionID
                .TextMatrix(i, 1) = "" & rs!Sectiondsc
                .TextMatrix(i, 2) = "" & rs!SectionRoomNo
                .TextMatrix(i, 3) = "" & rs!classteacher
                .TextMatrix(i, 4) = "" & get_name(1, rs!classteacher)
                .TextMatrix(i, 5) = "" & rs!SecMonitor1
                .TextMatrix(i, 6) = "" & get_name(2, rs!SecMonitor1)
                .TextMatrix(i, 7) = "" & rs!SecMonitor2
                .TextMatrix(i, 8) = "" & get_name(2, rs!SecMonitor2)
                .TextMatrix(i, 9) = "" & rs!EffectiveDate
                .TextMatrix(i, 10) = "" & rs!trackid
                i = i + 1
            rs.MoveNext
        Loop
    End With
 Else
     MSFlexGrid2.Rows = 1

 End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title
End Sub
