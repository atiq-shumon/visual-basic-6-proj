VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form21 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  User & Privilege"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "Form21.frx":0000
   LinkTopic       =   "Form21"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6000
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   10583
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      BackColor       =   16777215
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Create User"
      TabPicture(0)   =   "Form21.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "User Privilege"
      TabPicture(1)   =   "Form21.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   5685
         Left            =   -75000
         TabIndex        =   2
         Top             =   315
         Width           =   7350
         Begin MSDataGridLib.DataGrid dtgGiven 
            Bindings        =   "Form21.frx":0902
            Height          =   3975
            Left            =   4005
            TabIndex        =   14
            Top             =   1350
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   0
            ColumnHeaders   =   0   'False
            ForeColor       =   192
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   0
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Given Previleges"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "scr_no"
               Caption         =   "Screen no."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "descript"
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               RecordSelectors =   0   'False
               BeginProperty Column00 
                  DividerStyle    =   0
                  ColumnWidth     =   0
               EndProperty
               BeginProperty Column01 
                  DividerStyle    =   0
                  ColumnWidth     =   4034.835
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dtgAll 
            Bindings        =   "Form21.frx":0917
            Height          =   3975
            Left            =   225
            TabIndex        =   13
            Top             =   1350
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BorderStyle     =   0
            ColumnHeaders   =   0   'False
            ForeColor       =   16711680
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   0
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Available Previleges"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "scr_no"
               Caption         =   "Screen"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "descript"
               Caption         =   "Privileges"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   5609.764
               EndProperty
            EndProperty
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   7470
            ScaleHeight     =   375
            ScaleWidth      =   510
            TabIndex        =   20
            Top             =   0
            Width           =   510
            Begin VB.CommandButton Command1 
               Appearance      =   0  'Flat
               Height          =   435
               Index           =   0
               Left            =   -45
               Picture         =   "Form21.frx":092C
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   3255
            Left            =   3285
            ScaleHeight     =   3255
            ScaleWidth      =   690
            TabIndex        =   15
            Top             =   1800
            Width           =   690
            Begin VB.CommandButton cmdSingle_In 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               Height          =   615
               Left            =   45
               Picture         =   "Form21.frx":12CE
               Style           =   1  'Graphical
               TabIndex        =   19
               ToolTipText     =   "  Give permission  "
               Top             =   135
               Width           =   600
            End
            Begin VB.CommandButton cmdAll_In 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               Height          =   615
               Left            =   45
               Picture         =   "Form21.frx":15D8
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "  Give all permission  "
               Top             =   945
               Width           =   600
            End
            Begin VB.CommandButton cmdSingle_Out 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               Height          =   615
               Left            =   45
               Picture         =   "Form21.frx":18E2
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "  Delete permission  "
               Top             =   1755
               Width           =   600
            End
            Begin VB.CommandButton cmdAll_Out 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               Height          =   615
               Left            =   45
               Picture         =   "Form21.frx":1BEC
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "  Delete all permission  "
               Top             =   2565
               Width           =   600
            End
         End
         Begin MSForms.ComboBox cboAccess_Area 
            Height          =   330
            Left            =   990
            TabIndex        =   25
            Top             =   810
            Width           =   2310
            VariousPropertyBits=   746604571
            ForeColor       =   12582912
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "4075;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   16761024
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboUser_ID 
            Height          =   330
            Left            =   990
            TabIndex        =   24
            Top             =   360
            Width           =   6135
            VariousPropertyBits=   746604571
            ForeColor       =   12582912
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "10821;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   16761024
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "User ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   7
            Left            =   270
            TabIndex        =   23
            Top             =   360
            Width           =   525
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   4065
            Index           =   4
            Left            =   180
            Top             =   1305
            Width           =   6945
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Access Area"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   435
            Index           =   5
            Left            =   270
            TabIndex        =   22
            Top             =   765
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   5685
         Left            =   0
         TabIndex        =   1
         Top             =   315
         Width           =   7350
         Begin VB.CommandButton cmdClose 
            Height          =   480
            Left            =   4995
            Picture         =   "Form21.frx":1EF6
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   4725
            Width           =   1185
         End
         Begin VB.CommandButton cmdClear 
            Height          =   480
            Left            =   2310
            Picture         =   "Form21.frx":3978
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   4725
            Width           =   1185
         End
         Begin VB.CommandButton cmdSave 
            Height          =   480
            Left            =   990
            Picture         =   "Form21.frx":530A
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   4725
            Width           =   1185
         End
         Begin VB.CommandButton cmdDelete 
            Height          =   480
            Left            =   3645
            Picture         =   "Form21.frx":6C9C
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   4725
            Width           =   1185
         End
         Begin VB.CheckBox chkAccess 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Restrict access temporarily"
            ForeColor       =   &H00800000&
            Height          =   420
            Left            =   5445
            TabIndex        =   26
            Top             =   315
            Value           =   1  'Checked
            Width           =   1680
         End
         Begin MSDataGridLib.DataGrid dtgResult 
            Height          =   2085
            Left            =   315
            TabIndex        =   7
            Top             =   2475
            Width           =   6675
            _ExtentX        =   11774
            _ExtentY        =   3678
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BorderStyle     =   0
            ColumnHeaders   =   -1  'True
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   6
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   3014.929
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1709.858
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   450.142
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.StatusBar StatusBar1 
            Height          =   330
            Left            =   45
            TabIndex        =   32
            Top             =   5310
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   1
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Object.Width           =   12788
                  MinWidth        =   12788
               EndProperty
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSForms.ComboBox cboUserType 
            Height          =   330
            Left            =   1755
            TabIndex        =   33
            Top             =   1080
            Width           =   3435
            VariousPropertyBits=   746604571
            ForeColor       =   12582912
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "6059;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   16761024
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   0
            Left            =   6345
            TabIndex        =   31
            Top             =   2115
            Visible         =   0   'False
            Width           =   645
         End
         Begin MSForms.TextBox txtConPass 
            Height          =   330
            Left            =   1755
            TabIndex        =   12
            Top             =   1890
            Width           =   3435
            VariousPropertyBits=   746604571
            ForeColor       =   12582912
            BorderStyle     =   1
            Size            =   "6059;582"
            PasswordChar    =   35
            BorderColor     =   16761024
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtPass 
            Height          =   330
            Left            =   1755
            TabIndex        =   11
            Top             =   1485
            Width           =   3435
            VariousPropertyBits=   746604571
            ForeColor       =   12582912
            BorderStyle     =   1
            Size            =   "6059;582"
            PasswordChar    =   35
            BorderColor     =   16761024
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "User Type"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   4
            Left            =   405
            TabIndex        =   10
            Top             =   1125
            Width           =   1320
         End
         Begin MSForms.TextBox txtEmpID 
            Height          =   330
            Left            =   1755
            TabIndex        =   9
            Top             =   270
            Width           =   1545
            VariousPropertyBits=   746604571
            ForeColor       =   12582912
            BorderStyle     =   1
            Size            =   "2725;582"
            BorderColor     =   16761024
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtName 
            Height          =   330
            Index           =   0
            Left            =   1755
            TabIndex        =   8
            Top             =   675
            Width           =   3435
            VariousPropertyBits=   746604571
            ForeColor       =   12582912
            BorderStyle     =   1
            Size            =   "6059;582"
            BorderColor     =   16761024
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   2160
            Index           =   3
            Left            =   270
            Top             =   2430
            Width           =   6750
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "User ID"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   405
            TabIndex        =   6
            Top             =   315
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "User Name"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   405
            TabIndex        =   5
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   405
            TabIndex        =   4
            Top             =   1575
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   3
            Left            =   405
            TabIndex        =   3
            Top             =   1935
            Width           =   1320
         End
      End
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim User_Pass As New Security
Dim User_Rs As New ADODB.Recordset
Dim Access_Rs As New ADODB.Recordset

Dim G_PrivRs As ADODB.Recordset
Dim N_PrivRs As ADODB.Recordset

Dim SSTab_Index As Integer
Dim Pw As New EnDecode.clsEndecoder
Dim Perm_Description As String


Private Sub cboAccess_Area_Click()
    Privileges
End Sub

Private Sub cboUser_ID_click()
    Privileges
End Sub

Private Sub cmdAll_In_Click()
    Give_Permission ("All")
End Sub

Private Sub cmdAll_Out_Click()
    Revoke_Permission ("All")
End Sub

Private Sub cmdClear_Click()
    Clear_Screen
    StatusBar1.Panels(1).Text = "No. of current user:  " & User_Rs.RecordCount
    txtEmpID.SetFocus
End Sub


Private Sub cmdClose_Click()
    Close_Msg Me
End Sub

Private Sub cmdDelete_Click()

    Dim A As String
    
    If MsgBox("Do you really want to delete user ID " + txtEmpID + " ?", vbYesNo, "Confirmation") = vbYes Then
        With User_Pass
            .ConnString = strCN.Connection
         A = .Delete_User(Trim(txtEmpID))
        End With
        Show_Data
    End If
    
End Sub

Private Sub cmdSave_Click()
    Dim Result As Boolean
    
    With User_Pass
        .ConnString = strCN.Connection
        .User_id = txtEmpID
        .User_Name = txtName(SSTab_Index)
        .User_Group = cboUserType
        .Password = txtPass
        .Confirm_Password = txtConPass
                
        If chkAccess Then
            .Access = Yes
        Else
            .Access = No
        End If
        
        Result = .Save
        
        Show_Data
        
        If Result = False Then  'if password mismatched then
            SetFocus_To txtConPass  'setfocus to the control with selection
        End If
        
    End With

End Sub

Private Sub cmdSingle_In_Click()
    On Error Resume Next
    Perm_Description = dtgAll.Columns(1)
    Give_Permission ("Single")
End Sub

Private Sub cmdSingle_Out_Click()
    On Error Resume Next
    Perm_Description = dtgGiven.Columns(1)
    
    Revoke_Permission ("Single")

End Sub

Private Sub dtgAll_dblClick()
On Error Resume Next
    Perm_Description = dtgAll.Columns(1)
    Give_Permission ("Single")
End Sub

Private Sub dtgGiven_dblClick()
On Error Resume Next
    Perm_Description = dtgGiven.Columns(1)
    Revoke_Permission ("Single")

End Sub

Private Sub dtgResult_Click()
'On Error Resume Next
Dim i As Integer

    With User_Rs
        txtEmpID = .Fields(0)
        txtName(SSTab_Index) = .Fields(1)
        '-------------------------------------------------------------
      
            For i = 0 To cboUserType.ListCount - 1
                If Trim(cboUserType.List(i)) = Trim(User_Rs.Fields(2)) Then
                    cboUserType.ListIndex = i
                    Exit For
                End If
            Next
       '-------------------------------------------------------------
        If .Fields(3) = "Yes" Then
            chkAccess.Value = 1
        Else
            chkAccess.Value = 0
        End If
        
            With Pw
                  .InputString = User_Rs.Fields(4)
                  .Decode
                  txtPass = .OutputString
                  txtConPass = txtPass
            End With
            
    StatusBar1.Panels(1).Text = "Selected user:  " & .Fields(1) & ",  User ID=" & .Fields(0)
            
    End With
    

    
    

End Sub

Private Sub Form_Load()

    txtEmpID.MaxLength = Id_Len

    Screen_Position Me
    SSTab_Index = 0
    Set_TabIndex
        
    With cboUserType
        .AddItem "Standard User"
        .AddItem "Restricted User"
        .AddItem "General User"
        .AddItem "Insert Operator"
        .AddItem "Backup Operator"
        .AddItem "Administrator"
    End With
    
    
    Show_Data
            
End Sub

Private Sub lblName_Change(index As Integer)
    txtName(index) = Trim(lblName(index))
    If Len(Trim(txtName(index))) > 0 Then
        SetFocus_To cboUserType
    Else
        SetFocus_To txtName
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Dim i As Integer
   
    
     SSTab_Index = SSTab1.Tab
         
     Select Case SSTab_Index
        Case 0
            Show_Data
        Case 1
        'populate User '--------------------------
            With User_Pass
                .ConnString = strCN.Connection
                    Set User_Rs = .GetAll_User
                        With User_Rs
                            If .RecordCount > 0 Then
                                .MoveFirst
                                cboUser_ID.Clear
                                Do Until User_Rs.EOF = True
                                      cboUser_ID.AddItem .Fields(0) & " . " & .Fields(1)
                                      .MoveNext
                                Loop
                             End If
                        End With
             End With
        
                'Populate Access Area
                With User_Pass
                
                    Set Access_Rs = .GetAccess_Area
                        With Access_Rs
                            .MoveFirst
                            
                            Do Until Access_Rs.EOF = True
                                cboAccess_Area.AddItem .Fields(0)
                                .MoveNext
                            Loop
                        End With
                End With
     End Select
     
End Sub


Private Sub txtEmpID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Get_Employee txtEmpID, Me, True
End Sub

Public Sub Set_TabIndex()
On Error Resume Next
    Select Case SSTab_Index
    
        Case 0
            
            txtEmpID.TabIndex = 0
            txtName(SSTab_Index).TabIndex = 1
            cboUserType.TabIndex = 2
            txtPass.TabIndex = 3
            txtConPass.TabIndex = 4
            cmdSave.TabIndex = 5
            cmdClear.TabIndex = 6
            cmdClose.TabIndex = 7
            txtEmpID(SSTab_Index).SetFocus
       
    End Select

End Sub


Public Sub Show_Data()
On Error Resume Next
    Dim Panel_Des As String
    Dim i As Integer
    
    With User_Pass
       .ConnString = strCN.Connection
       Set User_Rs = .GetAll_User
    End With
    
    If Not (User_Rs.EOF Or User_Rs.BOF) Then
        Set dtgResult.DataSource = User_Rs
        
        'grid columns resizing as per field length
    
        
              With dtgResult
                    .Columns(0).Width = 1170
                    .Columns(1).Width = 3014
                    .Columns(2).Width = 1620
                    .Columns(3).Width = 550
              End With
        
        
        With StatusBar1
                .Panels(1).Text = "No. of current user:  " & User_Rs.RecordCount
                '.Panels(2).Text = "Search pattern  '" + Search_Pattern + "'"
                '.Panels(3).Text = User_RsS.RecordCount & " record(s) found"
        End With
    
    Else
        
              
        For i = 0 To 4
              dtgResult.Columns(i).Width = 1605
        Next
        
        Set dtgResult.DataSource = Nothing
        
'        With StatusBar1
'                .Panels(1).Text = "Search option: " + lblSearchOption
'                .Panels(2).Text = "Search pattern  '" + Search_Pattern + "'"
'                .Panels(3).Text = "Search is complete. There is no result to display"
'                 cmdExport.Enabled = False
'        End With
    End If
End Sub


Public Sub Privileges()

    With User_Pass
           .ConnString = strCN.Connection
           Set G_PrivRs = .Get_Previleges(Given, cboAccess_Area, ChunkStr(cboUser_ID, ".", False))
           Set N_PrivRs = .Get_Previleges(NotGiven, cboAccess_Area, ChunkStr(cboUser_ID, ".", False))
    End With
        
        Set dtgAll.DataSource = N_PrivRs
        Set dtgGiven.DataSource = G_PrivRs

End Sub
Public Sub Give_Permission(pType As String)
Dim A As String
    
    With User_Pass
           .ConnString = strCN.Connection
           A = .Give_Permission(pType, cboAccess_Area, Perm_Description, ChunkStr(cboUser_ID, ".", False))
    End With
        
    Privileges
        
End Sub

Public Sub Revoke_Permission(pType As String)
Dim A As String
    
    With User_Pass
           .ConnString = strCN.Connection
           A = .Revoke_Permission(pType, cboAccess_Area, Perm_Description, ChunkStr(cboUser_ID, ".", False))
    End With
        
    Privileges
        
End Sub
