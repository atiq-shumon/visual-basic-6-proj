VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form9 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Organization Setup"
   ClientHeight    =   5580
   ClientLeft      =   2205
   ClientTop       =   2160
   ClientWidth     =   8400
   ForeColor       =   &H00800000&
   Icon            =   "frmJob_setup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form13"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8400
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   225
      Top             =   4815
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   6210
      Picture         =   "frmJob_setup.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4950
      Width           =   1140
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   2310
      Picture         =   "frmJob_setup.frx":234C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4950
      Width           =   1140
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   1080
      Picture         =   "frmJob_setup.frx":3CDE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4950
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   4845
      Picture         =   "frmJob_setup.frx":5670
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4950
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   3600
      Picture         =   "frmJob_setup.frx":725A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4950
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4700
      Left            =   135
      TabIndex        =   5
      Top             =   90
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Org. Profile"
      TabPicture(0)   =   "frmJob_setup.frx":8C64
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Job Type"
      TabPicture(1)   =   "frmJob_setup.frx":8C80
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Department"
      TabPicture(2)   =   "frmJob_setup.frx":8C9C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Designation"
      TabPicture(3)   =   "frmJob_setup.frx":8CB8
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Height          =   4350
         Left            =   -75000
         TabIndex        =   42
         Top             =   345
         Width           =   8115
         Begin VB.TextBox txtNotes 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1395
            TabIndex        =   49
            Top             =   3555
            Width           =   6180
         End
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   1395
            TabIndex        =   48
            Top             =   3060
            Width           =   6180
         End
         Begin VB.TextBox txtFax 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1395
            TabIndex        =   47
            Top             =   2655
            Width           =   4110
         End
         Begin VB.TextBox txtPhone 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1395
            TabIndex        =   46
            Top             =   2205
            Width           =   4110
         End
         Begin VB.TextBox txtAddress 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   735
            Left            =   1395
            MultiLine       =   -1  'True
            TabIndex        =   45
            Top             =   1305
            Width           =   4065
         End
         Begin VB.ComboBox cboOrg_Type 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "frmJob_setup.frx":8CD4
            Left            =   1395
            List            =   "frmJob_setup.frx":8CE4
            TabIndex        =   44
            Top             =   855
            Width           =   4110
         End
         Begin VB.TextBox txtname 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   420
            Left            =   1395
            TabIndex        =   43
            Top             =   315
            Width           =   6135
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   1770
            Left            =   5625
            Top             =   1170
            Width           =   1950
         End
         Begin VB.Image imgLogo 
            Height          =   1695
            Left            =   5670
            Picture         =   "frmJob_setup.frx":8D36
            Stretch         =   -1  'True
            ToolTipText     =   "   Click to load picture  "
            Top             =   1215
            Width           =   1815
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   405
            TabIndex        =   57
            Top             =   1305
            Width           =   570
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   405
            TabIndex        =   56
            Top             =   855
            Width           =   360
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   405
            TabIndex        =   55
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Notes"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   405
            TabIndex        =   54
            Top             =   3510
            Width           =   420
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   405
            TabIndex        =   53
            Top             =   3060
            Width           =   420
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   405
            TabIndex        =   52
            Top             =   2655
            Width           =   255
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   405
            TabIndex        =   51
            Top             =   2250
            Width           =   465
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Company Logo"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5760
            TabIndex        =   50
            Top             =   855
            Width           =   1590
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   4350
         Left            =   0
         TabIndex        =   22
         Top             =   345
         Width           =   8115
         Begin VB.TextBox txtLevel 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   7110
            TabIndex        =   35
            Top             =   450
            Width           =   735
         End
         Begin VB.TextBox txtDesignation 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   3555
            TabIndex        =   34
            Top             =   450
            Width           =   2715
         End
         Begin VB.TextBox txtDesigCode 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1575
            MaxLength       =   3
            TabIndex        =   33
            Top             =   450
            Width           =   780
         End
         Begin VB.OptionButton optEmp_Type 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Officer"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   1530
            TabIndex        =   32
            Top             =   945
            Value           =   -1  'True
            Width           =   870
         End
         Begin VB.OptionButton optEmp_Type 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Staff"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   2520
            TabIndex        =   31
            Top             =   945
            Width           =   870
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4410
            TabIndex        =   28
            Top             =   945
            Width           =   3390
            Begin VB.OptionButton optCommission 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Wage Commission"
               ForeColor       =   &H00800000&
               Height          =   240
               Index           =   1
               Left            =   1800
               TabIndex        =   30
               Top             =   0
               Width           =   1770
            End
            Begin VB.OptionButton optCommission 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Pay Commission"
               ForeColor       =   &H00800000&
               Height          =   240
               Index           =   0
               Left            =   270
               TabIndex        =   29
               Top             =   0
               Value           =   -1  'True
               Width           =   1545
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1530
            TabIndex        =   24
            Top             =   1305
            Width           =   3120
            Begin VB.OptionButton optPool 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Finance"
               ForeColor       =   &H00800000&
               Height          =   285
               Index           =   2
               Left            =   2070
               TabIndex        =   27
               Top             =   0
               Width           =   1050
            End
            Begin VB.OptionButton optPool 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Technical"
               ForeColor       =   &H00800000&
               Height          =   285
               Index           =   1
               Left            =   990
               TabIndex        =   26
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton optPool 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "General"
               ForeColor       =   &H00800000&
               Height          =   285
               Index           =   0
               Left            =   0
               TabIndex        =   25
               Top             =   0
               Value           =   -1  'True
               Width           =   1005
            End
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2400
            Index           =   0
            Left            =   315
            TabIndex        =   23
            Top             =   1710
            Width           =   7485
            _ExtentX        =   13203
            _ExtentY        =   4233
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            ColumnHeaders   =   0   'False
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
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
            ColumnCount     =   2
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Level"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   6435
            TabIndex        =   41
            Top             =   450
            Width           =   390
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Designation"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2565
            TabIndex        =   40
            Top             =   495
            Width           =   840
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Type"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   315
            TabIndex        =   39
            Top             =   945
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   315
            TabIndex        =   38
            Top             =   495
            Width           =   375
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pool / Cadre"
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   315
            TabIndex        =   37
            Top             =   1350
            Width           =   1050
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Salary Base"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   3555
            TabIndex        =   36
            Top             =   945
            Width           =   840
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   4350
         Left            =   -75000
         TabIndex        =   14
         Top             =   345
         Width           =   8115
         Begin VB.TextBox txtDesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   2
            Left            =   1215
            TabIndex        =   18
            Top             =   945
            Width           =   6585
         End
         Begin VB.TextBox txtDepartment 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   3600
            TabIndex        =   17
            Top             =   450
            Width           =   4200
         End
         Begin VB.TextBox txtDept_Code 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1215
            MaxLength       =   3
            TabIndex        =   16
            Top             =   450
            Width           =   1095
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2625
            Index           =   2
            Left            =   315
            TabIndex        =   15
            Top             =   1485
            Width           =   7485
            _ExtentX        =   13203
            _ExtentY        =   4630
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            ColumnHeaders   =   0   'False
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
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
            ColumnCount     =   2
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   315
            TabIndex        =   21
            Top             =   495
            Width           =   375
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   315
            TabIndex        =   20
            Top             =   945
            Width           =   795
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Department"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2565
            TabIndex        =   19
            Top             =   495
            Width           =   825
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   4350
         Left            =   -75000
         TabIndex        =   6
         Top             =   345
         Width           =   8115
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2625
            Index           =   1
            Left            =   315
            TabIndex        =   13
            Top             =   1485
            Width           =   7485
            _ExtentX        =   13203
            _ExtentY        =   4630
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            ColumnHeaders   =   0   'False
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   15
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
            ColumnCount     =   2
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtJTypeCode 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1215
            TabIndex        =   9
            Top             =   450
            Width           =   1095
         End
         Begin VB.TextBox txtJType 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   3600
            TabIndex        =   8
            Top             =   450
            Width           =   4200
         End
         Begin VB.TextBox txtDesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   1
            Left            =   1215
            TabIndex        =   7
            Top             =   945
            Width           =   6585
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Job Type"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2565
            TabIndex        =   12
            Top             =   495
            Width           =   660
         End
         Begin VB.Label lblDescription 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   315
            TabIndex        =   11
            Top             =   945
            Width           =   795
         End
         Begin VB.Label lblCode 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   315
            TabIndex        =   10
            Top             =   495
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Desig As New St_Desig
Private Desig_Rs As New Recordset
Private JType As New St_JbType
Private JType_Rs As New Recordset
Private Comp_Info As New Company_Info
Private Dept As New St_Department
Dim Default_Pic_Path As String
Dim New_Pic_Path As String
Dim Pool As Integer
Dim Emp_Type As Integer
Dim PW_Comm As Integer
Dim SSTab_Index As Integer
Private Sub cmdClear_Click()
Clear_Screen

If SSTab_Index = 0 Then
    txtDesigCode.SetFocus
Else
    txtJTypeCode.SetFocus
End If
    
End Sub
Private Sub cmdClose_Click()
    Close_Msg Me
End Sub
Private Sub cmdDelete_Click()
On Error GoTo Errdes
Select Case SSTab_Index

Case 3

    With Desig
        .Connstring = strCN.Connection_String
        .Desig_Code = txtDesigCode
        .Delete
    End With

Case 1
    With JType
        .Connstring = strCN.Connection_String
        .JType_Code = txtJTypeCode
        .Delete
    End With
    
Case 2

    With Dept
        .Connstring = strCN.Connection_String
        '.DEPT_CODE = txtDeptCode
        .DEPT_CODE = txtDept_Code
        .Delete
    End With

Case 0

End Select

Flash_Into_Grid
Clear_Screen

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub

Private Sub cmdSave_Click()
On Error GoTo Errdes
Select Case SSTab_Index

Case 3

    With Desig
        .Connstring = strCN.Connection_String
        .Desig_Code = txtDesigCode
        .designation = txtDesignation
        .Desig_Level = txtLevel
        .Pool = Pool
        .Emp_Type = Emp_Type
        .PW_Commission = PW_Comm
        .Save
        .Show_Message
    End With
    txtDesigCode.SetFocus

Case 1

    With JType
        .Connstring = strCN.Connection_String
        .JType_Code = Me.txtJTypeCode
        .JType_Nm = txtJType
        .Description = txtDesc(1)
        .Save
        .Show_Message
    End With
    txtJTypeCode.SetFocus

Case 2

     With Dept
        .Connstring = strCN.Connection_String
        .DEPT_CODE = txtDept_Code
        .DEPT_NM = txtDepartment
        .Description = txtDesc(2)
        .Save
        .Show_Message
    End With

Case 0
    
     With Comp_Info
        .Connstring = strCN.Connection_String
        .Co_Nm = txtname
        .Co_Type = cboOrg_Type
        .Address = txtAddress
        .Phone = txtPhone
        .Fax = txtFax
        .E_mail = txtEmail
        .Notes = txtNotes
        .Logo = New_Pic_Path
        .Save
        .Show_Message
    End With
    
End Select
    
    Flash_Into_Grid
    
    If SSTab_Index = 0 Then Exit Sub
    
    Clear_Screen

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
      
End Sub
Private Sub DataGrid1_Click(Index As Integer)
On Error GoTo Errdes
Select Case Index
Case 0

    txtDesigCode = DataGrid1(0).Columns(0)
    txtDesignation = DataGrid1(0).Columns(1)
    txtLevel = DataGrid1(0).Columns(3)
    
'    If DataGrid1(0).Columns(3) = 0 Then
'        optEmp_Type(0).Value = True
'    Else
'        optEmp_Type(1).Value = True
'    End If
'
'    If DataGrid1(0).Columns(4) = 0 Then
'        optCommission(0).Value = True
'    Else
'        optCommission(1).Value = True
'    End If
'
'    If DataGrid1(0).Columns(5) = 0 Then
'        optPool(0).Value = True
'    ElseIf DataGrid1(0).Columns(5) = 1 Then
'        optPool(1).Value = True
'    Else
'        optPool(2).Value = True
'    End If
Case 1

    txtJTypeCode = DataGrid1(1).Columns(0)
    txtJType = DataGrid1(1).Columns(1)
    txtDesc(1) = DataGrid1(1).Columns(2)

Case 2

    txtDept_Code = DataGrid1(2).Columns(0)
    txtDepartment = DataGrid1(2).Columns(1)
    txtDesc(2) = DataGrid1(2).Columns(2)


    



End Select

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Load()
On Error GoTo Errdes

SSTab_Index = 0
Screen_Position Me
Flash_Into_Grid

Set_Tab_Index

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
    
End Sub
Public Sub Flash_Into_Grid()
On Error GoTo Errdes

Dim RS As ADODB.Recordset

Select Case SSTab_Index

Case 3

    With Desig
        .Connstring = strCN.Connection_String
        Set RS = .GetAll
    
    End With
    
    Set DataGrid1(0).DataSource = RS
    
    With DataGrid1(0)
        .Columns(0).Width = 1500
        .Columns(1).Width = 8000
        '  .Columns(2).Width = 3400
    End With


Case 1

    With JType
        .Connstring = strCN.Connection_String
        Set RS = .GetAll
        End With
    
    Set DataGrid1(1).DataSource = RS
    
    With DataGrid1(1)
        .Columns(0).Width = 800
        .Columns(1).Width = 3000
        .Columns(2).Width = 6000
    End With
    

Case 2

    With Dept
        .Connstring = strCN.Connection_String
        Set RS = .GetAll
    End With
    
    Set DataGrid1(2).DataSource = RS
    
    With DataGrid1(2)
        .Columns(0).Width = 800
        .Columns(1).Width = 3000
        .Columns(2).Width = 6000
    End With

Case 0

    With Comp_Info
            .Connstring = strCN.Connection_String
            .Get_Company_Info
            
            txtname = .Co_Nm
            cboOrg_Type = .Co_Type
            txtAddress = .Address
            txtPhone = .Phone
            txtFax = .Fax
            txtEmail = .E_mail
            txtNotes = .Notes
        
        If Not .Logo = Empty Then
            imgLogo.Picture = LoadPicture(.Logo)
            Default_Pic_Path = .Logo
            New_Pic_Path = .Logo
        Else
            imgLogo.Picture = LoadPicture(Default_Pic_Path)
        End If
        
    End With

End Select
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
      
End Sub

Private Sub optCommission_Click(Index As Integer)
    PW_Comm = Index
End Sub

Private Sub optEmp_Type_Click(Index As Integer)

    Emp_Type = Index

End Sub

Private Sub optPool_Click(Index As Integer)

    Pool = Index
    
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
SSTab_Index = SSTab1.Tab
Set_Tab_Index
Flash_Into_Grid
End Sub
Public Sub Set_Tab_Index()
On Error GoTo Errdes
Select Case SSTab_Index

Case 0
    txtname.TabIndex = 0
    cboOrg_Type.TabIndex = 1
    txtAddress.TabIndex = 2
    txtPhone.TabIndex = 3
    txtFax.TabIndex = 4
    txtEmail.TabIndex = 5
    txtNotes.TabIndex = 6
    cmdSave.TabIndex = 7
    cmdClear.TabIndex = 8
    cmdClose.TabIndex = 9
    'txtName.SetFocus

Case 3
    txtDesigCode.TabIndex = 0
    txtDesignation.TabIndex = 1
    txtLevel.TabIndex = 2
    optEmp_Type(0).TabIndex = 3
    optEmp_Type(1).TabIndex = 4
    
    optCommission(0).TabIndex = 5
    optCommission(1).TabIndex = 6
    
    optPool(0).TabIndex = 7
    optPool(1).TabIndex = 8
    
    cmdSave.TabIndex = 9
    cmdClear.TabIndex = 10
    cmdClose.TabIndex = 11
    txtDesigCode.SetFocus

Case 1
    txtJTypeCode.TabIndex = 0
    txtJType.TabIndex = 1
    txtDesc(1).TabIndex = 2
    cmdSave.TabIndex = 3
    cmdClear.TabIndex = 4
    cmdClose.TabIndex = 5
    txtJTypeCode.SetFocus
    
Case 2

    txtDept_Code.TabIndex = 0
    txtDepartment.TabIndex = 1
    txtDesc(2).TabIndex = 2
    cmdSave.TabIndex = 3
    cmdClear.TabIndex = 4
    cmdClose.TabIndex = 5
    
    txtDept_Code.SetFocus

End Select

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Destroy Me
End Sub
Private Sub txtDept_Code_KeyPress(KeyAscii As Integer)
On Error GoTo Errdes

If KeyAscii = 13 And txtDept_Code <> "" Then

    With Dept
        .Connstring = strCN.Connection_String
        .DEPT_CODE = txtDept_Code
        .GetX
        
        txtDept_Code = .DEPT_CODE
        txtDepartment = .DEPT_NM
        txtDesc(2) = .Description
    End With

End If

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub txtDesigCode_KeyPress(KeyAscii As Integer)
On Error GoTo Errdes

If KeyAscii = 13 And txtDesigCode <> "" Then

    With Desig
        .Connstring = strCN.Connection_String
        .Desig_Code = txtDesigCode
        .GetX
        txtDesigCode = .Desig_Code
        txtDesignation = .designation
        txtLevel = .Desig_Level
        
        optCommission(.PW_Commission).Value = True
        optEmp_Type(.Emp_Type).Value = True
        optPool(.Pool).Value = True
    
    End With

End If
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub
Private Sub txtJTypeCode_KeyPress(KeyAscii As Integer)
On Error GoTo Errdes

If KeyAscii = 13 And txtJTypeCode <> "" Then

    With JType
        .Connstring = strCN.Connection_String
        .JType_Code = txtJTypeCode
        .GetX
        
        txtJTypeCode = .JType_Code
        txtJType = .JType_Nm
        txtDesc(1) = .Description
    
    End With

End If

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub txtLevel_KeyPress(KeyAscii As Integer)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
Public Sub Load_Photo(ComDiag As CommonDialog, Img As Image, Optional Photo_Path As String)
'On Error GoTo Errdes
'Dim resp As String
'
'
'Start:  With ComDiag
'            .Filter = "Photograph,*.bmp;*.jpg;*.gif|*.bmp;*.jpg;*.gif"
'            .Action = 1
'                If .FileName = "" Then
'                    Exit Sub
'                Else
'                    New_Pic_Path = .FileName
'                    Img.Picture = LoadPicture(New_Pic_Path)
'                End If
'        End With
'
''------------------------------------------------------
'
'        resp = MsgBox("           Is it the right Picture ?", vbYesNoCancel + vbQuestion + vbDefaultButton2, "Message")
'
'
'        If resp = vbCancel Then
'            Img.Picture = LoadPicture(Default_Pic_Path)
'            Exit Sub
'        End If
'
'        If resp = vbNo Then
'            Img.Picture = LoadPicture(Default_Pic_Path)
'            GoTo Start
'            Exit Sub
'        End If
'
'        If resp = vbYes Then
'                New_Pic_Path = ComDiag.FileName
'            Exit Sub
'
'        End If
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub imgLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Call Load_Photo(CommonDialog1, imgLogo, Default_Pic_Path)
End If
End Sub
Private Sub Clear_Screen()
On Error GoTo Errdes

If SSTab1.Tab = 0 Then
    txtname.Text = ""
    txtAddress = ""
    txtPhone = ""
    txtEmail = ""
    txtNotes = ""
    txtname.SetFocus
ElseIf SSTab1.Tab = 1 Then
    txtJTypeCode = ""
    txtJType = ""
    txtDesc(1) = ""
    txtJTypeCode.SetFocus
ElseIf SSTab1.Tab = 2 Then
    txtDept_Code = ""
    txtDepartment = ""
    txtDesc(2) = ""
    txtDept_Code.SetFocus
Else
    txtDesigCode = ""
    txtDesignation = ""
    txtLevel = ""
    txtDesigCode.SetFocus
End If

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

