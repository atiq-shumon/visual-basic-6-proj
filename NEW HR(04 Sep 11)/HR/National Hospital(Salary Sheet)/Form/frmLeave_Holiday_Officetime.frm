VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form10 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Leave Setup"
   ClientHeight    =   5955
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   8175
   ClipControls    =   0   'False
   Icon            =   "frmLeave_Holiday_Officetime.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110.247
   ScaleMode       =   0  'User
   ScaleWidth      =   7676.75
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   3465
      Picture         =   "frmLeave_Holiday_Officetime.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5310
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   4755
      Picture         =   "frmLeave_Holiday_Officetime.frx":22D4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5310
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   810
      Picture         =   "frmLeave_Holiday_Officetime.frx":3EBE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5310
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   2130
      Picture         =   "frmLeave_Holiday_Officetime.frx":5850
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5310
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   6075
      Picture         =   "frmLeave_Holiday_Officetime.frx":71E2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5310
      Width           =   1185
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5070
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   8943
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Leave Setup"
      TabPicture(0)   =   "frmLeave_Holiday_Officetime.frx":8C64
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Holiday Setup"
      TabPicture(1)   =   "frmLeave_Holiday_Officetime.frx":8C80
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Office Time"
      TabPicture(2)   =   "frmLeave_Holiday_Officetime.frx":8C9C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(2)"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   4740
         Index           =   2
         Left            =   -75000
         TabIndex        =   39
         Top             =   315
         Width           =   8025
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   1455
            Left            =   360
            TabIndex        =   59
            Top             =   2925
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   2566
            _Version        =   393216
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
                  LCID            =   2057
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
                  LCID            =   2057
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
         Begin VB.ComboBox cmbSp_End_Day 
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
            Height          =   330
            ItemData        =   "frmLeave_Holiday_Officetime.frx":8CB8
            Left            =   4230
            List            =   "frmLeave_Holiday_Officetime.frx":8CEC
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   2070
            Width           =   1320
         End
         Begin VB.ComboBox cmbSp_Start_Time 
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
            Height          =   330
            ItemData        =   "frmLeave_Holiday_Officetime.frx":8D2F
            Left            =   2340
            List            =   "frmLeave_Holiday_Officetime.frx":8D84
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   2070
            Width           =   1320
         End
         Begin VB.ComboBox cmbSp_Start_Day 
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
            Height          =   330
            ItemData        =   "frmLeave_Holiday_Officetime.frx":8E90
            Left            =   450
            List            =   "frmLeave_Holiday_Officetime.frx":8EAC
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   2070
            Width           =   1320
         End
         Begin VB.ComboBox cmbSp_End_time 
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
            Height          =   330
            ItemData        =   "frmLeave_Holiday_Officetime.frx":8EFC
            Left            =   6030
            List            =   "frmLeave_Holiday_Officetime.frx":8F51
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   2070
            Width           =   1320
         End
         Begin VB.ComboBox cmbAbs_Time 
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
            Height          =   330
            ItemData        =   "frmLeave_Holiday_Officetime.frx":9060
            Left            =   4230
            List            =   "frmLeave_Holiday_Officetime.frx":90B5
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   855
            Width           =   1320
         End
         Begin VB.ComboBox cmbEnd_time 
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
            Height          =   330
            ItemData        =   "frmLeave_Holiday_Officetime.frx":91C1
            Left            =   2340
            List            =   "frmLeave_Holiday_Officetime.frx":9216
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   855
            Width           =   1320
         End
         Begin VB.ComboBox cmbStart_time 
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
            Height          =   330
            ItemData        =   "frmLeave_Holiday_Officetime.frx":9322
            Left            =   450
            List            =   "frmLeave_Holiday_Officetime.frx":9377
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   855
            Width           =   1320
         End
         Begin MSComCtl2.DTPicker dtpEffect_date 
            Height          =   330
            Left            =   6030
            TabIndex        =   43
            Top             =   855
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyyy"
            Format          =   22675459
            CurrentDate     =   37316
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Special Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   4
            Left            =   405
            TabIndex        =   57
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Usual Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   3
            Left            =   450
            TabIndex        =   56
            Top             =   180
            Width           =   945
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   1545
            Index           =   14
            Left            =   315
            Top             =   2880
            Width           =   7305
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Day"
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
            Left            =   450
            TabIndex        =   55
            Top             =   1800
            Width           =   285
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Start time"
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
            Index           =   1
            Left            =   2340
            TabIndex        =   54
            Top             =   1800
            Width           =   675
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Relax Time"
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
            Index           =   1
            Left            =   4230
            TabIndex        =   53
            Top             =   1800
            Width           =   780
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "End time"
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
            Left            =   6075
            TabIndex        =   52
            Top             =   1800
            Width           =   600
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   420
            Index           =   9
            Left            =   5985
            Top             =   2025
            Width           =   1410
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   420
            Index           =   8
            Left            =   4185
            Top             =   2025
            Width           =   1410
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   420
            Index           =   7
            Left            =   405
            Top             =   2025
            Width           =   1410
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   420
            Index           =   2
            Left            =   2295
            Top             =   2025
            Width           =   1410
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Absent After"
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
            Index           =   1
            Left            =   4275
            TabIndex        =   47
            Top             =   585
            Width           =   945
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Effect date"
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
            Index           =   0
            Left            =   6030
            TabIndex        =   46
            Top             =   585
            Width           =   795
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Start time"
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
            Index           =   2
            Left            =   450
            TabIndex        =   45
            Top             =   585
            Width           =   675
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "End time"
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
            Index           =   0
            Left            =   2340
            TabIndex        =   44
            Top             =   585
            Width           =   600
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   420
            Index           =   13
            Left            =   5985
            Top             =   810
            Width           =   1410
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   420
            Index           =   12
            Left            =   4185
            Top             =   810
            Width           =   1410
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   420
            Index           =   11
            Left            =   2295
            Top             =   810
            Width           =   1410
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   420
            Index           =   10
            Left            =   405
            Top             =   810
            Width           =   1410
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   4740
         Left            =   -75000
         TabIndex        =   22
         Top             =   315
         Width           =   8025
         Begin MSDataGridLib.DataGrid DataGrid4 
            Height          =   2445
            Left            =   225
            TabIndex        =   58
            Top             =   1530
            Width           =   7485
            _ExtentX        =   13203
            _ExtentY        =   4313
            _Version        =   393216
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
                  LCID            =   2057
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
                  LCID            =   2057
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
         Begin VB.ComboBox cmbCategory 
            Appearance      =   0  'Flat
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
            Height          =   330
            ItemData        =   "frmLeave_Holiday_Officetime.frx":9484
            Left            =   1170
            List            =   "frmLeave_Holiday_Officetime.frx":9497
            TabIndex        =   31
            Top             =   945
            Width           =   1905
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
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
            Height          =   330
            Index           =   4
            Left            =   4410
            TabIndex        =   30
            Top             =   360
            Width           =   3165
         End
         Begin VB.ComboBox cmbYear 
            Appearance      =   0  'Flat
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
            Height          =   330
            ItemData        =   "frmLeave_Holiday_Officetime.frx":94D8
            Left            =   1170
            List            =   "frmLeave_Holiday_Officetime.frx":9506
            TabIndex        =   29
            Top             =   360
            Width           =   960
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Weekend"
            ForeColor       =   &H00800000&
            Height          =   510
            Left            =   315
            TabIndex        =   24
            Top             =   4050
            Width           =   6945
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Friday"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   6
               Left            =   6120
               TabIndex        =   0
               Top             =   225
               Value           =   -1  'True
               Width           =   780
            End
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Thursday"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   5
               Left            =   5115
               TabIndex        =   28
               Top             =   225
               Width           =   1050
            End
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Wednesday"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   4
               Left            =   3930
               TabIndex        =   27
               Top             =   225
               Width           =   1275
            End
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Tuesday"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   2970
               TabIndex        =   26
               Top             =   225
               Width           =   1050
            End
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Monday"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   2010
               TabIndex        =   25
               Top             =   225
               Width           =   1005
            End
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Sunday"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   1095
               TabIndex        =   1
               Top             =   225
               Width           =   1005
            End
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Saturday"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   2
               Top             =   225
               Width           =   1005
            End
         End
         Begin VB.CommandButton CmdHoliday 
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   7290
            Picture         =   "frmLeave_Holiday_Officetime.frx":955E
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   4140
            Width           =   420
         End
         Begin MSComCtl2.DTPicker dtpEnd_date 
            Height          =   330
            Left            =   6480
            TabIndex        =   32
            Top             =   945
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
            _Version        =   393216
            CalendarForeColor=   8388608
            CalendarTitleForeColor=   8388608
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   22675459
            CurrentDate     =   37004
         End
         Begin MSComCtl2.DTPicker dtpStr_date 
            Height          =   330
            Left            =   4365
            TabIndex        =   33
            Top             =   945
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
            _Version        =   393216
            CalendarForeColor=   8388608
            CalendarTitleForeColor=   8388608
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   22675459
            CurrentDate     =   36995
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To (Date)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   5670
            TabIndex        =   38
            Top             =   1035
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   37
            Top             =   990
            Width           =   630
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From (Date)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   3195
            TabIndex        =   36
            Top             =   1035
            Width           =   825
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Holiday  Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   3195
            TabIndex        =   35
            Top             =   405
            Width           =   1035
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   420
            Index           =   0
            Left            =   4320
            Top             =   315
            Width           =   3345
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   2535
            Index           =   1
            Left            =   180
            Top             =   1485
            Width           =   7575
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   420
            Index           =   3
            Left            =   4320
            Top             =   900
            Width           =   1320
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   420
            Index           =   4
            Left            =   6435
            Top             =   900
            Width           =   1320
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   420
            Index           =   5
            Left            =   1125
            Top             =   900
            Width           =   1995
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Year"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   34
            Top             =   360
            Width           =   330
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   420
            Index           =   6
            Left            =   1125
            Top             =   315
            Width           =   1050
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000E&
         Height          =   4740
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   315
         Width           =   8025
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   0
            Left            =   1395
            TabIndex        =   15
            Top             =   360
            Width           =   1050
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   2
            Left            =   1395
            TabIndex        =   14
            Top             =   765
            Width           =   1050
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   3
            Left            =   6705
            TabIndex        =   13
            Top             =   765
            Width           =   1005
         End
         Begin VB.OptionButton Opt 
            BackColor       =   &H8000000E&
            Caption         =   "Yes"
            ForeColor       =   &H8000000D&
            Height          =   240
            Index           =   0
            Left            =   4050
            TabIndex        =   12
            Top             =   855
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton Opt 
            BackColor       =   &H8000000E&
            Caption         =   "No"
            ForeColor       =   &H8000000D&
            Height          =   330
            Index           =   1
            Left            =   4905
            TabIndex        =   11
            Top             =   810
            Width           =   510
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   1
            Left            =   4050
            TabIndex        =   10
            Top             =   360
            Width           =   3660
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3075
            Left            =   270
            TabIndex        =   16
            Top             =   1350
            Width           =   7395
            _ExtentX        =   13044
            _ExtentY        =   5424
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
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
                  ColumnWidth     =   1425.26
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1425.26
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000E&
            Caption         =   "Leave Code"
            ForeColor       =   &H8000000D&
            Height          =   330
            Index           =   1
            Left            =   315
            TabIndex        =   21
            Top             =   405
            Width           =   1185
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000E&
            Caption         =   "Leave Name"
            ForeColor       =   &H8000000D&
            Height          =   285
            Index           =   1
            Left            =   2610
            TabIndex        =   20
            Top             =   450
            Width           =   1230
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000E&
            Caption         =   " Leave Days"
            ForeColor       =   &H8000000D&
            Height          =   240
            Index           =   1
            Left            =   270
            TabIndex        =   19
            Top             =   855
            Width           =   915
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000E&
            Caption         =   "Carry Days"
            ForeColor       =   &H8000000D&
            Height          =   285
            Index           =   1
            Left            =   5670
            TabIndex        =   18
            Top             =   855
            Width           =   1365
         End
         Begin VB.Label Label5 
            BackColor       =   &H8000000E&
            Caption         =   "Carry Arrear"
            ForeColor       =   &H8000000D&
            Height          =   240
            Index           =   1
            Left            =   2610
            TabIndex        =   17
            Top             =   855
            Width           =   1365
         End
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Leave_Info As New LeaveSetUp
Private HolidaySetUp As New clsSt_Holiday
Private OfficeTimingSetUp As New clsSt_Office_Time
Dim SSTab_Index As Integer

Private Sub cmbAbs_Time_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   ' dtpEffect_date
End If
End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    dtpStr_date.SetFocus
End If
End Sub

Private Sub cmbEnd_time_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmbAbs_Time.SetFocus
End If
End Sub

Private Sub cmbSp_End_Day_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmbSp_End_time.SetFocus
End If
End Sub

Private Sub cmbSp_End_time_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdSave.SetFocus
End If
End Sub

Private Sub cmbSp_Start_Day_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmbSp_Start_Time.SetFocus
End If
End Sub

Private Sub cmbSp_Start_Time_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmbSp_End_Day.SetFocus
End If
End Sub

Private Sub cmbStart_time_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmbEnd_time.SetFocus
End If
End Sub

Private Sub cmbYear_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtfields(4).SetFocus
End If
End Sub

Private Sub cmdClear_Click()
On Error GoTo Errdes

If SSTab1.Tab = 0 Then

   For i = 0 To 3
    txtfields(i).Text = ""
   Next
   
   Opt(0).Value = True
   txtfields(0).SetFocus
   
ElseIf SSTab1.Tab = 1 Then
    txtfields(4).Text = ""
    cmbyear.Text = cmbyear.List(0)
    cmbCategory.Text = cmbCategory.List(0)
    dtpStr_date = Date
    dtpEnd_date = Date
    cmbyear.SetFocus

Else
   cmbStart_time.SetFocus

End If

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If SSTab1.Tab = 0 Then
    Call LeaveSetup_Delete
ElseIf SSTab1.Tab = 1 Then
    Call HolidaySetup_Delete
Else
    Call OfficeTimeSetup_Delete
End If
End Sub
Private Sub LeaveSetup_Delete()
'On Error GoTo Errdes
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub HolidaySetup_Delete()
'On Error GoTo Errdes
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub OfficeTimeSetup_Delete()
'On Error GoTo Errdes
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub cmdSave_Click()
If SSTab1.Tab = 0 Then
    Call LeaveSetup_Save
ElseIf SSTab1.Tab = 1 Then
    Call HolidaySetup_Save
Else
    Call OfficeTimeSetup_Save
End If

End Sub
Private Sub LeaveSetup_Save()
On Error GoTo Errdes
Select Case SSTab_Index
Dim conn As New Connection
Case 0

With Leave_Info
    
    .Connstring = strCN.Connection_String
    .LEAVE_CODE = txtfields(0)
    .Leave_Name = txtfields(1)
    
    If Opt(0).Value = True Then
        .Carry_Arrear_Days = 0
    Else
        .Carry_Arrear_Days = 1
    End If
    
    .Days = txtfields(2)
    .Carry_Max_Days = txtfields(3)
    .Save
       
End With
    get_LeaveInfo_Into_Grid

End Select

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub HolidaySetup_Save()
On Error GoTo Errdes
    
    With HolidaySetUp
    
        .Connstring = strCN.Connection_String
        .Year_To_Show = Trim(cmbyear)
        
        .Holiday_Name = txtfields(4)
        
        If Trim(cmbCategory) = "Government Holiday" Then
            .H_Type = 0
        ElseIf Trim(cmbCategory) = "Public Holiday" Then
             .H_Type = 1
        ElseIf Trim(cmbCategory) = "Weekend" Then
             .H_Type = 2
        ElseIf Trim(cmbCategory) = "Hartal" Then
             .H_Type = 3
        ElseIf Trim(cmbCategory) = "Others" Then
             .H_Type = 4
        End If
        
        .From_Dt = dtpStr_date
        .To_Dt = dtpEnd_date
        
        .Save
    End With
    get_LeaveInfo_Into_Grid
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub OfficeTimeSetup_Save()
On Error GoTo Errdes
 With OfficeTimingSetUp
 
        .Connstring = strCN.Connection_String
        .Start_Time = Trim(cmbStart_time)
        .End_Time = Trim(cmbEnd_time)
        .Relax_Time = Trim(cmbSp_End_Day)
        .Absent_Time = Trim(cmbAbs_Time)
        .Special_Start_Time = Trim(cmbSp_Start_Time)
        .Special_End_Time = Trim(cmbSp_End_time)
        .Special_Day = Trim(cmbSp_Start_Day)
        .Effect_Dt = Format(dtpEffect_date, "dd/mm/yyyy")
        .Save
    
 End With
    get_LeaveInfo_Into_Grid
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub dtpEffect_date_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmbSp_Start_Day.SetFocus
End If
End Sub

Private Sub dtpEnd_date_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  '  Option1(7).SetFocus
End If
End Sub

Private Sub dtpStr_date_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    dtpEnd_date.SetFocus
End If
End Sub

Private Sub Form_Load()
If SSTab1.Tab = 0 Then
    get_LeaveInfo_Into_Grid
ElseIf SSTab1.Tab = 1 Then
    get_LeaveInfo_Into_Grid
ElseIf SSTab1.Tab = 2 Then
    get_LeaveInfo_Into_Grid
End If
End Sub

Private Sub Opt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Select Case Index

Case 0
    If Opt(0).Value = True Then
        txtfields(3).SetFocus
    Else
        Opt(1).SetFocus
    End If
    
Case 1
    If Opt(1).Value = True Then
        txtfields(3).SetFocus
    End If

End Select
End If

End Sub

Private Sub Option1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Select Case Index
Case 0

    If Option1(0).Value = True Then
        cmdSave.SetFocus
    End If
    
Case 1

     If Option1(1).Value = True Then
        cmdSave.SetFocus
    End If

Case 2

     If Option1(2).Value = True Then
        cmdSave.SetFocus
    End If

Case 3

    If Option1(3).Value = True Then
        cmdSave.SetFocus
    End If

Case 4

     If Option1(4).Value = True Then
        cmdSave.SetFocus
    End If
    
Case 5
     If Option1(5).Value = True Then
        cmdSave.SetFocus
    End If

End Select
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
    get_LeaveInfo_Into_Grid
ElseIf SSTab1.Tab = 1 Then
    get_LeaveInfo_Into_Grid
End If
End Sub

Private Sub txtfields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

    Select Case Index
    
    Case 0
        txtfields(1).SetFocus
    Case 1
        txtfields(2).SetFocus
    Case 2
        Opt(0).SetFocus
    Case 3
        cmdSave.SetFocus
    Case 4
        cmbCategory.SetFocus
        
    End Select
    
End If

End Sub
Private Sub get_LeaveInfo_Into_Grid()
On Error GoTo Errdes
Dim conn2 As New Connection
Dim RS2 As New Recordset
Dim cmd As New Command
conn2.ConnectionString = strCN.Connection_String
conn2.Open
cmd.ActiveConnection = conn2
cmd.CommandType = adCmdText

If SSTab1.Tab = 0 Then
    cmd.CommandText = "select * from st_leave order by LEAVE_CODE "
ElseIf SSTab1.Tab = 1 Then
     cmd.CommandText = "Select * from st_holiday order by YEAR_TO_SH "
ElseIf SSTab1.Tab = 2 Then
      cmd.CommandText = "Select * from st_office_time "
End If

cmd.Properties("iRowsetChange") = True
cmd.Properties("updatability") = 7
RS2.CursorLocation = adUseClient
RS2.Open cmd.CommandText, conn2, adOpenDynamic, adLockOptimistic

If SSTab1.Tab = 0 Then

    If Not RS2.BOF Or RS2.EOF Then
        Set DataGrid1.DataSource = RS2
    Else
        Set DataGrid1.DataSource = Nothing
    End If
    
ElseIf SSTab1.Tab = 1 Then

    If Not RS2.BOF Or RS2.EOF Then
        Set DataGrid4.DataSource = RS2
    Else
        Set DataGrid4.DataSource = Nothing
    End If
    
ElseIf SSTab1.Tab = 2 Then

    If Not RS2.BOF Or RS2.EOF Then
       Set DataGrid2.DataSource = RS2
    Else
       Set DataGrid2.DataSource = Nothing
    End If

 
End If

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "Daffdoil Software Ltd"
End Sub
