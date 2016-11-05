VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form11 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Loan (PF & Others Loan)"
   ClientHeight    =   6900
   ClientLeft      =   2085
   ClientTop       =   1755
   ClientWidth     =   8025
   Icon            =   "frmPFLoan_App.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form41"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   8025
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   5580
      Picture         =   "frmPFLoan_App.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6210
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   1605
      Picture         =   "frmPFLoan_App.frx":234C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6210
      Width           =   1140
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   315
      Picture         =   "frmPFLoan_App.frx":3CDE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6210
      Width           =   1140
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   4200
      Picture         =   "frmPFLoan_App.frx":5670
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6210
      Width           =   1230
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   2895
      Picture         =   "frmPFLoan_App.frx":725A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6210
      Width           =   1140
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5820
      Left            =   135
      TabIndex        =   8
      Top             =   180
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   10266
      _Version        =   393216
      Tab             =   2
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
      TabCaption(0)   =   "Salary Advance"
      TabPicture(0)   =   "frmPFLoan_App.frx":8C64
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Loan"
      TabPicture(1)   =   "frmPFLoan_App.frx":8C80
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Loan Refund"
      TabPicture(2)   =   "frmPFLoan_App.frx":8C9C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame2(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   5505
         Index           =   0
         Left            =   -75000
         TabIndex        =   48
         Top             =   360
         Width           =   7755
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            Left            =   1425
            TabIndex        =   4
            Top             =   375
            Width           =   1410
         End
         Begin VB.CommandButton cmdView 
            Height          =   315
            Index           =   0
            Left            =   2885
            Picture         =   "frmPFLoan_App.frx":8CB8
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   375
            Width           =   330
         End
         Begin VB.TextBox txtNotes 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   465
            Index           =   0
            Left            =   1485
            MaxLength       =   15
            TabIndex        =   49
            Top             =   2235
            Width           =   5910
         End
         Begin MSComCtl2.DTPicker dtpIssue_Dt 
            Height          =   330
            Index           =   0
            Left            =   4500
            TabIndex        =   50
            Top             =   1260
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   582
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   12582912
            CalendarTrailingForeColor=   16576
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   64028673
            CurrentDate     =   37722
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2310
            Index           =   0
            Left            =   360
            TabIndex        =   19
            Top             =   2880
            Width           =   7035
            _ExtentX        =   12409
            _ExtentY        =   4075
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BorderStyle     =   0
            ForeColor       =   9895936
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
         Begin VB.Shape Shape3 
            BorderColor     =   &H00FFC0C0&
            Height          =   345
            Left            =   2835
            Top             =   360
            Width           =   420
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FFC0C0&
            Height          =   345
            Left            =   1395
            Top             =   360
            Width           =   1455
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   2400
            Index           =   6
            Left            =   315
            Top             =   2835
            Width           =   7125
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Designation"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   315
            TabIndex        =   20
            Top             =   825
            Width           =   840
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Issue Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   3555
            TabIndex        =   21
            Top             =   1260
            Width           =   765
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Department"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   3555
            TabIndex        =   22
            Top             =   810
            Width           =   825
         End
         Begin VB.Label Label31 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee ID"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   315
            TabIndex        =   26
            Top             =   375
            Width           =   900
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nos. Inst."
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   315
            TabIndex        =   63
            Top             =   1320
            Width           =   675
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inst. Paid"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   315
            TabIndex        =   62
            Top             =   1755
            Width           =   660
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Notes"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   315
            TabIndex        =   61
            Top             =   2175
            Width           =   420
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   555
            Index           =   1
            Left            =   1395
            Top             =   2160
            Width           =   6045
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   3555
            TabIndex        =   54
            Top             =   375
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   3555
            TabIndex        =   53
            Top             =   1755
            Width           =   540
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   5715
            TabIndex        =   52
            Top             =   1770
            Width           =   585
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   5505
         Index           =   1
         Left            =   -75000
         TabIndex        =   27
         Top             =   315
         Width           =   7755
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   1425
            TabIndex        =   31
            Top             =   315
            Width           =   1410
         End
         Begin VB.CommandButton cmdView 
            Height          =   285
            Index           =   1
            Left            =   2835
            Picture         =   "frmPFLoan_App.frx":9582
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   315
            Width           =   330
         End
         Begin VB.TextBox txtNotes 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   465
            Index           =   1
            Left            =   1485
            MaxLength       =   15
            TabIndex        =   28
            Top             =   2235
            Width           =   5910
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2310
            Index           =   1
            Left            =   360
            TabIndex        =   29
            Top             =   2880
            Width           =   7035
            _ExtentX        =   12409
            _ExtentY        =   4075
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BorderStyle     =   0
            ForeColor       =   9895936
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
         Begin MSComCtl2.DTPicker dtpIssue_Dt 
            Height          =   330
            Index           =   1
            Left            =   4440
            TabIndex        =   30
            Top             =   1260
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   582
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   12582912
            CalendarTrailingForeColor=   16576
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   64028673
            CurrentDate     =   37722
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Int."
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   6090
            TabIndex        =   40
            Top             =   1380
            Width           =   225
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFC0C0&
            Height          =   375
            Left            =   1395
            Top             =   270
            Width           =   1815
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount "
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   47
            Top             =   1733
            Width           =   585
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Slab Amount"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   46
            Top             =   1755
            Width           =   900
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   3555
            TabIndex        =   45
            Top             =   375
            Width           =   420
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   555
            Index           =   0
            Left            =   1395
            Top             =   2160
            Width           =   6045
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Notes"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   315
            TabIndex        =   39
            Top             =   2175
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Loan ID"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   315
            TabIndex        =   38
            Top             =   1305
            Width           =   570
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nos. Inst."
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   3540
            TabIndex        =   37
            Top             =   1740
            Width           =   675
         End
         Begin VB.Label Label31 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee ID"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   315
            TabIndex        =   36
            Top             =   375
            Width           =   900
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Department"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   3555
            TabIndex        =   35
            Top             =   810
            Width           =   825
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Issue Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   3495
            TabIndex        =   34
            Top             =   1260
            Width           =   765
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Designation"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   315
            TabIndex        =   33
            Top             =   825
            Width           =   840
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   2400
            Index           =   2
            Left            =   315
            Top             =   2835
            Width           =   7125
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   5505
         Index           =   2
         Left            =   0
         TabIndex        =   9
         Top             =   315
         Width           =   7755
         Begin VB.TextBox txtAmount 
            Height          =   315
            Index           =   3
            Left            =   3060
            TabIndex        =   66
            Top             =   1800
            Width           =   1035
         End
         Begin VB.ComboBox Combo 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   900
            Width           =   1995
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   1420
            TabIndex        =   41
            Top             =   360
            Width           =   1410
         End
         Begin VB.TextBox txtNotes 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   465
            Index           =   2
            Left            =   1485
            MaxLength       =   15
            TabIndex        =   5
            Top             =   2235
            Width           =   5910
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2310
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   2880
            Width           =   7035
            _ExtentX        =   12409
            _ExtentY        =   4075
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BorderStyle     =   0
            ForeColor       =   9895936
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
         Begin MSComCtl2.DTPicker dtpIssue_Dt 
            Height          =   330
            Index           =   2
            Left            =   4500
            TabIndex        =   3
            Top             =   1260
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   582
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   12582912
            CalendarTrailingForeColor=   16576
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   64028673
            CurrentDate     =   37722
         End
         Begin VB.Label lblName 
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            Height          =   255
            Index           =   0
            Left            =   4170
            TabIndex        =   64
            Top             =   420
            Width           =   3015
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H00FFC0C0&
            Height          =   375
            Left            =   2835
            Top             =   315
            Width           =   15
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H00FFC0C0&
            Height          =   375
            Left            =   1395
            Top             =   315
            Width           =   1455
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inst.Due"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   6120
            TabIndex        =   42
            Top             =   1800
            Width           =   600
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   4275
            TabIndex        =   25
            Top             =   1770
            Width           =   585
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Slub Amount "
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   2025
            TabIndex        =   24
            Top             =   1755
            Width           =   945
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   3555
            TabIndex        =   23
            Top             =   375
            Width           =   420
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   555
            Index           =   3
            Left            =   1395
            Top             =   2160
            Width           =   6045
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Notes"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   315
            TabIndex        =   17
            Top             =   2175
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inst.Paid"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   315
            TabIndex        =   16
            Top             =   1740
            Width           =   615
         End
         Begin VB.Label Label31 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee ID"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   315
            TabIndex        =   15
            Top             =   375
            Width           =   900
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Department"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   3555
            TabIndex        =   14
            Top             =   810
            Width           =   825
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Refund Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   3510
            TabIndex        =   13
            Top             =   1305
            Width           =   915
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Designation"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   315
            TabIndex        =   12
            Top             =   825
            Width           =   840
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   2400
            Index           =   4
            Left            =   315
            Top             =   2835
            Width           =   7125
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Loan Refu.ID"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   315
            TabIndex        =   11
            Top             =   1305
            Width           =   960
         End
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Advance_Info As New clsAdvance_Info
Private loanInfo As New LoanRefundInfo
Private Refund As New Loan_Ref_Info
Private Refund_Rs As New Recordset
Dim SSTab_Index As Integer
Private Job_Info As New clsEmp_Job_Detail
Dim Ln_Id As String
Dim Track_Id As Long
Dim AmountPaidByLoanTaker, AmountTakenByLoanTaker, MinusBetweenTakenAndPaidValue, SlabAmountfortheLoanTaked, TotalNoofInstammentforLoan
Private Sub cboLn_Nature_Click(Index As Integer)
    Get_New_Ln_ID
End Sub
Private Sub ComboBox1_Change()
End Sub



Private Sub cmdPrint_Click()
Dim f2 As New frmLoanLedgerReport
f2.Show 1
End Sub
Private Sub cmdView_Click(Index As Integer)
On Error GoTo Erdesc
Dim f2 As New frmDataSelectforLoan
Dim getconnected As New Connection
Dim cmd As New Command
Dim myrs As New ADODB.Recordset
    getconnected.ConnectionString = strCN.Connection_String
    getconnected.Open
    cmd.ActiveConnection = getconnected
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT EMP_INFO.EMP_ID, EMP_INFO.EMP_NM FROM EMP_INFO  "
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    myrs.CursorLocation = adUseClient
    
    myrs.Open cmd.CommandText, getconnected, adOpenDynamic, adLockOptimistic
    
Select Case Index

Case 0

     Set f2.adoRecordset = myrs
     Set f2.OwnerForm = Me
     f2.Width = 6500
     f2.grdDataGrid.Columns(0).Caption = "Emp ID"
     f2.grdDataGrid.Columns(1).Caption = "Name"
     f2.grdDataGrid.Columns(0).Width = 1800
     f2.grdDataGrid.Columns(1).Width = 5500
     f2.intPutSel = 1
     f2.Show 1
     Combo1(0) = myrs.Fields(0)
     lblName(0) = myrs.Fields(1)
     
Case 1
    Set f2.adoRecordset = myrs
     Set f2.OwnerForm = Me
     f2.Width = 6500
     f2.grdDataGrid.Columns(0).Caption = "Emp ID"
     f2.grdDataGrid.Columns(1).Caption = "Name"
     f2.grdDataGrid.Columns(0).Width = 1800
     f2.grdDataGrid.Columns(1).Width = 5500
     f2.intPutSel = 1
     f2.Show 1
     Combo1(1) = myrs.Fields(0)
     lblName(1) = myrs.Fields(1)
     
Case 2

'    Set f2.adoRecordset = myrs
'     Set f2.OwnerForm = Me
'     f2.Width = 6500
'     f2.grdDataGrid.Columns(0).Caption = "Emp ID"
'     f2.grdDataGrid.Columns(1).Caption = "Name"
'     f2.grdDataGrid.Columns(0).Width = 1800
'     f2.grdDataGrid.Columns(1).Width = 5500
'     f2.intPutSel = 0
'     f2.Show 1
'     Combo1(2) = myrs.Fields(0)
'     lblName(2) = myrs.Fields(1)
End Select
Exit Sub
Erdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Combo_Click()
On Error GoTo Errdesc

Dim LoanInterestAmount, TotalAmounttoPayperMonth

If Trim(Val(Combo)) = "12" Then
    If Len((txtAmount(3))) = 0 Then txtAmount(3) = 0
    LoanInterestAmount = Val(txtAmount(3)) * 5 / 100
    TotalAmounttoPayperMonth = Val(LoanInterestAmount + txtAmount(3)) / Trim(Val(Combo))
    txtAmount(1) = Round(TotalAmounttoPayperMonth)
    txtAmount(4) = Round(LoanInterestAmount) 'set by zahid
    
ElseIf Trim(Val(Combo)) = "18" Then
    If Len((txtAmount(3))) = 0 Then txtAmount(3) = 0
    LoanInterestAmount = Val(txtAmount(3)) * 7.5 / 100
    TotalAmounttoPayperMonth = Val(LoanInterestAmount + Val(txtAmount(3))) / Trim(Val(Combo))
    txtAmount(1) = Round(TotalAmounttoPayperMonth)
    txtAmount(4) = Round(LoanInterestAmount) 'set by zahid
    
ElseIf Trim(Val(Combo)) = "24" Then
    If Len((txtAmount(3))) = 0 Then txtAmount(3) = 0
    LoanInterestAmount = Val(txtAmount(3)) * 10 / 100
    TotalAmounttoPayperMonth = Val(LoanInterestAmount + Val(txtAmount(3))) / Trim(Val(Combo))
    txtAmount(1) = Round(TotalAmounttoPayperMonth)
    txtAmount(4) = Round(LoanInterestAmount) 'set by zahid
    
End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub



Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Errdesc
If KeyCode = 13 Then
    
    Get_Employee Combo1(Index), Me, True, Index
    
    With Job_Info
        .Connstring = strCN.Connection_String
        .Emp_ID = Combo1(SSTab_Index)
        .Get_Employee
        lblName(SSTab_Index) = .Emp_Nm
       '' lblDept(SSTab_Index) = .Dept
    End With
    
    If SSTab1.Tab = 1 Then
        Get_LoanId_LOISATE_NOOFINSTA_SLABINSTALL
        ''txtLnID(1).SetFocus
    
    End If
    
    Flash_Into_Grid
    
    If SSTab1.Tab = 2 Then
        Get_Total_Slab_Amount_Pay
        'Get_LoanId_LOISATE_NOOFINSTA_SLABINSTALL
        ''txtLnID(2).SetFocus
    End If
    
    Get_Total_Amount_Has_tobe_Paid
    TotalAmount_ofMoneyLeftfor_Loan
    
    
    If SSTab1.Tab = 0 Then
        ''' txtInstall(0).SetFocus
    ElseIf SSTab1.Tab = 1 Then
        ''txtLnID(1).SetFocus
    ElseIf SSTab1.Tab = 2 Then
        ''txtLnID(2).SetFocus
    End If
End If
TabControl_For_Helping_User
             

Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub

Private Sub dtpIssue_Dt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Select Case Index
Case 0
        ''txtAdvID(0) = 0
        ''txtAmount(0).SetFocus
Case 1
   txtAmount(3).SetFocus
Case 2
    
End Select
End If
End Sub
Private Sub Form_Load()
On Error GoTo Errdesc
    SSTab_Index = 0
    Ln_Id = "0"
    Set_TabIndex
    get_Value_Into_Combo
    TabControl_For_Form_Load
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub cmdClear_Click()
    Clear_Screen
    Combo1(SSTab_Index).SetFocus
End Sub

Private Sub cmdClose_Click()
    Close_Msg Me
End Sub
Private Sub cmdSave_Click()
'On Error GoTo Errdesc
  Select Case SSTab_Index
      
      Case 0
      
        With Advance_Info
        
            .Connstring = strCN.Connection_String
            .Emp_ID = Combo1(0)
            .Adv_issue_dt = dtpIssue_Dt(1)
            .Adv_Amt = txtAmount(0)
            '''.Num_Inst = txtInstall(0)
            .Notes = txtNotes(0)
            '''.Balance = lblBalance(0)
            .Save
        End With
        
        MsgBox "Data Saved Succesfully", vbInformation, "IT Division, DNMIH"
        TabControl_For_Form_Load
        
      Case 1
      
          With loanInfo
              .Connstring = strCN.Connection_String
              .Emp_ID = Combo1(1)
              ''.Loan_Id = txtLnID(1)
              .LaonIssedDate = dtpIssue_Dt(1)
              .IssuedAmount = txtAmount(3)
              .NoOfInstallment = Trim(Combo)
              .SlabInstallmentAmount = txtAmount(1)
              .Notes = txtNotes(1)
              .Save
            End With
            
            MsgBox "Data Saved Succesfully", vbInformation, "IT Division, DNMIH"
            TabControl_For_Form_Load
            
       Case 2
    
            With loanInfo
              .Connstring = strCN.Connection_String
              .Emp_ID = Combo1(2)
              .LoanRefundedDate = dtpIssue_Dt(2)
              ''.NoOfInstallmentPaid = txtInstall(2)
              .AmountPaid = txtAmount(2)
              .Notes = txtNotes(2)
              '''.LoanRefundNo = Trim(txtLnID(2))
               .EntrDate = Date$
              .Loan_Sub_Save
          End With
          
       MsgBox "Data Saved Succesfully", vbInformation, "IT Division, DNMIH"
       TabControl_For_Form_Load
  End Select

  Flash_Into_Grid
      
' Call Clear_Screen(Me)
 ' txtEmpID(SSTab_Index).SetFocus
  Combo1(SSTab_Index).SetFocus
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
  
End Sub
Private Sub DataGrid1_Click(Index As Integer)
On Error Resume Next

If SSTab_Index = 0 Or SSTab_Index = 1 Then
    
   ' txtEmpID(SSTab_Index) = DataGrid1(0).Columns(0)
    Combo1(SSTab_Index) = DataGrid1(0).Columns(0)
    
    dtpIssue_Dt(SSTab_Index) = DataGrid1(0).Columns(2)
    txtAmount(SSTab_Index) = DataGrid1(0).Columns(3)
    
    txtNotes(SSTab_Index) = DataGrid1(0).Columns(5)
   
   
Else
'    txtRefundCode = Refund_Rs!Refund_Code
'    txtRefund = Refund_Rs!Refund_Nm
'    txtDesc(1) = Refund_Rs!Description
   
End If

    
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
Public Sub Flash_Into_Grid()
Select Case SSTab_Index

    Case 0

        With Advance_Info
            .Connstring = strCN.Connection_String
            .Emp_ID = Combo1(SSTab_Index)
        Set DataGrid1(0).DataSource = .GetAll
        
        End With
        
'            Set DataGrid1(SSTab_Index).DataSource = Loan_Info_Rs
'            DataGrid1(SSTab_Index).Refresh
        
    Case 1

'        With Loan_Info
'            .ConnString = strCN.Connection_String
'            .Ln_Nature = "Other"
'            .Emp_Id = txtEmpID(SSTab_Index)
'            Set Loan_Info_Rs = .GetAll
'        End With
'
'            Set DataGrid1(SSTab_Index).DataSource = Loan_Info_Rs
'            DataGrid1(SSTab_Index).Refresh
'
'        Case Else
'
'            With Refund
'                .ConnString = strCN.Connection_String
'                Set Refund_Rs = .GetAll
'            End With
'
'            Set DataGrid1(1).DataSource = Refund_Rs
'                 DataGrid1(1).Refresh
'
'                        With DataGrid1(1)
'                    .Columns(0).Width = 930
'                    '.Columns(0).DataField = Loan_Info_Rs!Fields(0)
'
'                    .Columns(1).Width = 3030
'                    '.Columns(1).DataField = Loan_Info_Rs!Fields(1)
'
'                    '.Columns(2).Width = 3254
'
'                End With

    End Select
       
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo Errdesc
    SSTab_Index = SSTab1.Tab
    Set_TabIndex
    get_Value_Into_Combo
    TabControl_For_Form_Load
        
        
        
    If SSTab_Index = 0 Then
        Combo1(0).SetFocus
    ElseIf SSTab_Index = 1 Then
        Combo1(1).SetFocus
    ElseIf SSTab_Index = 2 Then
        Combo1(2).SetFocus
    End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Public Sub Set_TabIndex()

On Error Resume Next

Select Case SSTab_Index
    
        Case 0
            
            'txtEmpID(SSTab_Index).TabIndex = 0
            Combo1(SSTab_Index).TabIndex = 0
            txtAmount(SSTab_Index).TabIndex = 2
            

            txtNotes(SSTab_Index).TabIndex = 6
            cmdSave.TabIndex = 7
            cmdClear.TabIndex = 8
            cmdClose.TabIndex = 9
            'txtEmpID(SSTab_Index).SetFocus
            'Combo1(SSTab_Index).SetFocus
        
        Case 1
        
        Combo.Clear
        Combo.AddItem 12
        Combo.AddItem 18
        Combo.AddItem 24
        
     'select LOANREFUNDNO from loaninformation_sub
    Dim Connect As New Connection
    Dim cmd As New Command
    Dim myrs5 As New ADODB.Recordset

    Connect.ConnectionString = strCN.Connection_String
    Connect.Open
    cmd.ActiveConnection = Connect
    cmd.CommandType = adCmdText
    cmd.CommandText = "select max(LOANREFUNDNO) from loaninformation_sub where emp_id ='" & Trim(Combo1(2)) & "'"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    myrs5.CursorLocation = adUseClient
    
    myrs5.Open cmd.CommandText, Connect, adOpenDynamic, adLockOptimistic
    
    If myrs5.BOF = False Then
        TotalNoofInstammentforLoan = myrs5.Fields(0)
       
    Else
        TotalNoofInstammentforLoan = 0
        
    End If
    

'            txtEmpID(SSTab_Index).TabIndex = 0
'            txtAmount(SSTab_Index).TabIndex = 4
'            txtInstall(SSTab_Index).TabIndex = 3
'
'            txtNotes(SSTab_Index).TabIndex = 5
'            cmdSave.TabIndex = 6
'            cmdClear.TabIndex = 8
'            cmdClose.TabIndex = 9
'            txtEmpID(SSTab_Index).SetFocus

'       combo1(1).TabIndex = 0
'       txtLnID(1).TabIndex = 1
'       dtpIssue_Dt(1).TabIndex = 2
'       txtInstall(1).TabIndex = 3
'       txtAmount(3).TabIndex = 4
'       txtNotes(1).TabIndex = 5
'       cmdSave.TabIndex = 6
'       cmdClear.TabIndex = 7
'        combo1(1).SetFocus
       
    End Select

End Sub
Private Sub Get_New_Ln_ID()
On Error GoTo Errdesc
    Dim con As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset

    con.Open strCN.Connection_String
    
    Set cmd.ActiveConnection = con
    
    cmd.CommandType = adCmdText
   ' CMD.CommandText = "select dbo.Get_Loan_ID('" & cboLn_Nature(SSTab_Index) & "')"
    Set RS = cmd.Execute
        
   ''' txtLnID(SSTab_Index) = RS.Fields(0)
            
    RS.Close
    con.Close
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Destroy Me
End Sub


Private Sub txtInsAmt_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
     KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
Private Sub txtInstall_Change(Index As Integer)
'On Error GoTo Errdesc
'Select Case Index
'Case 2
'
'    Get_Total_Amount_Has_tobe_Paid
'
'    If Trim(Len(txtInstall(2))) < 0 Then
'        txtInstall(2) = 0
'        txtInstall(1) = Val(TotalNoofInstammentforLoan) - Val(txtInstall(2))
'    Else
'        If Val(TotalNoofInstammentforLoan) < Val(txtInstall(2)) Then
'            MsgBox "Invalid Data! ", vbCritical, "Daffodil Software"
'            txtInstall(2).SetFocus
'            Exit Sub
'        Else
'            txtInstall(1) = Val(TotalNoofInstammentforLoan) - Val(txtInstall(2))
'        End If
'    End If
'
'
    
    
''End Select
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub txtInstall_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
Select Case Index
Case 0
    dtpIssue_Dt(0).SetFocus
Case 1
    txtAmount(3).SetFocus
    
Case 2

    To_Get_For_Installment_Value
    'txtNotes(2).SetFocus


End Select
End If
End Sub

Private Sub txtInstall_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    ' KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub


Private Sub txtInterest_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
Private Sub Get_Total_Slab_Amount_Pay()
On Error GoTo Errdes

Dim getconnected As New Connection
Dim cmd As New Command
Dim myrs As New ADODB.Recordset

    getconnected.ConnectionString = strCN.Connection_String
    getconnected.Open
    cmd.ActiveConnection = getconnected
    cmd.CommandType = adCmdText
    cmd.CommandText = "select sum(AMOUNTPAID) from loaninformation_sub where emp_id ='" & Trim(Combo1(2)) & "'"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    myrs.CursorLocation = adUseClient
    
    myrs.Open cmd.CommandText, getconnected, adOpenDynamic, adLockOptimistic
    
    If myrs.BOF = False Then
        AmountPaidByLoanTaker = myrs.Fields(0)
        ''lblBalance(2) = "" & AmountPaidByLoanTaker
    End If
Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
  
End Sub
Private Sub Get_Total_Amount_Has_tobe_Paid()
On Error GoTo Errdes

Dim getconnected As New Connection
Dim cmd As New Command
Dim myrs As New ADODB.Recordset

    getconnected.ConnectionString = strCN.Connection_String
    getconnected.Open
    cmd.ActiveConnection = getconnected
    cmd.CommandType = adCmdText
    
    If SSTab_Index = 0 Then
           cmd.CommandText = "select NOOFINSTALLMENT,SLABINSTALLMENTAMOUNT,ISSUEDAMOUNT from loaninformation_main where emp_id ='" & Trim(Combo1(0)) & "'"
    ElseIf SSTab_Index = 1 Then
        cmd.CommandText = "select NOOFINSTALLMENT,SLABINSTALLMENTAMOUNT,ISSUEDAMOUNT from loaninformation_main where emp_id ='" & Trim(Combo1(1)) & "'"
    ElseIf SSTab_Index = 2 Then
        cmd.CommandText = "select NOOFINSTALLMENT,SLABINSTALLMENTAMOUNT,ISSUEDAMOUNT from loaninformation_main where emp_id ='" & Trim(Combo1(2)) & "'"
    End If
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    myrs.CursorLocation = adUseClient
    
    myrs.Open cmd.CommandText, getconnected, adOpenDynamic, adLockOptimistic
    
    If myrs.BOF = False Then
        AmountTakenByLoanTaker = "" & myrs.Fields(0)
        SlabAmountfortheLoanTaked = "" & myrs.Fields(1)
        TotalNoofInstammentforLoan = "" & myrs.Fields(2)
    End If
Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub
Private Sub TotalAmount_ofMoneyLeftfor_Loan()
On Error GoTo Errdesc
If IsNull(AmountTakenByLoanTaker) Then
    AmountPaidByLoanTaker = 0
End If

If IsNull(AmountPaidByLoanTaker) Then
    AmountPaidByLoanTaker = 0
End If
txtAmount(2) = SlabAmountfortheLoanTaked
'''lblBalance(2) = Val(AmountTakenByLoanTaker) - Val(AmountPaidByLoanTaker)
MinusBetweenTakenAndPaidValue = Val(AmountTakenByLoanTaker) - Val(AmountPaidByLoanTaker)
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub txtLnID_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then

Select Case Index
    Case 0
    Case 1
        dtpIssue_Dt(1).SetFocus
    Case 2
        dtpIssue_Dt(2).SetFocus
End Select

End If
End Sub

Private Sub txtNotes_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Select Case Index
Case 0
    cmdSave.SetFocus
Case 1
    cmdSave.SetFocus
Case 2
    cmdSave.SetFocus
End Select
End If

End Sub
Private Sub To_Get_For_Installment_Value()
On Error GoTo Errdes
    Dim Connect As New Connection
    Dim cmd As New Command
    Dim myrs5 As New ADODB.Recordset
    Dim LessValue, MoreValue, EqalValue
    Dim Tracevalue As Boolean
    
    Connect.ConnectionString = strCN.Connection_String
    Connect.Open
    cmd.ActiveConnection = Connect
    cmd.CommandType = adCmdText
    cmd.CommandText = "select count(NOOFINTALLMENTPAID) from loaninformation_sub where emp_id ='" & Trim(Combo1(2)) & "'"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    myrs5.CursorLocation = adUseClient
    
    myrs5.Open cmd.CommandText, Connect, adOpenDynamic, adLockOptimistic
    
    If myrs5.BOF = False Then
    
        TotalNoofInstammentforLoan = myrs5.Fields(0)
        If IsNull(TotalNoofInstammentforLoan) Then TotalNoofInstammentforLoan = 0
        MoreValue = Val(TotalNoofInstammentforLoan) + 2
        LessValue = Val(TotalNoofInstammentforLoan) - 1
        EqalValue = Val(TotalNoofInstammentforLoan)
        
'        If Val(txtInstall(2)) >= MoreValue Then
'                MsgBox "Invalid Installment No.", vbCritical, "IT Division, DNMIH"
'                txtInstall(2).SetFocus
'                Exit Sub
'        End If
        
'        If Val(txtInstall(2)) <= LessValue Then
'                MsgBox "Invalid Installment No.", vbCritical, "IT Division, DNMIH"
'                txtInstall(2).SetFocus
'                Exit Sub
'        End If
'
'        If Val(txtInstall(2)) = EqalValue Then
'                MsgBox "Invalid Installment No.", vbCritical, "IT Division, DNMIH"
'                txtInstall(2).SetFocus
'                Exit Sub
'        End If
            
    End If
    
    Get_Total_Amount_Has_tobe_Paid
   
'   If Len(Trim(txtInstall(2))) < 0 Then
'        txtInstall(2) = 0
'        txtInstall(1) = Val(TotalNoofInstammentforLoan) - Val(txtInstall(2))
'    Else
'        txtInstall(1) = Val(TotalNoofInstammentforLoan) - Val(txtInstall(2))
'    End If

Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
 
End Sub
Private Sub Get_LoanId_LOISATE_NOOFINSTA_SLABINSTALL()
On Error GoTo Errdes
Dim getconnected As New Connection
Dim cmd As New Command
Dim myrs As New ADODB.Recordset

    getconnected.ConnectionString = strCN.Connection_String
    getconnected.Open
    cmd.ActiveConnection = getconnected
    cmd.CommandType = adCmdText
    cmd.CommandText = "select LOAN_ID,LOANISSUEDATE,ISSUEDAMOUNT,NOOFINSTALLMENT,SLABINSTALLMENTAMOUNT,notes from loaninformation_main where emp_id ='" & Trim(Combo1(1)) & "'"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    myrs.CursorLocation = adUseClient
    
    myrs.Open cmd.CommandText, getconnected, adOpenDynamic, adLockOptimistic
    
    If myrs.BOF = False Then
     '   txtLnID(1) = "" & myrs.Fields(0)
'        dtpIssue_Dt(1) = myrs.Fields(1)
'        txtAmount(3) = "" & myrs.Fields(3) 'change by zahid
'        Combo = "" & myrs.Fields(2) 'change by zahid
'        txtAmount(1) = "" & myrs.Fields(4)
'        txtNotes(1) = "" & myrs.Fields(5)
    End If

Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub
Private Sub TabControl_For_Form_Load()
On Error GoTo Errdes
    Dim getconnect As New Connection
    Dim cmd As New Command
    Dim myrs10 As New ADODB.Recordset
    
    getconnect.ConnectionString = strCN.Connection_String
    getconnect.Open
    cmd.ActiveConnection = getconnect
    cmd.CommandType = adCmdText
    
    If SSTab1.Tab = 0 Then
        cmd.CommandText = "Select * from Advance_Info order by EMP_ID,NUM_INST "
    ElseIf SSTab1.Tab = 1 Then
         'cmd.CommandText = "Select * from LoanInformation_main order by Emp_id,NOOFINSTALLMENT"
cmd.CommandText = "Select Emp_ID,Loan_Id,LOANISSUEDATE,IssuedAmount as Installment,NoOfInstallment as Amount,SlabInstallmentAmount,Notes from LoanInformation_main order by Emp_id,NOOFINSTALLMENT"
         
    ElseIf SSTab1.Tab = 2 Then
         cmd.CommandText = "Select * from LoanInformation_Sub order by Emp_id,NOOFINTALLMENTPAID"
    End If
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    myrs10.CursorLocation = adUseClient
    
    myrs10.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
    
    If SSTab1.Tab = 0 Then
        If Not (myrs10.BOF Or myrs10.EOF) Then
             Set DataGrid1(0).DataSource = myrs10
        End If
        ElseIf SSTab1.Tab = 1 Then
        
        If Not (myrs10.BOF Or myrs10.EOF) Then
                Set DataGrid1(1).DataSource = myrs10
        End If
        
        ElseIf SSTab1.Tab = 2 Then
        
        If Not (myrs10.BOF Or myrs10.EOF) Then
                Set DataGrid1(2).DataSource = myrs10
        End If
        
    End If
Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub get_Value_Into_Combo()
On Error GoTo Errdes
Dim cmd As New Command
Dim conn10 As New Connection
Dim rs10 As New Recordset

conn10.ConnectionString = strCN.Connection_String
conn10.Open
cmd.ActiveConnection = conn10
cmd.CommandType = adCmdText

cmd.CommandText = "select emp_id from emp_info order by emp_id"
rs10.CursorLocation = adUseClient
rs10.Open cmd.CommandText, conn10, adOpenDynamic, adLockOptimistic

If rs10.RecordCount > 0 Then

 If SSTab1.Tab = 0 Then
    Do Until rs10.EOF
        Combo1(0).AddItem rs10.Fields(0)
        rs10.MoveNext
    Loop
 ElseIf SSTab1.Tab = 1 Then
     Do Until rs10.EOF
        Combo1(1).AddItem rs10.Fields(0)
        rs10.MoveNext
    Loop
 ElseIf SSTab1.Tab = 2 Then
     Do Until rs10.EOF
        Combo1(2).AddItem rs10.Fields(0)
        rs10.MoveNext
    Loop
 End If
End If

rs10.Close
conn10.Close

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub
Private Sub TabControl_For_Helping_User()
On Error GoTo Errdes
    Dim getconnect As New Connection
    Dim cmd As New Command
    Dim myrs10 As New ADODB.Recordset
    
    getconnect.ConnectionString = strCN.Connection_String
    getconnect.Open
    cmd.ActiveConnection = getconnect
    cmd.CommandType = adCmdText
    
    If SSTab1.Tab = 0 Then
        cmd.CommandText = "Select * from Advance_Info  where emp_id='" & Combo1(0) & "'"
    ElseIf SSTab1.Tab = 1 Then
         'cmd.CommandText = "Select * from LoanInformation_main where emp_id='" & Combo1(1) & "'"
         cmd.CommandText = "Select Emp_ID,Loan_Id,LOANISSUEDATE,IssuedAmount as Installment,NoOfInstallment as Amount,SlabInstallmentAmount,Notes from LoanInformation_main where emp_id='" & Combo1(1) & "'"
         
    ElseIf SSTab1.Tab = 2 Then
         cmd.CommandText = "Select * from LoanInformation_Sub  where emp_id='" & Combo1(2) & "'"
    End If
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    myrs10.CursorLocation = adUseClient
    
    myrs10.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
    
    If SSTab1.Tab = 0 Then
        If Not (myrs10.BOF Or myrs10.EOF) Then
             Set DataGrid1(0).DataSource = myrs10
        End If
        ElseIf SSTab1.Tab = 1 Then
        
        If Not (myrs10.BOF Or myrs10.EOF) Then
                Set DataGrid1(1).DataSource = myrs10
        End If
        
        ElseIf SSTab1.Tab = 2 Then
        
        If Not (myrs10.BOF Or myrs10.EOF) Then
                Set DataGrid1(2).DataSource = myrs10
        End If
        
    End If
Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

