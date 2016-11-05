VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGratuity 
   BackColor       =   &H80000005&
   Caption         =   "Gratuity Fund"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form6"
   ScaleHeight     =   6645
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   5850
      Picture         =   "frmGratuity.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5910
      Width           =   1140
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   4650
      Picture         =   "frmGratuity.frx":1A82
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5910
      Width           =   1140
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   3450
      Picture         =   "frmGratuity.frx":366C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5910
      Width           =   1140
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   2250
      Picture         =   "frmGratuity.frx":5076
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5910
      Width           =   1140
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   1050
      Picture         =   "frmGratuity.frx":6A08
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5910
      Width           =   1140
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   9763
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   8388608
      TabCaption(0)   =   "Gratuity Fund Receive"
      TabPicture(0)   =   "frmGratuity.frx":839A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Gratuity Paypment"
      TabPicture(1)   =   "frmGratuity.frx":83B6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Gratuity Capital Fund"
      TabPicture(2)   =   "frmGratuity.frx":83D2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(2)"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H80000009&
         Height          =   5145
         Index           =   2
         Left            =   -75000
         TabIndex        =   42
         Top             =   330
         Width           =   7485
         Begin VB.ComboBox cmbBankName 
            Height          =   315
            Left            =   1600
            TabIndex        =   51
            Top             =   1626
            Width           =   2145
         End
         Begin VB.ComboBox cmbAccountType 
            Height          =   315
            Left            =   1600
            TabIndex        =   52
            Top             =   772
            Width           =   2145
         End
         Begin MSDataGridLib.DataGrid DataGrid3 
            Height          =   2205
            Left            =   120
            TabIndex        =   50
            Top             =   2610
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   3889
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
            Caption         =   "Gratuity Capital Fund"
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
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   11
            Left            =   1560
            Top             =   1575
            Width           =   2205
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   10
            Left            =   1560
            Top             =   720
            Width           =   2205
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Track ID"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   26
            Left            =   480
            TabIndex        =   49
            Top             =   420
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Account Type"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   25
            Left            =   480
            TabIndex        =   48
            Top             =   832
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Bank Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   24
            Left            =   480
            TabIndex        =   47
            Top             =   1656
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Amount"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   22
            Left            =   540
            TabIndex        =   46
            Top             =   2070
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Account No"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   21
            Left            =   480
            TabIndex        =   45
            Top             =   1244
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000009&
         Height          =   5235
         Index           =   0
         Left            =   30
         TabIndex        =   2
         Top             =   270
         Width           =   7458
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   6
            ItemData        =   "frmGratuity.frx":83EE
            Left            =   4560
            List            =   "frmGratuity.frx":83F0
            TabIndex        =   12
            Top             =   2100
            Width           =   2335
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   5
            ItemData        =   "frmGratuity.frx":83F2
            Left            =   4560
            List            =   "frmGratuity.frx":83F4
            TabIndex        =   13
            Top             =   2550
            Width           =   2355
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   4
            ItemData        =   "frmGratuity.frx":83F6
            Left            =   1530
            List            =   "frmGratuity.frx":83F8
            TabIndex        =   14
            Top             =   2148
            Width           =   1905
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   4980
            TabIndex        =   23
            Top             =   300
            Width           =   1905
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "frmGratuity.frx":83FA
            Left            =   1560
            List            =   "frmGratuity.frx":83FC
            TabIndex        =   22
            Top             =   792
            Width           =   2210
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   1995
            Left            =   180
            TabIndex        =   16
            Top             =   3090
            Width           =   7125
            _ExtentX        =   12568
            _ExtentY        =   3519
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
            Caption         =   "Gratuity Fund Receive"
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
         Begin MSComCtl2.DTPicker DPReceiveDate 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   24
            Top             =   1686
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   8388608
            CalendarTitleForeColor=   8388608
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   63963139
            CurrentDate     =   36998
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   6
            Left            =   4530
            Top             =   2050
            Width           =   2415
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   5
            Left            =   4530
            Top             =   2520
            Width           =   2415
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   4
            Left            =   1530
            Top             =   2100
            Width           =   1905
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Account Type"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   19
            Left            =   210
            TabIndex        =   15
            Top             =   2190
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Voucher No"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   18
            Left            =   3990
            TabIndex        =   34
            Top             =   1740
            Width           =   855
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   1
            Left            =   4950
            Top             =   270
            Width           =   1965
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   0
            Left            =   1530
            Top             =   750
            Width           =   2265
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Bank Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   3600
            TabIndex        =   11
            Top             =   2640
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Account No"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   3600
            TabIndex        =   10
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Check No "
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   4050
            TabIndex        =   9
            Top             =   870
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Payment Type"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   3795
            TabIndex        =   8
            Top             =   330
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Received Amount"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   210
            TabIndex        =   7
            Top             =   2640
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Receive Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   210
            TabIndex        =   6
            Top             =   1740
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Description"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   210
            TabIndex        =   5
            Top             =   1290
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Source of Fund "
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   210
            TabIndex        =   4
            Top             =   840
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Receive ID"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   3
            Top             =   390
            Width           =   810
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000009&
         Height          =   5205
         Index           =   1
         Left            =   -75000
         TabIndex        =   1
         Top             =   300
         Width           =   7485
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   9
            Left            =   1800
            TabIndex        =   35
            Top             =   2082
            Width           =   1785
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   8
            Left            =   4740
            TabIndex        =   36
            Top             =   2520
            Width           =   2145
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   7
            Left            =   4740
            TabIndex        =   41
            Top             =   2085
            Width           =   2145
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   1900
            Left            =   120
            TabIndex        =   39
            Top             =   3000
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   3360
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
            Caption         =   "Gratuity Payment"
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
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   3
            Left            =   4980
            TabIndex        =   38
            Top             =   330
            Width           =   1905
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   1800
            TabIndex        =   37
            Top             =   798
            Width           =   1785
         End
         Begin MSComCtl2.DTPicker DPReceiveDate 
            Height          =   315
            Index           =   1
            Left            =   1800
            TabIndex        =   40
            Top             =   1650
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   8388608
            CalendarTitleForeColor=   8388608
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   63963139
            CurrentDate     =   36998
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   345
            Index           =   12
            Left            =   1770
            Top             =   2070
            Width           =   1845
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   9
            Left            =   4700
            Top             =   2490
            Width           =   2155
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   8
            Left            =   1770
            Top             =   1620
            Width           =   1830
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   7
            Left            =   4700
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Account Type"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   23
            Left            =   210
            TabIndex        =   43
            Top             =   2112
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Voucher No"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   20
            Left            =   3720
            TabIndex        =   44
            Top             =   1740
            Width           =   855
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   3
            Left            =   4950
            Top             =   300
            Width           =   1965
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   2
            Left            =   1770
            Top             =   750
            Width           =   1845
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Bank Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   3750
            TabIndex        =   33
            Top             =   2550
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Account No"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   3750
            TabIndex        =   32
            Top             =   2100
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Check No"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   3840
            TabIndex        =   31
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Payment Type"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   3840
            TabIndex        =   30
            Top             =   390
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Amount"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   29
            Top             =   2550
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Payment Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   210
            TabIndex        =   28
            Top             =   1674
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Description"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   210
            TabIndex        =   27
            Top             =   1236
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Purpose of Payment"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   210
            TabIndex        =   26
            Top             =   798
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Gratuity Payment ID"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   210
            TabIndex        =   25
            Top             =   360
            Width           =   1410
         End
      End
   End
End
Attribute VB_Name = "frmGratuity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private oGratuity As New clsGratuity
'Private oGratuityPayment As New clsGratuityPayment
'Private oGraCapitalFund As New clsGratuityCapitalFund
'Dim SSTab_Index As Integer
'
'
'
'Private Sub cmdClear_Click()
'
'Clear_Screen
'
'End Sub
'
'Private Sub cmdClose_Click()
'Close_Msg Me
'End Sub
'
'
'Private Sub cmdDelete_Click()
'On Error GoTo Errdesc
'Select Case SSTab_Index
'Case 0
'        With oGratuity
'            .Connstring = strCN.Connection_String
'            .GratuityReceiveId = txtField(0)
'            .Delete
'        End With
'        MsgBox "Data Deleted Successfully", vbInformation, "IT Division, DNMIH"
'        Clear_Screen
'        Combo1(0).SetFocus
'Case 1
'        With oGratuityPayment
'            .Connstring = strCN.Connection_String
'            .GratuityPaymentId = txtField(6)
'            .Delete
'        End With
'        MsgBox "Data Deleted Successfully", vbInformation, "IT Division, DNMIH"
'        Clear_Screen
'        Combo1(2).SetFocus
'Case 2
'        With oGraCapitalFund
'            .Connstring = strCN.Connection_String
'            .TrackId = txtField(12)
'            .Delete
'        End With
'
'        MsgBox "Data Deleted Successfully", vbInformation, "IT Division, DNMIH"
'        Clear_Screen
'        cmbAccountType.SetFocus
'
'
'End Select
'Show_Data_Form_Load
'Exit Sub
'Errdesc:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'
'End Sub
'
'Private Sub cmdPrint_Click()
'Dim f As New frmGratuityReport
'f.Show 1
'
'End Sub
'
'Private Sub cmdSave_Click()
'  Select Case SSTab_Index
'
'      Case 0
'
'        With oGratuity
'
'            .Connstring = strCN.Connection_String
'            .GratuityReceiveId = txtField(0)
'            .SourceOfFund = Get_Code(Combo1(0).Text)
'            .Description = txtField(1)
'            .PaymentReceiveDate = DPReceiveDate(0)
'            .ReceiveAmount = txtField(2)
'            .PaymentReceivedType = Combo1(1)
'            .CheckNo = txtField(3)
'            .AccountNo = Combo1(6)
'            .BankCode = Get_Code(Combo1(5).Text)
'            .VoucherNo = txtField(15)
'            .AccountType = Get_Code(Combo1(4).Text)
'            .Save
'        End With
'
'        MsgBox "Data Saved Succesfully", vbInformation, "IT Division, DNMIH"
'        'TabControl_For_Form_Load
'
'      Case 1
'
'          With oGratuityPayment
'            .Connstring = strCN.Connection_String
'            .GratuityPaymentId = txtField(6)
'            .PaymentPurpose = Get_Code(Combo1(2).Text)
'            .Description = txtField(7)
'            .PaymentDate = DPReceiveDate(1)
'            .Amount = txtField(8)
'            .PaymentType = Combo1(3)
'            .CheckNo = txtField(9)
'            .AccountNo = Combo1(7)
'            .BankCode = Get_Code(Combo1(8))
'            .VoucherNo = txtField(4)
'            .AccountType = Get_Code(Combo1(9))
'            .Save
'            End With
'
'            MsgBox "Data Saved Succesfully", vbInformation, "IT Division, DNMIH"
'           ' TabControl_For_Form_Load
'
'    Case 2
'
'            With oGraCapitalFund
'            .Connstring = strCN.Connection_String
'            .TrackId = txtField(12)
'            .AccountType = Get_Code(cmbAccountType.Text)
'            .AccountNo = txtField(13)
'            .BankCode = Get_Code(cmbBankName.Text)
'            .Amount = txtField(14)
'            .Save
'            End With
'
'        MsgBox "Data Saved Succesfully", vbInformation, "IT Division, DNMIH"
'
'
'  End Select
'Show_Data_Form_Load
'
'Exit Sub
'Errdesc:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'
'End Sub
'
'
'
'Private Sub DataGrid1_Click()
'On Error GoTo Errdes
'txtField(0) = DataGrid1.Columns(0)
'Combo1(0).Text = DataGrid1.Columns(1) + "~" + DataGrid1.Columns(11)
'txtField(1).Text = DataGrid1.Columns(2)
'DPReceiveDate(0).Value = DataGrid1.Columns(3)
'txtField(2).Text = DataGrid1.Columns(4)
'Combo1(1).Text = DataGrid1.Columns(5)
'txtField(3).Text = DataGrid1.Columns(6)
'Combo1(6).Text = DataGrid1.Columns(7)
'Combo1(5).Text = DataGrid1.Columns(8) + "~" + DataGrid1.Columns(13)
'txtField(15).Text = DataGrid1.Columns(9)
'Combo1(4).Text = DataGrid1.Columns(10) + "~" + DataGrid1.Columns(12)
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'
'End Sub
'Private Sub DataGrid2_Click()
'On Error GoTo Errdes
'txtField(6) = DataGrid2.Columns(0)
'Combo1(2).Text = DataGrid2.Columns(1) + "~" + DataGrid2.Columns(11)
'txtField(7).Text = DataGrid2.Columns(2)
'DPReceiveDate(1).Value = DataGrid2.Columns(3)
'txtField(8).Text = DataGrid2.Columns(4)
'Combo1(3).Text = DataGrid2.Columns(5)
'txtField(9).Text = DataGrid2.Columns(6)
'Combo1(7).Text = DataGrid2.Columns(7)
'Combo1(8).Text = DataGrid2.Columns(8) + "~" + DataGrid2.Columns(13)
'txtField(4).Text = DataGrid2.Columns(9)
'Combo1(9).Text = DataGrid2.Columns(10) + "~" + DataGrid2.Columns(12)
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'
'End Sub
'
'Private Sub DataGrid3_Click()
'On Error GoTo Errdes
'txtField(12) = DataGrid3.Columns(0)
'cmbAccountType.Text = DataGrid3.Columns(1) + "~" + DataGrid3.Columns(5)
'txtField(13).Text = DataGrid3.Columns(2)
'cmbBankName.Text = DataGrid3.Columns(3) + "~" + DataGrid3.Columns(6)
'txtField(14).Text = DataGrid3.Columns(4)
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'
'End Sub
'
'
'Private Sub Form_Load()
'On Error GoTo Errdes
'    Screen_Position Me
'    SSTab_Index = 0
'    Set_TabIndex
'    get_Value_Into_Payment_Purpose
'    LOAD_PAYMENT_RECE_TYPE Me
'    LOAD_PAYMENT_TYPE Me
'    get_Value_Into_Account_type
'    get_Value_Into_Bank_Name
'    get_Value_Into_Account_No
'    Dim cmd As New Command
'    Dim conn1 As New Connection
'    Dim RS As New Recordset
'
'    conn1.ConnectionString = strCN.Connection_String
'    conn1.Open
'    cmd.ActiveConnection = conn1
'    cmd.CommandType = adCmdText
'
'    cmd.CommandText = "select SOURCE_ID,SOURCE_NAME from L_SOURCE_OF_FUND order by SOURCE_ID"
'    RS.CursorLocation = adUseClient
'    RS.Open cmd.CommandText, conn1, adOpenDynamic, adLockOptimistic
'
'        If RS.RecordCount > 0 Then
'            Do Until RS.EOF
'            Combo1(0).AddItem RS.Fields(1) & "~" & RS.Fields(0)
'            RS.MoveNext
'            Loop
'
'        End If
'
'    RS.Close
'    conn1.Close
' SSTab1.Tab = 0
'Show_Data_Form_Load
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'
'End Sub
'
'Private Sub Show_Data_Form_Load()
'On Error GoTo Errdes
'    Dim getconnect As New Connection
'    Dim cmd As New Command
'    Dim myrs10 As New ADODB.Recordset
'
'    getconnect.ConnectionString = strCN.Connection_String
'    getconnect.Open
'    cmd.ActiveConnection = getconnect
'    cmd.CommandType = adCmdText
'
'    If SSTab1.Tab = 0 Then
'        'cmd.CommandText = "Select GRA_RECEIVE_ID as ReceiveID,SOURCE_OF_GRATUITY as SourceOfFund,Description,PAYMENT_RECE_DATE as ReceiveDate,RECEIVED_AMOUNT ReceiveAmount,PAYMENT_RECE_TYPE as ReceiveType,CHECK_NO as CheckNo,ACCOUNT_NO as AccountNo,BANK_CODE as BankName from GRATUITY_RECEIVE order by GRA_RECEIVE_ID"
'
'        cmd.CommandText = "Select A.GRA_RECEIVE_ID as ReceiveId,LS.SOURCE_NAME as SourceOfFund," _
'            & "A.Description,A.PAYMENT_RECE_DATE as ReceiveDate,A.RECEIVED_AMOUNT ReceiveAmount," _
'            & "A.PAYMENT_RECE_TYPE as ReceiveType,A.CHECK_NO as CheckNo,A.ACCOUNT_NO as AccountNo," _
'            & "LB.BANK_NAME as BankName,A.VOUCHER_NO AS VoucherNo,LA.TYPE_NAME AS TypeName," _
'            & "LS.SOURCE_ID AS SourceId,LA.TYPE_ID AS TypeId,LB.BANK_ID AS BankId" _
'            & " From GRATUITY_RECEIVE A ,L_SOURCE_OF_FUND LS,L_BANK LB,L_ACCOUNT_TYPE LA" _
'            & " Where A.SOURCE_OF_GRATUITY = LS.SOURCE_ID AND A.ACCOUNT_TYPE=LA.TYPE_ID" _
'            & " AND A.BANK_CODE=LB.BANK_ID order by A.GRA_RECEIVE_ID"
'
'    ElseIf SSTab1.Tab = 1 Then
'        'cmd.CommandText = "Select GRA_PAYMENT_ID as PaymentID,PURPOSE_OF_PAYMENT as PaymentPurpose,Description,PAYMENT_DATE as PaymentDate,AMOUNT,PAYMENT_TYPE as PaymentType,CHECK_NO as CheckNo,ACCOUNT_NO as AccountNo,BANK_NAME as BankName from GRATUITY_PAYMENT order by GRA_PAYMENT_ID"
'
''        cmd.CommandText = "Select A.GRA_PAYMENT_ID as PaymentID," _
''        & "(SELECT LP.PURPOSE_NAME FROM L_GRA_PAYMENT_PURPOSE LP WHERE A.PURPOSE_OF_PAYMENT=LP.PURPOSE_ID) as PaymentPurpose," _
''        & "A.Description,A.PAYMENT_DATE as PaymentDate," _
''        & "A.AMOUNT,A.PAYMENT_TYPE as PaymentType,A.CHECK_NO as CheckNo," _
''        & "A.ACCOUNT_NO as AccountNo,A.BANK_NAME as BankName," _
''        & "(SELECT LP.PURPOSE_ID FROM L_GRA_PAYMENT_PURPOSE LP WHERE A.PURPOSE_OF_PAYMENT=LP.PURPOSE_ID) as PurposeId" _
''        & "from GRATUITY_PAYMENT A order by A.GRA_PAYMENT_ID"
'
'        cmd.CommandText = "SELECT A.GRA_PAYMENT_ID as PaymentID,LP.PURPOSE_NAME as PaymentPurpose," _
'            & "A.Description,A.PAYMENT_DATE as PaymentDate,A.AMOUNT as PaymentAmount," _
'            & "A.PAYMENT_TYPE as PaymentType,A.CHECK_NO as CheckNo,A.ACCOUNT_NO as AccountNo," _
'            & "LB.BANK_NAME as BankName,A.VOUCHER_NO AS VoucherNo,LA.TYPE_NAME AS TypeName," _
'            & "LP.PURPOSE_ID as PurposeId,LA.TYPE_ID AS TypeId,LB.BANK_ID AS BankId" _
'            & " From GRATUITY_PAYMENT A,L_GRA_PAYMENT_PURPOSE LP,L_BANK LB,L_ACCOUNT_TYPE LA" _
'            & " Where A.PURPOSE_OF_PAYMENT = LP.PURPOSE_ID AND A.ACCOUNT_TYPE=LA.TYPE_ID" _
'            & " AND A.BANK_CODE=LB.BANK_ID order by A.GRA_PAYMENT_ID"
'
'
'    ElseIf SSTab1.Tab = 2 Then
'       'cmd.CommandText = "Select Track_Id as TrackId,ACCOUNT_TYPE as AccountType,ACCOUNT_NO as AccountNo,BANK_NAME as BankName,AMOUNT as Amount from GRA_CAPITAL_FUND order by TRACK_ID"
'       'cmd.CommandText = "Select A.Track_Id as TrackId, (SELECT B.TYPE_NAME FROM L_ACCOUNT_TYPE B WHERE A.ACCOUNT_TYPE=B.TYPE_ID)  AS AccountType, ACCOUNT_NO as AccountNo,(SELECT LB.BANK_NAME FROM L_BANK LB WHERE A.BANK_NAME=LB.BANK_ID) as BankName,AMOUNT as Amount,(SELECT B.TYPE_ID FROM L_ACCOUNT_TYPE B WHERE A.ACCOUNT_TYPE=B.TYPE_ID) as  TYPEID,(SELECT LB.BANK_ID FROM L_BANK LB WHERE A.BANK_NAME=LB.BANK_ID) as BankId from GRA_CAPITAL_FUND A  order by A.TRACK_ID"
'        cmd.CommandText = "Select A.Track_Id as TrackId," _
'        & "B.TYPE_NAME AS AccountType,A.ACCOUNT_NO as AccountNo," _
'        & "LB.BANK_NAME as BankName,A.AMOUNT as Amount," _
'        & "B.TYPE_ID as TypeId,LB.BANK_ID as BankId" _
'        & " From GRA_CAPITAL_FUND A,L_ACCOUNT_TYPE B,L_BANK LB" _
'        & " Where A.ACCOUNT_TYPE = B.TYPE_ID AND A.BANK_CODE=LB.BANK_ID ORDER BY A.TRACK_ID"
'
'
'    End If
'
'    cmd.Properties("iRowsetChange") = True
'    cmd.Properties("updatability") = 7
'    myrs10.CursorLocation = adUseClient
'
'    myrs10.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
'
'    If SSTab1.Tab = 0 Then
'        If Not (myrs10.BOF Or myrs10.EOF) Then
'             Set DataGrid1.DataSource = myrs10
'
'        End If
'    ElseIf SSTab1.Tab = 1 Then
'        If Not (myrs10.BOF Or myrs10.EOF) Then
'                Set DataGrid2.DataSource = myrs10
'        End If
'
'    ElseIf SSTab1.Tab = 2 Then
'        If Not (myrs10.BOF Or myrs10.EOF) Then
'                Set DataGrid3.DataSource = myrs10
'        End If
'
'
'    End If
'Exit Sub
'Errdes:
' MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'
'Private Sub SSTab1_Click(PreviousTab As Integer)
'On Error GoTo Errdesc
'    SSTab_Index = SSTab1.Tab
'    Set_TabIndex
'    'get_Value_Into_Payment_Purpose
'    'TabControl_For_Form_Load
'    Show_Data_Form_Load
'    If SSTab_Index = 1 Then
'    Combo1(3).SetFocus
'    End If
'Exit Sub
'Errdesc:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
'If KeyCode = 13 Then
'Show_Data
''Set_TabIndex
'End If
'End Sub
'Private Sub Show_Data()
'On Error GoTo Errdes
'    Dim getconnect As New Connection
'    Dim cmd As New Command
'    Dim myrs10 As New ADODB.Recordset
'    getconnect.ConnectionString = strCN.Connection_String
'    getconnect.Open
'    cmd.ActiveConnection = getconnect
'    cmd.CommandType = adCmdText
'
'    If SSTab1.Tab = 0 Then
'        cmd.CommandText = "Select GRA_RECEIVE_ID as ReceiveID,SOURCE_OF_GRATUITY as SourceOfFund,Description,PAYMENT_RECE_DATE as ReceiveDate,RECEIVED_AMOUNT ReceiveAmount,PAYMENT_RECE_TYPE as ReceiveType,CHECK_NO as CheckNo,ACCOUNT_NO as AccountNo,BANK_NAME as BankName from GRATUITY_RECEIVE where GRA_RECEIVE_ID='" & txtField(0).Text & "'"
'    ElseIf SSTab1.Tab = 1 Then
'        cmd.CommandText = "Select GRA_PAYMENT_ID as PaymentID,PURPOSE_OF_PAYMENT as PaymentPurpose,Description,PAYMENT_DATE as PaymentDate,AMOUNT,PAYMENT_TYPE as PaymentType,CHECK_NO as CheckNo,ACCOUNT_NO as AccountNo,BANK_NAME as BankName from GRATUITY_PAYMENT where GRA_PAYMENT_ID='" & txtField(6).Text & "'"
'
'    ElseIf SSTab1.Tab = 1 Then
'        cmd.CommandText = "Select  TRACK_ID, ACCOUNT_TYPE, ACCOUNT_NO, bank_name, AMOUNT from gra_capital_fund where TRACK_ID='" & txtField(12).Text & "'"
'
'    End If
'    cmd.Properties("iRowsetChange") = True
'    cmd.Properties("updatability") = 7
'    myrs10.CursorLocation = adUseClient
'    myrs10.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
'    If SSTab1.Tab = 0 Then
'        If Not (myrs10.BOF Or myrs10.EOF) Then
'              txtField(0) = myrs10(0)
'              Combo1(0) = myrs10(1)
'              txtField(1) = myrs10(2)
'              DPReceiveDate(0) = myrs10(3)
'              txtField(2) = myrs10(4)
'              Combo1(1) = myrs10(5)
'              txtField(3) = "" & myrs10(6)
'              txtField(4) = myrs10(7)
'              txtField(5) = myrs10(8)
'        End If
'    ElseIf SSTab1.Tab = 1 Then
'        If Not (myrs10.BOF Or myrs10.EOF) Then
'               ' Set DataGrid2.DataSource = myrs10
'        End If
'
'    End If
'Exit Sub
'Errdes:
' MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'
'Private Sub get_Value_Into_Payment_Purpose()
'On Error GoTo Errdes
'    Dim getconnect As New Connection
'    Dim cmd As New Command
'    Dim RS As New ADODB.Recordset
'    getconnect.ConnectionString = strCN.Connection_String
'    getconnect.Open
'    cmd.ActiveConnection = getconnect
'    cmd.CommandType = adCmdText
'
'    cmd.CommandText = "select PURPOSE_ID,PURPOSE_NAME from L_GRA_PAYMENT_PURPOSE order by PURPOSE_ID"
'
'    cmd.Properties("iRowsetChange") = True
'    cmd.Properties("updatability") = 7
'    RS.CursorLocation = adUseClient
'    RS.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
'       If RS.RecordCount > 0 Then
'            Do Until RS.EOF
'            Combo1(2).AddItem RS.Fields(1) & "~" & RS.Fields(0)
'            RS.MoveNext
'            Loop
'        End If
'
'
'
'    Exit Sub
'Errdes:
' MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub get_Value_Into_Account_type()
'On Error GoTo Errdes
'    Dim getconnect As New Connection
'    Dim cmd As New Command
'    Dim RS As New ADODB.Recordset
'    getconnect.ConnectionString = strCN.Connection_String
'    getconnect.Open
'    cmd.ActiveConnection = getconnect
'    cmd.CommandType = adCmdText
'
'    cmd.CommandText = "select TYPE_ID,TYPE_NAME from L_ACCOUNT_TYPE order by TYPE_ID"
'
'    cmd.Properties("iRowsetChange") = True
'    cmd.Properties("updatability") = 7
'    RS.CursorLocation = adUseClient
'    RS.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
'       If RS.RecordCount > 0 Then
'            Do Until RS.EOF
'            cmbAccountType.AddItem RS.Fields(1) & "~" & RS.Fields(0)
'            Combo1(4).AddItem RS.Fields(1) & "~" & RS.Fields(0)
'            Combo1(9).AddItem RS.Fields(1) & "~" & RS.Fields(0)
'            RS.MoveNext
'            Loop
'            'cmbAccountType.ListIndex = 0
'        End If
'
'
'
'    Exit Sub
'Errdes:
' MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub get_Value_Into_Bank_Name()
'On Error GoTo Errdes
'    Dim getconnect As New Connection
'    Dim cmd As New Command
'    Dim RS As New ADODB.Recordset
'    getconnect.ConnectionString = strCN.Connection_String
'    getconnect.Open
'    cmd.ActiveConnection = getconnect
'    cmd.CommandType = adCmdText
'
'    cmd.CommandText = "select BANK_ID,BANK_NAME from L_BANK order by BANK_ID"
'
'    cmd.Properties("iRowsetChange") = True
'    cmd.Properties("updatability") = 7
'    RS.CursorLocation = adUseClient
'    RS.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
'       If RS.RecordCount > 0 Then
'            Do Until RS.EOF
'            cmbBankName.AddItem RS.Fields(1) & "~" & RS.Fields(0)
'            Combo1(5).AddItem RS.Fields(1) & "~" & RS.Fields(0)
'            Combo1(8).AddItem RS.Fields(1) & "~" & RS.Fields(0)
'            RS.MoveNext
'            Loop
'        End If
'
'
'
'    Exit Sub
'Errdes:
' MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub get_Value_Into_Account_No()
'On Error GoTo Errdes
'    Dim getconnect As New Connection
'    Dim cmd As New Command
'    Dim RS As New ADODB.Recordset
'    getconnect.ConnectionString = strCN.Connection_String
'    getconnect.Open
'    cmd.ActiveConnection = getconnect
'    cmd.CommandType = adCmdText
'
'    cmd.CommandText = "select ACCOUNT_NO from GRA_CAPITAL_FUND"
'
'    cmd.Properties("iRowsetChange") = True
'    cmd.Properties("updatability") = 7
'    RS.CursorLocation = adUseClient
'    RS.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
'       If RS.RecordCount > 0 Then
'            Do Until RS.EOF
'            Combo1(6).AddItem RS.Fields(0)
'            Combo1(7).AddItem RS.Fields(0)
'            RS.MoveNext
'            Loop
'        End If
'
'
'
'    Exit Sub
'Errdes:
' MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'
'
'Public Sub Set_TabIndex()
'On Error GoTo Errdes
'
'Select Case SSTab_Index
'
'Case 0  'Gratuity Receive
'
'                'txtEmp_ID(0).TabIndex = 0
'    Combo1(1).TabIndex = 0
'    Combo1(0).TabIndex = 1
'    txtField(3).TabIndex = 2
'    txtField(1).TabIndex = 3
'    DPReceiveDate(0).TabIndex = 4
'    Combo1(6).TabIndex = 5
'    txtField(2).TabIndex = 6
'    Combo1(5).TabIndex = 7
'    cmdSave.TabIndex = 8
'    cmdClear.TabIndex = 9
'    cmdDelete.TabIndex = 10
'    cmdPrint.TabIndex = 11
'    cmdClear.TabIndex = 12
'
'
'
'Case 1  'Gratuity Payment
'
''    Combo1(3).TabIndex = 0
''    Combo1(2).TabIndex = 1
''    txtField(9).TabIndex = 2
''    txtField(7).TabIndex = 3
''    DPReceiveDate(1).TabIndex = 4
''    txtField(4).TabIndex = 5
''    txtField(8).TabIndex = 6
''    txtField(11).TabIndex = 7
'Case 2  ' Gartuity Capital fund
'   'txtField(12).TabIndex = 0
'    cmbAccountType.TabIndex = 0
'    txtField(13).TabIndex = 1
'    cmbBankName.TabIndex = 2
'    txtField(14).TabIndex = 3
'
'End Select
'
'    'txtEmp_ID(SSTab_Index).SetFocus
'    'txtEmp_ID(SSTab_Index).SelStart = Len(txtEmp_ID(SSTab_Index))
'
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'
