VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Employee Information"
   ClientHeight    =   6885
   ClientLeft      =   810
   ClientTop       =   1500
   ClientWidth     =   10395
   HelpContextID   =   101
   Icon            =   "frmEmp_Information.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog ComDiag 
      Left            =   135
      Top             =   6255
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   4635
      Picture         =   "frmEmp_Information.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   6300
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   5925
      Picture         =   "frmEmp_Information.frx":22D4
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   6300
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   1980
      Picture         =   "frmEmp_Information.frx":3EBE
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   6300
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   3300
      Picture         =   "frmEmp_Information.frx":5850
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   6300
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   7245
      Picture         =   "frmEmp_Information.frx":71E2
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   6300
      Width           =   1185
   End
   Begin VB.ListBox lstTips 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3FEFF&
      ForeColor       =   &H000000C0&
      Height          =   225
      ItemData        =   "frmEmp_Information.frx":8C64
      Left            =   180
      List            =   "frmEmp_Information.frx":8C66
      TabIndex        =   0
      Top             =   6210
      Visible         =   0   'False
      Width           =   1680
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6150
      Left            =   60
      TabIndex        =   1
      Top             =   45
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   10848
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   8908
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Personal Information"
      TabPicture(0)   =   "frmEmp_Information.frx":8C68
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Job Details"
      TabPicture(1)   =   "frmEmp_Information.frx":8C84
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   5805
         Left            =   0
         TabIndex        =   44
         Top             =   330
         Width           =   10275
         Begin VB.CommandButton MaxSerialButton 
            Appearance      =   0  'Flat
            BackColor       =   &H00915411&
            Caption         =   "Show Max Serial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   7800
            MaskColor       =   &H00915411&
            TabIndex        =   107
            Top             =   1230
            Width           =   1575
         End
         Begin VB.TextBox txtDesignationLevel 
            Height          =   405
            Left            =   6090
            Locked          =   -1  'True
            TabIndex        =   106
            Top             =   1260
            Width           =   975
         End
         Begin VB.TextBox txtStaffSerial 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   420
            Left            =   7170
            MaxLength       =   4
            TabIndex        =   105
            Text            =   "1"
            Top             =   1260
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpParmanent_Dt 
            Height          =   330
            Left            =   6165
            TabIndex        =   103
            Top             =   1800
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   582
            _Version        =   393216
            Format          =   64356353
            CurrentDate     =   38908
         End
         Begin VB.ComboBox Combo2 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "frmEmp_Information.frx":8CA0
            Left            =   6255
            List            =   "frmEmp_Information.frx":8CB0
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   3195
            Width           =   2400
         End
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1950
            TabIndex        =   100
            Top             =   5370
            Width           =   2415
         End
         Begin VB.ComboBox cboBank 
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   6255
            TabIndex        =   90
            Text            =   "cboBank"
            Top             =   3960
            Width           =   2400
         End
         Begin VB.ComboBox cboBankBranch 
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1935
            TabIndex        =   89
            Text            =   "cboBankBranch"
            Top             =   4410
            Width           =   3525
         End
         Begin VB.ComboBox cboScale 
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1935
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   3195
            Width           =   2400
         End
         Begin VB.TextBox txtPrevious_Balance 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   6255
            MultiLine       =   -1  'True
            TabIndex        =   87
            Top             =   5265
            Width           =   2385
         End
         Begin VB.TextBox txtEmp_ID 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   1
            Left            =   1935
            MaxLength       =   10
            TabIndex        =   85
            Top             =   405
            Width           =   2400
         End
         Begin VB.TextBox txtAccountNo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   6300
            MultiLine       =   -1  'True
            TabIndex        =   84
            Top             =   4500
            Width           =   2385
         End
         Begin VB.TextBox txtPF_Mem_No 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   6255
            MultiLine       =   -1  'True
            TabIndex        =   83
            Top             =   4905
            Width           =   2430
         End
         Begin VB.TextBox txtBasic 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1935
            MultiLine       =   -1  'True
            TabIndex        =   82
            Top             =   3645
            Width           =   2385
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   2025
            TabIndex        =   71
            Top             =   4005
            Width           =   2265
            Begin VB.OptionButton optPayment_Mode 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Cash"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   73
               Top             =   0
               Value           =   -1  'True
               Width           =   780
            End
            Begin VB.OptionButton optPayment_Mode 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Bank"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   1
               Left            =   1230
               TabIndex        =   72
               Top             =   0
               Width           =   780
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   2070
            TabIndex        =   68
            Top             =   4905
            Width           =   2130
            Begin VB.OptionButton optPF 
               BackColor       =   &H00FFFFFF&
               Caption         =   "No"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   0
               Left            =   1200
               TabIndex        =   70
               Top             =   0
               Width           =   645
            End
            Begin VB.OptionButton optPF 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Yes"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   69
               Top             =   0
               Value           =   -1  'True
               Width           =   825
            End
         End
         Begin VB.TextBox txtResponsibility 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   1935
            MultiLine       =   -1  'True
            TabIndex        =   50
            Top             =   2745
            Width           =   7875
         End
         Begin VB.TextBox txtService_Bk_No 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   6165
            TabIndex        =   49
            Top             =   2295
            Width           =   2205
         End
         Begin VB.TextBox txtFile_Ref_No 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1935
            TabIndex        =   48
            Top             =   2295
            Width           =   2385
         End
         Begin VB.ComboBox cboDesig 
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1935
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   855
            Width           =   2400
         End
         Begin VB.ComboBox cboDept 
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   6075
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   855
            Width           =   2955
         End
         Begin VB.ComboBox cboType 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "frmEmp_Information.frx":8CEA
            Left            =   1935
            List            =   "frmEmp_Information.frx":8CEC
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   1305
            Width           =   2400
         End
         Begin MSComCtl2.DTPicker dtpEmp_join_date 
            Height          =   315
            Left            =   1935
            TabIndex        =   51
            Top             =   1800
            Width           =   2400
            _ExtentX        =   4233
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
            Format          =   64356355
            CurrentDate     =   36998
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Desig. Serial :"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4650
            TabIndex        =   104
            Top             =   1350
            Width           =   975
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   390
            Index           =   30
            Left            =   6210
            Top             =   3150
            Width           =   2490
         End
         Begin VB.Label Label31 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Class of Emp"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   4770
            TabIndex        =   101
            Top             =   3240
            Width           =   915
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   285
            Index           =   2
            Left            =   1890
            Top             =   5340
            Width           =   2490
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "BMDC Registration"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   99
            Top             =   5370
            Width           =   1350
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   375
            Index           =   23
            Left            =   6210
            Top             =   3915
            Width           =   2490
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   375
            Index           =   22
            Left            =   1890
            Top             =   4365
            Width           =   3615
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   375
            Index           =   21
            Left            =   1890
            Top             =   3150
            Width           =   2490
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   285
            Index           =   20
            Left            =   6210
            Top             =   5220
            Width           =   2490
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Previous Balance"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   4680
            TabIndex        =   86
            Top             =   5265
            Width           =   1245
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   285
            Index           =   19
            Left            =   6210
            Top             =   4455
            Width           =   2490
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   285
            Index           =   15
            Left            =   6210
            Top             =   4860
            Width           =   2490
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   285
            Index           =   14
            Left            =   1890
            Top             =   3600
            Width           =   2490
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pay Scale"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   360
            TabIndex        =   81
            Top             =   3195
            Width           =   720
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Basic Salary"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   80
            Top             =   3600
            Width           =   870
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Mode of Payment "
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   4
            Left            =   360
            TabIndex        =   79
            Top             =   4005
            Width           =   1515
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   285
            Index           =   40
            Left            =   1890
            Top             =   3960
            Width           =   2490
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name "
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   4680
            TabIndex        =   78
            Top             =   4005
            Width           =   885
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "A/C No."
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   5580
            TabIndex        =   77
            Top             =   4455
            Width           =   585
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PF Membership"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   76
            Top             =   4905
            Width           =   1095
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PF Mem. No."
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   4680
            TabIndex        =   75
            Top             =   4905
            Width           =   930
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   285
            Index           =   43
            Left            =   1890
            Top             =   4860
            Width           =   2490
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Branch Name "
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   360
            TabIndex        =   74
            Top             =   4455
            Width           =   1020
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Permanent  date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   4680
            TabIndex        =   67
            Top             =   1845
            Width           =   1170
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   420
            Index           =   13
            Left            =   6120
            Top             =   1755
            Width           =   2355
         End
         Begin VB.Label lblName 
            BackColor       =   &H00FFFFFF&
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   1
            Left            =   6075
            TabIndex        =   61
            Top             =   405
            Width           =   3705
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   345
            Index           =   55
            Left            =   6120
            Top             =   2250
            Width           =   2355
         End
         Begin VB.Label Label31 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Service Bk. No."
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   4680
            TabIndex        =   60
            Top             =   2325
            Width           =   1125
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   345
            Index           =   51
            Left            =   1890
            Top             =   2250
            Width           =   2490
         End
         Begin VB.Label Label31 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "File Ref. No."
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   59
            Top             =   2235
            Width           =   885
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Joining date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   58
            Top             =   1800
            Width           =   855
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
            Left            =   360
            TabIndex        =   57
            Top             =   900
            Width           =   840
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   4680
            TabIndex        =   56
            Top             =   450
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee ID"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   55
            Top             =   435
            Width           =   900
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Department"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4680
            TabIndex        =   54
            Top             =   870
            Width           =   825
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Job type"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   360
            TabIndex        =   53
            Top             =   1365
            Width           =   600
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   330
            Index           =   39
            Left            =   1890
            Top             =   360
            Width           =   2490
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   330
            Index           =   38
            Left            =   6030
            Top             =   360
            Width           =   3840
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   375
            Index           =   37
            Left            =   1890
            Top             =   810
            Width           =   2490
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   375
            Index           =   36
            Left            =   1890
            Top             =   1260
            Width           =   2490
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   375
            Index           =   33
            Left            =   6030
            Top             =   810
            Width           =   3045
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   375
            Index           =   29
            Left            =   1890
            Top             =   1755
            Width           =   2490
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   330
            Index           =   27
            Left            =   1890
            Top             =   2700
            Width           =   7980
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Responsibility"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   360
            TabIndex        =   52
            Top             =   2700
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   5850
         Index           =   0
         Left            =   -75000
         TabIndex        =   2
         Top             =   300
         Width           =   10275
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            Left            =   1890
            TabIndex        =   98
            Top             =   315
            Width           =   1230
         End
         Begin VB.CommandButton cmdView 
            Height          =   315
            Index           =   0
            Left            =   3150
            Picture         =   "frmEmp_Information.frx":8CEE
            Style           =   1  'Graphical
            TabIndex        =   97
            Top             =   300
            Width           =   375
         End
         Begin VB.TextBox txtEmail 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   6750
            TabIndex        =   93
            Top             =   5310
            Width           =   2970
         End
         Begin VB.TextBox txtPerm_Country 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   6750
            TabIndex        =   92
            Text            =   "Bangladesh"
            Top             =   4545
            Width           =   2970
         End
         Begin VB.TextBox txtPres_Country 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   1845
            TabIndex        =   91
            Text            =   "Bangladesh"
            Top             =   4545
            Width           =   2970
         End
         Begin VB.Frame Frame8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   465
            Left            =   3825
            TabIndex        =   20
            Top             =   1485
            Width           =   1770
            Begin VB.OptionButton optEmp_gender 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Male"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   1
               Left            =   180
               TabIndex        =   22
               Top             =   135
               Value           =   -1  'True
               Width           =   675
            End
            Begin VB.OptionButton optEmp_gender 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Female"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   0
               Left            =   900
               TabIndex        =   21
               Top             =   135
               Width           =   810
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H00FFC0C0&
               Height          =   345
               Index           =   26
               Left            =   90
               Top             =   90
               Width           =   1680
            End
         End
         Begin VB.ComboBox cboReligion 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "frmEmp_Information.frx":95B8
            Left            =   1800
            List            =   "frmEmp_Information.frx":95BA
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   2025
            Width           =   3060
         End
         Begin VB.ComboBox cboNational 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "frmEmp_Information.frx":95BC
            Left            =   1800
            List            =   "frmEmp_Information.frx":95C6
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   2385
            Width           =   3060
         End
         Begin VB.TextBox txtAddress_Pres 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   555
            Left            =   1845
            TabIndex        =   17
            Top             =   2760
            Width           =   2925
         End
         Begin VB.TextBox txtDist_Pres 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   1845
            TabIndex        =   16
            Top             =   3465
            Width           =   2925
         End
         Begin VB.TextBox txtPost_Pres 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   1845
            TabIndex        =   15
            Top             =   3825
            Width           =   2925
         End
         Begin VB.TextBox txtPS_Pres 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   1845
            TabIndex        =   14
            Top             =   4185
            Width           =   2925
         End
         Begin VB.TextBox txtTelephone_Pres 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   1845
            TabIndex        =   13
            Top             =   4905
            Width           =   2970
         End
         Begin VB.ComboBox cboMarital 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "frmEmp_Information.frx":95DF
            Left            =   6435
            List            =   "frmEmp_Information.frx":95EC
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2025
            Width           =   1350
         End
         Begin VB.TextBox txtAddress_Perm 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   555
            Left            =   6750
            TabIndex        =   11
            Top             =   2790
            Width           =   2925
         End
         Begin VB.TextBox txtDist_Perm 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   6750
            TabIndex        =   10
            Top             =   3465
            Width           =   2925
         End
         Begin VB.TextBox txtPost_Perm 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   6750
            TabIndex        =   9
            Top             =   3825
            Width           =   2925
         End
         Begin VB.TextBox txtPS_Perm 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   6750
            TabIndex        =   8
            Top             =   4185
            Width           =   2925
         End
         Begin VB.TextBox txtMobile 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   1845
            TabIndex        =   7
            Top             =   5310
            Width           =   2925
         End
         Begin VB.TextBox txtEmp_Name 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4230
            TabIndex        =   6
            Top             =   360
            Width           =   3510
         End
         Begin VB.TextBox txtFather 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1890
            TabIndex        =   5
            Top             =   765
            Width           =   5850
         End
         Begin VB.TextBox txtMother 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1890
            TabIndex        =   4
            Top             =   1170
            Width           =   5850
         End
         Begin VB.TextBox txtCode 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   6525
            TabIndex        =   3
            Top             =   1620
            Width           =   1170
         End
         Begin MSComCtl2.DTPicker dtpEmp_d_of_b 
            Height          =   330
            Left            =   1800
            TabIndex        =   23
            Top             =   1575
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   64356355
            CurrentDate     =   36948
         End
         Begin VB.Image imgPhoto 
            Height          =   1815
            Left            =   8100
            Top             =   420
            Width           =   1575
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00FFC0C0&
            Height          =   375
            Left            =   3150
            Top             =   270
            Width           =   420
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FFC0C0&
            Height          =   380
            Left            =   1845
            Top             =   270
            Width           =   1300
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Country"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   5355
            TabIndex        =   96
            Top             =   4500
            Width           =   540
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   5355
            TabIndex        =   95
            Top             =   5310
            Width           =   435
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   405
            TabIndex        =   94
            Top             =   4905
            Width           =   765
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Index           =   28
            Left            =   6705
            Top             =   5265
            Width           =   3030
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Index           =   25
            Left            =   1800
            Top             =   5265
            Width           =   3030
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   345
            Index           =   24
            Left            =   1800
            Top             =   4860
            Width           =   3030
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Birth"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   405
            TabIndex        =   43
            Top             =   1665
            Width           =   885
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   450
            TabIndex        =   42
            Top             =   5340
            Width           =   465
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Police Station"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   5355
            TabIndex        =   41
            Top             =   4185
            Width           =   975
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   3330
            TabIndex        =   40
            Top             =   1620
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marital status"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   5355
            TabIndex        =   39
            Top             =   2100
            Width           =   930
         End
         Begin VB.Label Label31 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Code No."
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   5670
            TabIndex        =   38
            Top             =   1650
            Width           =   675
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Permanent Address"
            ForeColor       =   &H00800000&
            Height          =   420
            Index           =   1
            Left            =   5355
            TabIndex        =   37
            Top             =   2790
            Width           =   885
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "District"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   5355
            TabIndex        =   36
            Top             =   3465
            Width           =   480
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Office"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   5355
            TabIndex        =   35
            Top             =   3810
            Width           =   780
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Police Station"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   405
            TabIndex        =   34
            Top             =   4140
            Width           =   975
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Country"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   405
            TabIndex        =   33
            Top             =   4530
            Width           =   540
         End
         Begin VB.Label Label33 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Father's/          Husband's Name"
            ForeColor       =   &H00800000&
            Height          =   420
            Left            =   405
            TabIndex        =   32
            Top             =   720
            Width           =   1305
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   3600
            TabIndex        =   31
            Top             =   360
            Width           =   420
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
            Left            =   405
            TabIndex        =   30
            Top             =   405
            Width           =   900
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Present Address"
            ForeColor       =   &H00800000&
            Height          =   420
            Index           =   0
            Left            =   405
            TabIndex        =   29
            Top             =   2745
            Width           =   660
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "District"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   405
            TabIndex        =   28
            Top             =   3420
            Width           =   480
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Office"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   405
            TabIndex        =   27
            Top             =   3765
            Width           =   780
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Religion"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   405
            TabIndex        =   26
            Top             =   2010
            Width           =   570
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Mother's Name"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   405
            TabIndex        =   25
            Top             =   1260
            Width           =   1305
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00F7B0AC&
            Height          =   2010
            Index           =   0
            Left            =   7920
            Top             =   315
            Width           =   1815
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nationality"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   405
            TabIndex        =   24
            Top             =   2340
            Width           =   735
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   615
            Index           =   3
            Left            =   1800
            Top             =   2745
            Width           =   3030
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Index           =   4
            Left            =   1800
            Top             =   3420
            Width           =   3030
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Index           =   5
            Left            =   1800
            Top             =   3780
            Width           =   3030
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Index           =   6
            Left            =   1800
            Top             =   4140
            Width           =   3030
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Index           =   7
            Left            =   1800
            Top             =   4500
            Width           =   3030
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   615
            Index           =   8
            Left            =   6705
            Top             =   2745
            Width           =   3030
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Index           =   9
            Left            =   6705
            Top             =   3420
            Width           =   3030
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Index           =   10
            Left            =   6705
            Top             =   3780
            Width           =   3030
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Index           =   11
            Left            =   6705
            Top             =   4140
            Width           =   3030
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Index           =   12
            Left            =   6705
            Top             =   4500
            Width           =   3030
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   345
            Index           =   1
            Left            =   4050
            Top             =   315
            Width           =   3750
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   345
            Index           =   16
            Left            =   1845
            Top             =   720
            Width           =   5955
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   345
            Index           =   17
            Left            =   1845
            Top             =   1125
            Width           =   5955
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   345
            Index           =   18
            Left            =   6480
            Top             =   1575
            Width           =   1320
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Per_Info As New emp_info
Private Job_Info As New clsEmp_Job_Detail
Dim SSTab_Index As Integer
Dim Default_Pic_Path As String
Dim New_Pic_Path As String
Dim Gender As Integer
Dim Payment_Mode As Integer
Dim PF_Membership As Integer
Dim TraceValidation As Boolean
Dim Get_EmpVlass As String

Private Sub cboDesig_Click()
  Get_Emp_Designation_Level (cboDesig)
End Sub

Private Sub cmdClear_Click()
On Error Resume Next
    Clear_Screen
    'txtEmp_ID(SSTab_Index).SetFocus
    Combo1(0).SetFocus
End Sub
Private Sub cmdClose_Click()
    Close_Msg Me
End Sub





Private Sub cmdPrint_Click()
 rptmode = 4
 Form20.Show vbModal
    
End Sub

Private Sub cmdSave_Click()
On Error GoTo Errdes
Select Case SSTab_Index

Case 0

    If optEmp_gender(1) Then Gender = 1
    If optEmp_gender(0) Then Gender = 0


With Per_Info    ''' Employee Personal Information

    .Connstring = strCN.Connection_String
    '.EMP_ID = txtEmp_ID(0)
    .Emp_ID = Combo1(0).Text
    .Emp_Nm = txtEmp_Name
    .Emp_Fat_Nm = txtFather
    .Emp_Mat_Nm = txtMother
    .Code_No = txtCode
    .DOB = dtpEmp_d_of_b
    
    .Religion = cboReligion.ListIndex
    .Nationality = cboNational.ListIndex
    .Marital_Stat = cboMarital.ListIndex
    
    .Address_Perm = txtAddress_Perm
    .Address_Pres = txtAddress_Pres
    
    .District_Perm = txtDist_Perm
    .District_Pres = txtDist_Pres
    
    .PS_Perm = txtPS_Perm
    .PS_Pres = txtPS_Pres
    
    .Post_Perm = txtPost_Perm
    .Post_Pres = txtPost_Pres
    
    .Country_Perm = txtPerm_Country
    .Country_Pres = txtPres_Country
    
    .TELEPHONE = txtTelephone_Pres
    .E_mail = txtEmail
    .Cellphone = txtMobile
    
    .Gender = Gender
    

    .Photo = New_Pic_Path
    .Save
    .Show_Message
    
End With

Show_Personal_Info

Case 1     ''' Job Information

   
    
    If optPayment_Mode(0) Then Payment_Mode = 0
    If optPayment_Mode(1) Then Payment_Mode = 1
    
    If optPF(0) Then PF_Membership = 0
    If optPF(1) Then PF_Membership = 1
    If Len(cboScale) = 0 Then
       MsgBox "Pay Scale Required", vbInformation, "Required....."
       cboScale.SetFocus
       Exit Sub
    End If
    
     If Len(Combo2) = 0 Then
       MsgBox "Employee Class Required", vbInformation, "Required...."
       Combo2.SetFocus
       Exit Sub
    End If
    
    If Len(Combo2) = 0 Then
       MsgBox "Employee Class Required", vbInformation, "Required...."
       Combo2.SetFocus
       Exit Sub
    End If
    
    If Len(lblName(1).Caption) = 0 Then
       MsgBox "Employee Name Required", vbInformation, "Required...."
       txtEmp_ID(1).SetFocus
       Exit Sub
    End If
'''''''''''''''''''''''Employee_Salary_Validation---------------------- for the remove the data entry
If TraceValidation = True Then
    Exit Sub
End If
With Job_Info
        
    .Connstring = strCN.Connection_String
    .Emp_ID = txtEmp_ID(1)
    
    .designation = cboDesig
    .Dept = cboDept
    .JobType = cboType
    .Jdate = dtpEmp_join_date
    .Reseponsibility = txtResponsibility
    .File_ref_number = txtFile_Ref_No
    .Service_bk_number = txtService_Bk_No
    .Scale_code = cboScale
    .Basic_Sal = IIf((Len(Trim(txtBasic)) = 0), 0, txtBasic)
    .EmpPositionSerial = txtStaffSerial
    .EmpDesignationLevel = txtDesignationLevel
'    If .Pdate = "00:00:00" Then
'        '.Pdate = "__/__/__"
'        .Pdate = Null
'    ElseIf .Pdate = "12:00:00 AM" Then
'            .Pdate = Format(dtpParmanent_Dt.Text, "dd/mm/yy")
'    Else
'        .Pdate = IIf((.Pdate = "__/__/__"), Null, Format(dtpParmanent_Dt.Text, "dd/mm/yy"))
'    End If
    
    .Pdate = dtpParmanent_Dt
    
    
    .Pf_mem = PF_Membership
    .Pf_mem_no = txtPF_Mem_No
    If txtPrevious_Balance = "" Then
       .Pre_Balance = 0
    Else
      .Pre_Balance = txtPrevious_Balance
    End If
    .Mode_of_payment = Payment_Mode
    .bank_name = cboBank
    .Branch_name = cboBankBranch
    .Acc_No = Trim(txtAccountNo)
    .BMDCREGI = Text1.Text
    .EmpClass = GetClassInInt
    
    
    .Save
    .Show_Message
End With

Show_Job_Info

End Select

Exit Sub
Errdes:

If Err.Number = 13 Then

    MsgBox "Permanent Date is not Avialable", vbInformation, "IT Division, DNMIH"
Else

    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End If
End Sub
Private Function GetClassInInt() As Integer
   If Combo2.Text = Combo2.List(0) Then
         GetClassInInt = 1
    ElseIf Combo2.Text = Combo2.List(1) Then
        GetClassInInt = 2
    ElseIf Combo2.Text = Combo2.List(2) Then
        GetClassInInt = 3
    Else
        GetClassInInt = 4
    End If
    
End Function
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer, Index As Integer)
    Select Case SSTab_Index
        
    Case 0
        
        Show_Personal_Info
    Case 1
        Show_Job_Info
        
    End Select


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Load()
On Error GoTo Errdes
    Screen_Position Me
    SSTab_Index = 0
    Set_TabIndex
    Load_Religion Me
    cboNational.ListIndex = 0
    cboMarital.ListIndex = 1
            'txtEmp_ID(0).MaxLength = Id_Len
    Default_Pic_Path = App.Path + "\Default_Pic.bmp"
    Get_Emp_ID_Into_Combo
    
    'select emp_id from emp_info;
    
    Dim cmd As New Command
    Dim conn1 As New Connection
    Dim rs1 As New Recordset
    
    conn1.ConnectionString = strCN.Connection_String
    conn1.Open
    cmd.ActiveConnection = conn1
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "select emp_id from emp_info order by emp_id"
    rs1.CursorLocation = adUseClient
    rs1.Open cmd.CommandText, conn1, adOpenDynamic, adLockOptimistic
    
    If rs1.RecordCount > 0 Then
     
        Do Until rs1.EOF
            Combo1(0).AddItem rs1.Fields(0)
            rs1.MoveNext
        Loop
        
     
    End If
    
    rs1.Close
    conn1.Close
 SSTab1.Tab = 0
 AUTHORIZATION (UserRole)
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub AUTHORIZATION(UserRole As String)
    Select Case UserRole
           Case "Accounts"
                Combo1(0).Enabled = False
                cboDesig.Enabled = False
                cboDept.Enabled = False
                cboType.Enabled = False
                Combo2.Enabled = False
                dtpEmp_join_date.Enabled = False
                dtpEmp_d_of_b.Enabled = False
                txtFile_Ref_No.Enabled = False
                dtpParmanent_Dt.Enabled = False
                MaxSerialButton.Enabled = False
                txtStaffSerial.Enabled = False
           Case "Personnel"
                cboBank.Enabled = False
                cboBankBranch.Enabled = False
                txtAccountNo.Enabled = False
                optPF.Item(0).Enabled = False
                optPF.Item(1).Enabled = False
                optPayment_Mode.Item(0).Enabled = False
                optPayment_Mode.Item(0).Enabled = False
    End Select
End Sub
Private Sub MaxSerialButton_Click()
  Dim i As Integer
  i = GetClassInInt
  Get_Max_Emp_Position_Serial (i)
End Sub

Private Sub optEmp_gender_Click(Index As Integer)
    Gender = Index
End Sub
Private Sub optPayment_Mode_Click(Index As Integer)
    Payment_Mode = Index
End Sub
Private Sub optPF_Click(Index As Integer)
    PF_Membership = Index
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo Errdes

SSTab_Index = SSTab1.Tab

Set_TabIndex    '' Rearrange tab index

Select Case SSTab_Index

Case 0
    
    Load_Religion Me
    cboNational.ListIndex = 0
    cboMarital.ListIndex = 1
    
    
    If Combo1(0) <> "" Then
         Combo1(0) = txtEmp_ID(1)
        Show_Personal_Info
    End If
    
Case 1
    
    Load_Desig Me
    Load_JbType Me
    Load_Department Me
    Load_PScale Me
    Load_BankNm Me
    Load_Bank_Branch_Nm Me
    
    If Trim(Combo1(0).Text) <> "" Then
        txtEmp_ID(1) = Combo1(0).Text
        Show_Job_Info
    End If
    
End Select
   
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Public Sub Set_TabIndex()
On Error GoTo Errdes

Select Case SSTab_Index

Case 0  ''Employee Personal Information

                'txtEmp_ID(0).TabIndex = 0
    Combo1(0).TabIndex = 0
    txtEmp_Name.TabIndex = 1
    txtFather.TabIndex = 2
    txtMother.TabIndex = 3
    dtpEmp_d_of_b.TabIndex = 4
    optEmp_gender(0).TabIndex = 5
    optEmp_gender(1).TabIndex = 6
    
    cboReligion.TabIndex = 7
    cboMarital.TabIndex = 8
    cboNational.TabIndex = 9
    
    txtAddress_Pres.TabIndex = 10
    txtDist_Pres.TabIndex = 11
    txtPost_Pres.TabIndex = 12
    txtPS_Pres.TabIndex = 13
    txtPres_Country.TabIndex = 14
    txtTelephone_Pres.TabIndex = 15
    
    
    txtAddress_Perm.TabIndex = 16
    txtDist_Perm.TabIndex = 17
    txtPost_Perm.TabIndex = 18
    txtPS_Perm.TabIndex = 19
    txtPerm_Country.TabIndex = 20
    
    txtMobile.TabIndex = 21
    txtEmail.TabIndex = 22
    
    cmdSave.TabIndex = 23
    cmdClear.TabIndex = 24
    cmdDelete.TabIndex = 25
    cmdClose.TabIndex = 26


Case 1  '' Employee Job Information

    txtEmp_ID(1).TabIndex = 0
    cboDesig.TabIndex = 1
    cboDept.TabIndex = 2
    cboType.TabIndex = 3
    dtpEmp_join_date.TabIndex = 4
    dtpParmanent_Dt.TabIndex = 5
    
    txtFile_Ref_No.TabIndex = 6
    txtService_Bk_No.TabIndex = 7
    
    txtResponsibility.TabIndex = 8
    
    cboScale.TabIndex = 9
    txtBasic.TabIndex = 10
    optPayment_Mode(0).TabIndex = 11
    optPayment_Mode(1).TabIndex = 12
    cboBank.TabIndex = 13
    cboBankBranch.TabIndex = 14
    txtAccountNo.TabIndex = 15
    
    optPF(1).TabIndex = 16
    optPF(0).TabIndex = 17
    
    txtPF_Mem_No.TabIndex = 18
    txtPrevious_Balance.TabIndex = 19
    
    cmdSave.TabIndex = 20
    cmdClear.TabIndex = 21
    cmdDelete.TabIndex = 22
    cmdClose.TabIndex = 23


End Select

    'txtEmp_ID(SSTab_Index).SetFocus
    'txtEmp_ID(SSTab_Index).SelStart = Len(txtEmp_ID(SSTab_Index))

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Destroy Me
End Sub
Public Sub Load_Photo(ComDiag As CommonDialog, Img As Image, Optional Photo_Path As String)
On Error GoTo Errdes
Dim resp As String

Start:  With ComDiag
            .Filter = "Photograph,*.bmp;*.jpg;*.gif|*.bmp;*.jpg;*.gif"
            .Action = 1
                If .FileName = "" Then
                    Exit Sub
                Else
                    New_Pic_Path = .FileName
                    Img.Picture = LoadPicture(New_Pic_Path)
                End If
        End With



        resp = MsgBox("           Is it the right Picture ?", vbYesNoCancel + vbQuestion + vbDefaultButton2, "Message")
        
        
        If resp = vbCancel Then
            Img.Picture = LoadPicture(Default_Pic_Path)
            Exit Sub
        End If
        
        If resp = vbNo Then
            Img.Picture = LoadPicture(Default_Pic_Path)
            GoTo Start
            Exit Sub
        End If
        
        If resp = vbYes Then
                New_Pic_Path = ComDiag.FileName
            Exit Sub
        
        End If
        
Exit Sub
Errdes:
    MsgBox Err.desctiption, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub imgPhoto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Call Load_Photo(ComDiag, imgPhoto, Default_Pic_Path)
End If
End Sub
Public Sub Show_Personal_Info()
On Error GoTo Errdes

With Per_Info
    .Connstring = strCN.Connection_String
    .Emp_ID = Combo1(0).Text
    .GetX
    
    Clear_Screen
    
    txtEmp_Name = .Emp_Nm
    txtFather = .Emp_Fat_Nm
    txtMother = .Emp_Mat_Nm


'    If Per_Info.Photo = True Or Per_Info_Rs.BOF = True Then
'         Clear_Screen
'         imgPhoto.Picture = LoadPicture(App.Path + "\Default_Pic.bmp")
'         Exit Sub
'    End If


    txtCode = .Code_No
    dtpEmp_d_of_b = .DOB
    
    txtAddress_Perm = .Address_Pres
    'txtAddress_Perm =
    txtAddress_Pres = .Address_Perm
    
    
    
    
    txtPerm_Country = .Country_Perm
    txtPres_Country = .Country_Pres
    
    txtPS_Perm = .PS_Perm
    txtPS_Pres = .PS_Pres
    
    txtTelephone_Pres = .TELEPHONE
    txtMobile = .Cellphone
    
    txtDist_Perm = .District_Perm
    txtDist_Pres = .District_Pres
    
    txtPost_Perm = .Post_Perm
    txtPost_Pres = .Post_Pres
    
'    cboMarital.ListIndex = CInt(.Marital_Stat)
'    cboReligion.ListIndex = CInt(.Religion)
    cboMarital.ListIndex = .Marital_Stat
    cboReligion.ListIndex = .Religion
    

    optEmp_gender(.Gender).Value = True
    
    txtEmail = .E_mail
    
End With


        'If Not !Photo = Empty Then
'            imgPhoto.Picture = LoadPicture(!Photo)
'            Default_Pic_Path = !Photo
'            New_Pic_Path = imgPhoto
'        Else
'            imgPhoto.Picture = LoadPicture(Default_Pic_Path)
        'End If

'    End With
'
Set Per_Info = Nothing

Exit Sub
Errdes:
If Err.Number = 13 Then
   Clear_Screen
Else
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End If
End Sub
Public Sub Show_Job_Info()
On Error GoTo Errdes
With Job_Info

    .Connstring = strCN.Connection_String
    .Emp_ID = Trim(txtEmp_ID(1))
    .GetX
    
    Clear_Screen
    
    txtEmp_ID(1) = .Emp_ID
    lblName(1) = .Emp_Nm
    
    dtpEmp_join_date = .Jdate
    
    'MsgBox .Pdate
    
'    If .Pdate = "00:00:00" Then
'        dtpParmanent_Dt = "__/__/__"
'    Else
'        dtpParmanent_Dt = Format(.Pdate, "dd/mm/yy")
'    End If
    
    'dtpParmanent_Dt = IIf(.Pdate, "00:00:00", Null, Format(.Pdate, "dd/mm/yy"))
    
    dtpParmanent_Dt = .Pdate
    
    cboDept = .Dept
    cboDesig = .designation
    cboType = .JobType
    optPayment_Mode(CInt(.Mode_of_payment)).Value = True
    optPF(CInt(.Pf_mem)).Value = True
    txtBasic = .Basic_Sal
    cboBank = .bank_name
    cboScale = .Scale_code
    cboBankBranch = .Branch_name
    txtFile_Ref_No = .File_ref_number
    txtService_Bk_No = .Service_bk_number
    txtResponsibility = .Reseponsibility
    txtPF_Mem_No = .Pf_mem_no
    txtPrevious_Balance = .Pre_Balance
    txtAccountNo = .Acc_No
    Text1 = .BMDCREGI
    txtStaffSerial = .EmpPositionSerial
    Get_Emp_Class
   
'    If .EmpClass = "1" Then
'        Combo2.Text = Combo2.List(0)
'    ElseIf .EmpClass = "2" Then
'        Combo2.Text = Combo2.List(1)
'    ElseIf .EmpClass = "3" Then
'        Combo2.Text = Combo2.List(2)
'    Else
'        Combo2.Text = Combo2.List(2)
'    End If
    
End With
Set Job_Info = Nothing
Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub txtEmp_ID_Change(Index As Integer)
  Select Case Index
        Case 1
              lblName(1).Caption = ""
  End Select

End Sub

Private Sub txtEmp_ID_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If txtEmp_ID(Index) <> "" And KeyCode = 13 Then

    Select Case SSTab_Index
        
    Case 0
        
        Show_Personal_Info
    Case 1
        Show_Job_Info
    
    End Select
        
    txtEmp_ID(SSTab_Index).SetFocus
    txtEmp_ID(SSTab_Index).SelStart = Len(txtEmp_ID(SSTab_Index))

End If
End Sub
Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    KeyAscii = IsNum(KeyAscii)
End Sub
Private Sub txtEmp_ID1_Change(Index As Integer)
End Sub
Private Sub Get_Emp_ID_Into_Combo()
On Error GoTo Errdes
Dim cmd As New Command
Dim conn4 As New Connection
Dim rs4 As New Recordset

conn4.ConnectionString = strCN.Connection_String
conn4.Open
cmd.ActiveConnection = conn4
cmd.CommandType = adCmdText

cmd.CommandText = "select Emp_Id from Emp_Info order by emp_id"
rs4.CursorLocation = adUseClient
rs4.Open cmd.CommandText, conn4, adOpenDynamic, adLockOptimistic

    If rs4.RecordCount > 0 Then
     
        Do Until rs4.EOF
            Combo1(0).AddItem rs4.Fields(0)
            rs4.MoveNext
        Loop
        
     
    End If
    
    rs4.Close
    conn4.Close

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Employee_Salary_Validation()
On Error GoTo Errdes
Dim cmd As New Command
Dim conn4 As New Connection
Dim rs4 As New Recordset

conn4.ConnectionString = strCN.Connection_String
conn4.Open
cmd.ActiveConnection = conn4
cmd.CommandType = adCmdText

cmd.CommandText = " SELECT STR_BASIC,EB_END FROM ST_PAYSCALE WHERE SCALE_CODE='" & cboScale & "' " + _
                " AND YR_REF=(SELECT MAX(YR_REF) FROM ST_PAYSCALE)"
rs4.CursorLocation = adUseClient
rs4.Open cmd.CommandText, conn4, adOpenDynamic, adLockOptimistic

    If rs4.RecordCount > 0 Then
            
            If Val(txtBasic) < Val(rs4.Fields(0)) Then
                MsgBox "Invalid Basic of the Employee", vbInformation, "IT Division, DNMIH"
                txtBasic.SetFocus
                TraceValidation = True
            ElseIf Val(txtBasic) > Val(rs4.Fields(1)) Then
                MsgBox "Invalid Basic of the Employee (OverPayment)", vbInformation, "IT Division, DNMIH"
                txtBasic.SetFocus
                TraceValidation = True
             Else
                TraceValidation = False
             End If
                
    End If
            
'            If Val(txtBasic) > Val(rs4.Fields(1)) Then
'                MsgBox "Invalid Basic of the Employee (OverPayment)", vbInformation, "IT Division, DNMIH"
'                txtBasic.SetFocus
'                TraceValidation = True
'             End If
            
    'End If
    
    
    rs4.Close
    conn4.Close

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Get_Emp_Class()
On Error GoTo Errdes
Dim cmd As New Command
Dim conn4 As New Connection
Dim rs4 As New Recordset

conn4.ConnectionString = strCN.Connection_String
conn4.Open
cmd.ActiveConnection = conn4
cmd.CommandType = adCmdText

cmd.CommandText = "select emp_class from EMP_JOB_INFO where Emp_Id='" & txtEmp_ID(1) & "'"
rs4.CursorLocation = adUseClient
rs4.Open cmd.CommandText, conn4, adOpenDynamic, adLockOptimistic

    If rs4.RecordCount > 0 Then
       
        If rs4.Fields(0) = "1" Then
           Combo2.Text = Combo2.List(0)
        ElseIf rs4.Fields(0) = "2" Then
           Combo2.Text = Combo2.List(1)
        ElseIf rs4.Fields(0) = "3" Then
           Combo2.Text = Combo2.List(2)
        Else
           Combo2.Text = Combo2.List(3)
        End If
     Else
        Combo2.Text = Combo2.List(0)
    End If
    
    rs4.Close
    conn4.Close

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Get_Max_Emp_Position_Serial(Class As Integer)
On Error GoTo Errdes
Dim cmd As New Command
Dim conn4 As New Connection
Dim rs4 As New Recordset

conn4.ConnectionString = strCN.Connection_String
conn4.Open
cmd.ActiveConnection = conn4
cmd.CommandType = adCmdText

cmd.CommandText = "select nvl(max(emp_position),0)+1 as position from EMP_JOB_INFO where emp_class=" & Class & ""
rs4.CursorLocation = adUseClient
rs4.Open cmd.CommandText, conn4, adOpenDynamic, adLockOptimistic

    If rs4.RecordCount > 0 Then
       txtStaffSerial = rs4.Fields(0)
        
    End If
    
    rs4.Close
    conn4.Close

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Get_Emp_Designation_Level(designation As String)
On Error GoTo Errdes
Dim cmd As New Command
Dim conn4 As New Connection
Dim rs4 As New Recordset

conn4.ConnectionString = strCN.Connection_String
conn4.Open
cmd.ActiveConnection = conn4
cmd.CommandType = adCmdText

cmd.CommandText = "select desig_Level as designation_Level from ST_Desig where Designation= '" & designation & "'"
rs4.CursorLocation = adUseClient
rs4.Open cmd.CommandText, conn4, adOpenDynamic, adLockOptimistic

    If rs4.RecordCount > 0 Then
       txtDesignationLevel = rs4.Fields(0)
        
    End If
    
    rs4.Close
    conn4.Close

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub txtStaffSerial_Change()
  If Not IsNumeric(txtStaffSerial) Then
      txtStaffSerial = ""
  End If
  
End Sub
