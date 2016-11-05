VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Salary Preparation"
   ClientHeight    =   8475
   ClientLeft      =   660
   ClientTop       =   1560
   ClientWidth     =   10515
   Icon            =   "frmSalary_Preparation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9255
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   7410
      Width           =   1185
   End
   Begin VB.CommandButton cmdNewBasic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Re-Calculate(Ctrl+R)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7140
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "Press to Re-calculate(Clrt+R)"
      Top             =   5670
      Width           =   1935
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   62
      Top             =   8055
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   194028
            MinWidth        =   194028
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstTips 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3FEFF&
      ForeColor       =   &H000000C0&
      Height          =   225
      ItemData        =   "frmSalary_Preparation.frx":08CA
      Left            =   4980
      List            =   "frmSalary_Preparation.frx":08CC
      TabIndex        =   28
      Top             =   6930
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5445
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7410
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Prepare(Ctrl+P)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Press to Update"
      Top             =   7410
      Width           =   1425
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7410
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7410
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   7335
      Index           =   0
      Left            =   -30
      TabIndex        =   9
      Top             =   -45
      Width           =   10455
      Begin VB.TextBox txtDesignationLevel 
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   6990
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtStaffSerial 
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5370
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   1410
         Width           =   645
      End
      Begin VB.TextBox txtEducationAssisAllow 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1755
         TabIndex        =   70
         TabStop         =   0   'False
         ToolTipText     =   "Education Assistance Allowance"
         Top             =   4950
         Width           =   915
      End
      Begin VB.TextBox txtDressAllowance 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   1785
         TabIndex        =   80
         Text            =   "0"
         Top             =   5325
         Width           =   915
      End
      Begin VB.ComboBox CboSalaryType 
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
         Height          =   315
         ItemData        =   "frmSalary_Preparation.frx":08CE
         Left            =   6420
         List            =   "frmSalary_Preparation.frx":08DE
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   360
         Width           =   1500
      End
      Begin VB.CommandButton cmdShowInformation 
         Caption         =   "Show Information"
         Height          =   345
         Left            =   7920
         TabIndex        =   74
         Top             =   330
         Width           =   1395
      End
      Begin VB.TextBox txtSDA 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1755
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   4590
         Width           =   915
      End
      Begin VB.TextBox txtWorkingDay 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   7215
         TabIndex        =   68
         Text            =   "0"
         Top             =   1470
         Width           =   375
      End
      Begin VB.TextBox lblRev_Stamp 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   4635
         TabIndex        =   65
         Top             =   3630
         Width           =   915
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1745
         TabIndex        =   0
         Top             =   360
         Width           =   1320
      End
      Begin VB.TextBox txtOthersDeduction 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   4635
         TabIndex        =   46
         Top             =   3975
         Width           =   915
      End
      Begin VB.CommandButton cmdView 
         Height          =   330
         Index           =   0
         Left            =   3095
         Picture         =   "frmSalary_Preparation.frx":0911
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   340
         Width           =   375
      End
      Begin VB.TextBox txtLeave 
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
         Height          =   240
         Left            =   9780
         TabIndex        =   48
         Top             =   1455
         Width           =   465
      End
      Begin VB.TextBox txtDept 
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   7710
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   975
         Width           =   2475
      End
      Begin VB.TextBox txtDesig 
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1440
         Width           =   3120
      End
      Begin VB.TextBox txtName 
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1005
         Width           =   5685
      End
      Begin VB.TextBox txtNet_Payable 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   240
         Left            =   7290
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   6510
         Width           =   915
      End
      Begin VB.TextBox txtTotal_Deduction 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   240
         Left            =   4635
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   6540
         Width           =   915
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   240
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   6540
         Width           =   915
      End
      Begin VB.TextBox txtBonus 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   1800
         TabIndex        =   41
         Top             =   5700
         Width           =   915
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   1740
         Left            =   7275
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   49
         Top             =   2400
         Width           =   2940
      End
      Begin VB.TextBox txtAttn 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   8715
         TabIndex        =   47
         Text            =   "0"
         Top             =   1455
         Width           =   345
      End
      Begin VB.TextBox txtAdvance 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   4635
         ScrollBars      =   1  'Horizontal
         TabIndex        =   44
         Top             =   3000
         Width           =   915
      End
      Begin VB.TextBox txtPF_Loan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   4635
         TabIndex        =   43
         Top             =   2685
         Width           =   915
      End
      Begin VB.TextBox txtPF_Contribution 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4635
         TabIndex        =   4
         Top             =   2370
         Width           =   915
      End
      Begin VB.TextBox txtNationalDisesterFund 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   4635
         TabIndex        =   45
         Top             =   3315
         Width           =   915
      End
      Begin VB.TextBox txtMed 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1755
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3030
         Width           =   915
      End
      Begin VB.TextBox txtOthers_Add 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   1725
         TabIndex        =   42
         Top             =   6060
         Width           =   975
      End
      Begin VB.TextBox txtArrear 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   1755
         TabIndex        =   40
         Top             =   3960
         Width           =   915
      End
      Begin VB.TextBox txtDA 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1755
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   4275
         Width           =   915
      End
      Begin VB.TextBox txtTiffin 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1755
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3660
         Width           =   915
      End
      Begin VB.TextBox txtConv 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1755
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3345
         Width           =   915
      End
      Begin VB.TextBox txtHR 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1755
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2715
         Width           =   915
      End
      Begin VB.TextBox txtBasic 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2430
         Width           =   915
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
         Left            =   1800
         TabIndex        =   1
         Top             =   405
         Width           =   1185
      End
      Begin VB.ComboBox cboMonth 
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
         Height          =   315
         Left            =   4980
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1440
      End
      Begin VB.ComboBox cboYear 
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
         Height          =   315
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Srl"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5070
         TabIndex        =   85
         Top             =   1470
         Width           =   180
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   26
         Left            =   1710
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&E.A.Allowance *"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   390
         TabIndex        =   83
         Top             =   4980
         Width           =   1140
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Dress Allowance"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   390
         TabIndex        =   82
         Top             =   5310
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   25
         Left            =   1710
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Salary  Type"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   6450
         TabIndex        =   79
         Top             =   120
         Width           =   885
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "( Ctrl+N )"
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
         Height          =   195
         Left            =   2220
         TabIndex        =   77
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "(Ctrl+J)"
         ForeColor       =   &H00008080&
         Height          =   225
         Left            =   9900
         TabIndex        =   76
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Job Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   8790
         TabIndex        =   75
         Top             =   90
         Width           =   1695
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00C0C0FF&
         Height          =   1785
         Left            =   -60
         Top             =   90
         Width           =   11685
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0C0FF&
         Height          =   375
         Left            =   1710
         Top             =   330
         Width           =   5655
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   13
         Left            =   1710
         Top             =   5655
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000004&
         BorderWidth     =   2
         X1              =   360
         X2              =   2835
         Y1              =   6420
         Y2              =   6435
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&N.A.*"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   390
         TabIndex        =   71
         Top             =   4620
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   24
         Left            =   1710
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   23
         Left            =   4590
         Top             =   3930
         Width           =   1095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Other(-)"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3240
         TabIndex        =   67
         Top             =   3990
         Width           =   525
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "(No Advance)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   5760
         TabIndex        =   66
         Top             =   3060
         Width           =   1305
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   3060
         Top             =   330
         Width           =   465
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   22
         Left            =   7110
         Top             =   1410
         Width           =   555
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Working Day"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6075
         TabIndex        =   64
         Top             =   1455
         Width           =   930
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Leave"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9225
         TabIndex        =   61
         Top             =   1470
         Width           =   450
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Left            =   9735
         Top             =   1410
         Width           =   555
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Bonus (Fesival)"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   375
         TabIndex        =   55
         Top             =   5685
         Width           =   1080
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   465
         TabIndex        =   53
         Top             =   1380
         Width           =   840
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   8430
         TabIndex        =   52
         Top             =   705
         Width           =   825
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   21
         Left            =   1710
         Top             =   1410
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   12
         Left            =   7590
         Top             =   960
         Width           =   2700
      End
      Begin VB.Label lblRev_Stamp2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5955
         TabIndex        =   5
         Top             =   3915
         Width           =   915
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   20
         Left            =   1710
         Top             =   960
         Width           =   5865
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   19
         Left            =   1710
         Top             =   6495
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   18
         Left            =   4590
         Top             =   6495
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   17
         Left            =   4590
         Top             =   3615
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   16
         Left            =   7245
         Top             =   6465
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   1920
         Index           =   15
         Left            =   7110
         Top             =   2355
         Width           =   3150
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   14
         Left            =   8610
         Top             =   1410
         Width           =   555
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   11
         Left            =   4590
         Top             =   2985
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   10
         Left            =   4590
         Top             =   2670
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   9
         Left            =   4590
         Top             =   2355
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   8
         Left            =   4590
         Top             =   3300
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   7
         Left            =   1710
         Top             =   2985
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   6
         Left            =   1710
         Top             =   6030
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   5
         Left            =   1710
         Top             =   4245
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   4
         Left            =   1710
         Top             =   3930
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   3
         Left            =   1710
         Top             =   3615
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   2
         Left            =   1710
         Top             =   3300
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   1
         Left            =   1710
         Top             =   2670
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   0
         Left            =   1710
         Top             =   2355
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   370
         Index           =   39
         Left            =   1710
         Top             =   325
         Width           =   1365
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Name"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   465
         TabIndex        =   34
         Top             =   1005
         Width           =   420
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Remarks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   240
         Left            =   7110
         TabIndex        =   33
         Top             =   1950
         Width           =   1500
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         Caption         =   "Employee  &ID"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   435
         TabIndex        =   32
         Top             =   375
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N.D Fund (-)"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3240
         TabIndex        =   31
         Top             =   3345
         Width           =   1005
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Arrear"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   405
         TabIndex        =   30
         Top             =   3960
         Width           =   825
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DA"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   405
         TabIndex        =   29
         Top             =   4320
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Payable =(Salary + Allowances) - (Total Deduction)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009BA76D&
         Height          =   225
         Index           =   7
         Left            =   4650
         TabIndex        =   27
         Top             =   5115
         Width           =   4605
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Payable"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   33
         Left            =   6210
         TabIndex        =   26
         Top             =   6525
         Width           =   870
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Payable"
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   32
         Left            =   345
         TabIndex        =   25
         Top             =   6570
         Width           =   1215
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deduction"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   31
         Left            =   3240
         TabIndex        =   24
         Top             =   6555
         Width           =   1140
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Revenue Stamp"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   29
         Left            =   3240
         TabIndex        =   23
         Top             =   3645
         Width           =   1155
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Atten&dance"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   7740
         TabIndex        =   22
         Top             =   1455
         Width           =   825
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Advance"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   26
         Left            =   3240
         TabIndex        =   21
         Top             =   3030
         Width           =   1125
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Others (+)"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   21
         Left            =   375
         TabIndex        =   20
         Top             =   6030
         Width           =   690
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Loan && Advances"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Index           =   1
         Left            =   3270
         TabIndex        =   19
         Top             =   1950
         Width           =   1665
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary && Allowances"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Index           =   0
         Left            =   420
         TabIndex        =   18
         Top             =   1950
         Width           =   1980
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PF Loan"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   3240
         TabIndex        =   17
         Top             =   2715
         Width           =   600
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PF Contribution"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   16
         Top             =   2370
         Width           =   1080
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tiffin"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   3
         Left            =   405
         TabIndex        =   15
         Top             =   3660
         Width           =   825
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Conveyance"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   405
         TabIndex        =   14
         Top             =   3345
         Width           =   960
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Medical Allown."
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   405
         TabIndex        =   13
         Top             =   3045
         Width           =   1140
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Salary"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   405
         TabIndex        =   12
         Top             =   2400
         Width           =   870
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "House Rent"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   405
         TabIndex        =   11
         Top             =   2730
         Width           =   885
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year                        Month"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   3840
         TabIndex        =   10
         Top             =   150
         Width           =   1860
      End
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "E.A=Education Assist.Allowance"
      Height          =   285
      Left            =   270
      TabIndex        =   84
      Top             =   7650
      Width           =   2355
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "N.A=National Allowance"
      Height          =   285
      Left            =   270
      TabIndex        =   81
      Top             =   7350
      Width           =   1815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Advance_Info As New clsAdvance_Info
Private objSalary As New Salary_Monthly
Private Salaray_Pre As New Cls_salary_Preparation
Private Utility As New clsUtility

Dim conn1 As New Connection
Dim rs1 As New Recordset
Dim conn2 As New Connection
Dim RS2 As New Recordset
Dim EmpContr, EmployeerContr, BasicSalaryOfEmp, EmpScaleID, Employee_Contribution, Employeer_Contribution, total_Contribution
Dim conn3 As New Connection
Dim rs3 As New Recordset
Dim conn4 As New Connection
Dim rs4 As New Recordset
Dim conn5 As New Connection
Dim rs5 As New Recordset
Dim HouseRentValue, MedicalValue, ConvenceCalue, TiffinValue
Dim conn6 As New Connection
Dim ChechWhetherSlaryPaidirNot
Dim AmountofLoanTaken, SlabAmountofLoan, NoofInstallmentPaidByEmp, TotalLoanPaidByEmp, TotalInstallment
Dim AdvanceTakenAmount, NoofIstalltimeAdvancePaid, PaidAdvanceInstallment, AdvanceBalanceAmount, DedeductionAdvanceEachMonth
Dim NoofIstalltimeAdvance, PFALLOWORNOT, EmployeeeHousrRent, ClassofEmployee
Dim conn007 As New Connection
Dim RS007 As New Recordset
Dim rs005 As New Recordset
Dim EmpIdExitinMain
Dim AdvanceTakenByEmp As Boolean
Dim LoanTakenFlag As Boolean
Dim srlType As String
Dim isSDAApplicable As Boolean

Private Sub cboMonth_Click()
   ClearAddition
   ClearDeduction
End Sub

Private Sub cboMonth_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    CboSalaryType.SetFocus
  End If
End Sub

Private Sub CboSalaryType_Click()
  Select Case CboSalaryType.Text
        Case "Reg."
           localSalaryType = "R"
           Form3.Caption = "Salary Preparation"
        Case "Supp."
           localSalaryType = "S"
           Form3.Caption = "Supplementary Salary Preparation"
        Case "Bonus Only"
           localSalaryType = "B"
           Form3.Caption = "Only Bonus Preparation"
        Case "Dress Allowance Only"
           localSalaryType = "D"
           Form3.Caption = "Dress Allowance Preparation"
  End Select
  
End Sub

Private Sub CboSalaryType_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     cmdShowInformation.SetFocus
  End If
End Sub

Private Sub cboYear_Click()
   ClearAddition
   ClearDeduction
End Sub

Private Sub cboYear_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     cboMonth.SetFocus
  End If
End Sub

Private Sub cmdBasic_Click()
    
End Sub

Private Sub cmdClear_Click()

    Clear_Screen
    Combo1(1).SetFocus
    StatusBar1.Panels(1) = ""
    Label15 = "(No Advance)"
End Sub

Private Sub cmd_Click()

End Sub

Private Sub cmdClose_Click()
    Close_Msg Me
End Sub
Private Sub cmdDelete_Click()

 If Combo1(1).Text = "" Then
       MsgBox "Employee ID Required", vbInformation, organizationInfo
       Combo1(1).SetFocus
       Exit Sub
 End If
 
 
If MsgBox("Are you Sure to Delete??", vbDefaultButton1 + vbYesNo) = vbYes Then


If Utility.SalaryUpdateValidation(cboMonth, cboYear) = 1 Then
   MsgBox "Warning:You can't Delete Previous Year's Salary" + Chr(13) + Chr(13) + "Please Contact with Administrator", vbCritical, "Warning...."
   cboYear.SetFocus
   Exit Sub
ElseIf Utility.SalaryUpdateValidation(cboMonth, cboYear) = 2 Then
  MsgBox "Warning:You can't Delete Previous Month's Salary" + Chr(13) + Chr(13) + "Please Contact with Administrator", vbCritical, "Warning...."
  cboMonth.SetFocus
  Exit Sub
End If


Dim temp_pf_Loan As Integer
Dim TEMP_OTHERS As Integer
On Error GoTo Errdesc
With objSalary
    .Connstring = strCN.Connection_String
    .Emp_ID = Combo1(1)
    .Salay_Month_Get = Trim(cboMonth)
    .Salary_YearGet = cboYear
    .SalaryType = localSalaryType
    .Delete
    MsgBox "Data Deleted Successfully", vbInformation, "Software Programmer,IT Division,DNMIH"
    temp_pf_Loan = txtPF_Loan
'    Clear_Screen
    txtPF_Loan = temp_pf_Loan
    StatusBar1.Panels(1) = ""
    Combo1(1).SetFocus
End With
End If ''''end of Delete confirmation
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub cmdNewBasic_Click()
On Error GoTo Errdesc
Dim cmd As New Command
Dim rs006 As New Recordset
Dim conn006 As New Connection

 If Combo1(1).Text = "" Then
       MsgBox "Employee ID Required", vbInformation, organizationInfo
       Combo1(1).SetFocus
       Exit Sub
 End If

If Utility.SalaryUpdateValidation(cboMonth, cboYear) = 1 Then
   MsgBox "Warning:You can't Re-calculate Previous Year's Salary" + Chr(13) + Chr(13) + "Please Contact with Administrator", vbCritical, "Warning...."
   cboYear.SetFocus
   Exit Sub
ElseIf Utility.SalaryUpdateValidation(cboMonth, cboYear) = 2 Then
  MsgBox "Warning:You can't Re-calculate Previous Month's Salary" + Chr(13) + Chr(13) + "Please Contact with Administrator", vbCritical, "Warning...."
  cboMonth.SetFocus
  Exit Sub
End If



conn006.ConnectionString = strCN.Connection_String
conn006.Open
cmd.ActiveConnection = conn006
cmd.CommandType = adCmdText

cmd.CommandText = "select Basic_sal,pf_mem,to_date(to_char(JDate,'dd-mon-yyyy'),'dd-mon-yyyy') joiningDate,JTYPE from " + _
                " Emp_job_info where  emp_id='" & Combo1(1) & "'"
rs006.CursorLocation = adUseClient
rs006.Open cmd.CommandText, conn006, adOpenDynamic, adLockOptimistic

If rs006.RecordCount > 0 Then
    
    If rs006.Fields(0) <> "" Then
        txtBasic = rs006.Fields(0)
        
    Else
        txtBasic = 0
    End If
    
    If rs006.Fields(1) = 1 Then
       txtPF_Contribution = Round((Val(txtBasic) * 0.1))
    Else
       txtPF_Contribution = 0
    End If
    
    If (rs006.Fields(3).Value <> jobTypePermanent) And rs006.Fields(2).Value > CDate(allowanceLawDate) Then
       isSDAApplicable = False
    Else
       isSDAApplicable = True
    End If
    
End If

rs006.Close
conn006.Close
Dim temp As Integer
temp = txtBasic
If Mid(CboSalaryType, 1, 1) = "B" Then
  txtBonus = txtBasic
  ClearDeduction
  ClearAddition
  txtBasic = temp
  txtBonus = txtBasic
Else
  BasicSalaryOfEmp = Val(txtBasic) ''GLOBAL VAR : BasicSalaryOfEmp
  GetEffects
End If

Exit Sub
Errdesc:
      MsgBox Err.Description, vbCritical, "Problem Occured....Contact with IT"

End Sub

Private Sub cmdPreview_Click()
Dim f As New frmSalaryReportForm
f.Show 1
End Sub
Private Sub cmdPrint_Click()
Dim f1 As New frmSalaryReportForm
f1.Show 1
End Sub
Private Sub cmdSave_Click()
'On Error GoTo Errdesc
 If Combo1(1).Text = "" Then
       MsgBox "Employee ID Required", vbInformation, organizationInfo
       Combo1(1).SetFocus
       Exit Sub
 End If

If txtDesig.Text = "" Then
    MsgBox "Employee Designation Required", vbCritical, organizationInfo
    cmdShowInformation.SetFocus
    Exit Sub
End If
If Utility.SalaryUpdateValidation(cboMonth, cboYear) = 1 Then
   MsgBox "Warning:You can't Update Previous Year's Salary" + Chr(13) + Chr(13) + "Please Contact with Administrator", vbCritical, "Warning...."
   cboYear.SetFocus
   Exit Sub
ElseIf Utility.SalaryUpdateValidation(cboMonth, cboYear) = 2 Then
  MsgBox "Warning:You can't Update Previous Month's Salary" + Chr(13) + Chr(13) + "Please Contact with Administrator", vbCritical, "Warning...."
  cboMonth.SetFocus
  Exit Sub
End If
If localSalaryType = "D" And txtBasic > 0 Then
   MsgBox "Pls. Change salary type to Regular " + Chr(13) + Chr(13) + "Please Contact with Administrator", vbCritical, "Warning...."
   CboSalaryType.SetFocus
   Exit Sub
End If


BonusPreparationStatus = 0
    With Salaray_Pre
        .Connstring = strCN.Connection_String
        .Emp_ID = Combo1(1)
        .DEPT_NM = txtDept
        .designation = txtDesig
        .Emp_Nm = txtName
        .PAY_MONTH = cboMonth
        .PAY_YEAR = cboYear
        .BASIC = IIf(Mid(CboSalaryType, 1, 1) = "B", 0, txtBasic)
        .H_RENT = txtHR
        .MED = txtMed
        .CONV = txtConv
        .TFN = txtTiffin
        .DA = txtDA
        .ATTN = txtAttn
        .LEAVE = txtLeave
        .ARREAR = txtArrear
        .Bonus = txtBonus
        .OTHERS_ADDITION = txtOthers_Add
        .PF_CONTRI_DEDUCTION = txtPF_Contribution
        .PF_LN_AMOUNT = txtPF_Loan
        .SALARY_ADVANCE = txtAdvance
        .OTHERS_DEDUCTION = txtOthersDeduction
        .R_STAMP = lblRev_Stamp
        .NET_PAYABLE = txtNet_Payable
        .CREATE_BY = "IT"
        .CREATE_DATE = Date
        .UPDATE_DATE = Date
        .Remarks = txtRemarks
        .WORKING_DAY = 0
        .SALARY_DISBURSE = 0
        .adddeduct_other = txtWorkingDay
        .SDA = txtSDA
        .SalaryType = localSalaryType
        .SalaryBonusBasic = txtBasic
        .DressAllowance = txtDressAllowance
        .NDFundDeduct = txtNationalDisesterFund
        .EducationAsstAllowance = txtEducationAssisAllow
        .EmployeeDesignationLevel = Trim(txtDesignationLevel.Text)
        .Save
'        Check_Whether_Advance_Paid_ByEmp
'        Check_Whether_LoanhasTaken_By_Employee
'        If AdvanceTakenByEmp = True Then
'          Affect_Save_onAdvance_Table
'        Else
'        End If
'       If LoanTakenFlag = True Then
'
'        To_Update_Loan_Refund_Table_From_Salary
'        Else
'        End If
    End With

    With Salaray_Pre
        .Connstring = strCN.Connection_String
        .Emp_ID = Combo1(1)
        .PAY_MONTH = cboMonth
        .PAY_YEAR = cboYear
        .Employee_Contribution = Employee_Contribution
        .EMPLOYER_CONTRIBUTION = Employeer_Contribution
        .CREATE_BY = "IT"
        .CREATE_DATE = Date$
        .PF_Save
        '.Show_Message
         MsgBox "Data Saved Successfully!", vbInformation, organizationInfo
         
    End With
'End If
Combo1(1).SetFocus
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, organizationInfo

End Sub
Private Sub CommandButton1_Click()
End Sub

Private Sub cmdShowInformation_Click()
''  On Error GoTo Errdesc
   If Combo1(1).Text = "" Then
       MsgBox "Employee ID Required", vbInformation, organizationInfo
       Combo1(1).SetFocus
       Exit Sub
   Else
      Get_All_Information
   End If
   
   GetPreparedSalary
   
'  GetPreparetxtBasic = BasicSalaryOfEmpdSalary
   
'  txtBasic = BasicSalaryOfEmp
   
   
   
'    Emp_Employeer_Contribution
    
'    Get_ParameterValue
'
'    Check_Whether_LoanhasTaken_By_Employee
'    Check_Whether_Advance_Paid_ByEmp
'    Check_Whether_Advance_HasTakenAndPAid_ByEmp
'   If PFALLOWORNOT = 0 Then
'        txtPF_Contribution = 0
'    End If
'    If lblRev_Stamp = 0 Then
'
'      lblRev_Stamp = 0
'    End If
'
'
'
    
    
   
    txtConv.SetFocus

    
  Exit Sub
  
  

  
  
Errdesc:
    MsgBox Err.Description, vbInformation, organizationInfo
   
   
   
End Sub

Private Sub cmdView_Click(Index As Integer)
On Error GoTo Errdesc
Dim f2 As New frmDataSelectforSalary
Dim getconnected As New Connection
Dim cmd As New Command
Dim myrs As New ADODB.Recordset

    getconnected.ConnectionString = strCN.Connection_String
    getconnected.Open
    cmd.ActiveConnection = getconnected
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT EMP_INFO.EMP_ID,EMP_INFO.EMP_NM FROM EMP_INFO  "

    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    myrs.CursorLocation = adUseClient

    myrs.Open cmd.CommandText, getconnected, adOpenDynamic, adLockOptimistic


     Set f2.adoRecordset = myrs
     Set f2.OwnerForm = Me
     f2.Width = 6500
     f2.grdDataGrid.Columns(0).Caption = "Emp ID"
     f2.grdDataGrid.Columns(1).Caption = "Name"
     f2.grdDataGrid.Columns(0).Width = 1800
     f2.grdDataGrid.Columns(1).Width = 5500
     f2.Show 1
     Combo1(1) = myrs.Fields(0)
     txtName = myrs.Fields(1)
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub Combo1_Change(Index As Integer)
  Select Case Index
      Case 1
            ClearInfo
            ClearAddition
            ClearDeduction
End Select

End Sub

Private Sub ClearInfo()
       txtName = ""
       txtDesig = ""
       txtDept = ""
End Sub
Private Sub ClearAddition()
        txtBasic = 0
        txtHR = 0
        txtMed = 0
        txtConv = 0
        txtTiffin = 0
        txtDA = 0
        txtSDA = 0
        txtAttn = 0
        txtLeave = ""
        txtArrear = 0
        txtBonus = 0 ''bonus
        txtOthers_Add = 0
        txtDressAllowance = 0
        txtEducationAssisAllow = 0
        
  End Sub
Private Sub ClearDeduction()
        txtPF_Contribution = 0
        txtPF_Loan = 0
        txtAdvance = 0
        txtNationalDisesterFund = 0
        lblRev_Stamp = 0
        txtRemarks = ""
        txtWorkingDay = 0
        txtOthersDeduction = 0
End Sub

Private Sub Combo1_Click(Index As Integer)
  Combo1_Change (1)
End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'On Error GoTo Errdesc
'   If KeyCode = 13 Then
'        Get_All_Information
'    End If
'Exit Sub
'Errdesc:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub Command1_Click()
  
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
         Case 1
            If KeyAscii = 13 Then
             cboYear.SetFocus
            End If
       End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
    If (Shift And vbCtrlMask) > 0 Then
        Select Case KeyCode
            Case vbKeyA: txtArrear.SetFocus
            Case vbKeyT: txtBonus.SetFocus
            Case vbKeyO: txtOthers_Add.SetFocus
            Case vbKeyH: txtNationalDisesterFund.SetFocus
            Case vbKeyR: txtRemarks.SetFocus
            Case vbKeyS: cmdSave.SetFocus
            Case vbKeyI: Combo1(1).SetFocus
            Case vbKeyI: txtAttn.SetFocus
            
        End Select
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 10 Then
      Label17_Click
  ElseIf KeyAscii = 18 Then
     cmdNewBasic_Click
  ElseIf KeyAscii = 14 Then
     Combo1(1).SetFocus
  ElseIf KeyAscii = 16 Then
     cmdSave_Click
  End If
 
End Sub

Private Sub Form_Load()
On Error GoTo Errdes
    Screen_Position Me
    localSalaryType = "R"
    CboSalaryType.ListIndex = 0
    Load_Yr Me
    Load_MonthNm Me
    'lblRev_Stamp = 4
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
     
        Do Until rs10.EOF
            Combo1(1).AddItem rs10.Fields(0)
            rs10.MoveNext
        Loop
        
     
    End If
    
    rs10.Close
    conn10.Close
     
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "Software Programmer,IT Division,DNMIH"
End Sub
Private Sub lblHol_OT_Click()
    Calculate
End Sub

Private Sub lblOT_Click()
    Calculate
End Sub

Private Sub txtAtt_Change()
    Calculate
End Sub

Private Sub txtAtt_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub Label17_Click()
 If Combo1(1).Text = "" Then
       MsgBox "Employee ID Required", vbInformation, organizationInfo
       Combo1(1).SetFocus
       Exit Sub
 End If
 
  Form2.SSTab1.Tab = 1
  Form2.txtEmp_ID(1) = Form3.Combo1(1).Text
  Form2.Combo1(0).Text = Form3.Combo1(1).Text
  Form2.Show 1
   
End Sub

Private Sub lblRev_Stamp_Change()
 Default_Zero txtNationalDisesterFund
    Calculate
End Sub

Private Sub lblRev_Stamp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'     lblRev_Stamp.Text = Format(lblRev_Stamp, "#,##0.00")
    txtOthersDeduction.SetFocus
End If
End Sub

Private Sub oPTION_SalaryType_Click(Index As Integer)
  Select Case Index
        Case 0
           localSalaryType = "R"
        Case 1
           localSalaryType = "S"
  End Select
 
End Sub

Private Sub txtEducationAssisAllow_Change()
  Default_Zero txtEducationAssisAllow
  Calculate
End Sub
Private Sub txtEducationAssisAllow_GotFocus()
  txtEducationAssisAllow.SelStart = 0
  txtEducationAssisAllow.SelLength = Len(txtEducationAssisAllow.Text)
End Sub
Private Sub txtOthersDeduction_Change()
'txtLeave = Val(txtOthersDeduction) - Val(txtAttn)
 Default_Zero txtAdvance
 Calculate
End Sub

Private Sub txtOthersDeduction_GotFocus()
txtOthersDeduction.SelStart = 0
txtOthersDeduction.SelLength = Len(txtOthersDeduction.Text)
End Sub

Private Sub txtOthersDeduction_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'    txtOthersDeduction.Text = Format(txtOthersDeduction, "#,##0.00")
    cmdSave.SetFocus
End If
End Sub

Private Sub txtOthersDeduction_LostFocus()
 If txtOthersDeduction.Text = "" Then
    txtOthersDeduction.Text = 0
 End If
End Sub

Private Sub txtAdvance_Change()
    Default_Zero txtAdvance
    Calculate
End Sub

Private Sub txtAdvance_GotFocus()
txtAdvance.SelStart = 0
txtAdvance.SelLength = Len(txtAdvance.Text)
End Sub

Private Sub txtAdvance_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'    txtAdvance.Text = Format(txtAdvance, "#,##0.00")
    txtNationalDisesterFund.SetFocus
End If
End Sub

Private Sub txtArrear_Change()
    Default_Zero txtArrear
    Calculate
End Sub

Private Sub txtArrear_GotFocus()
txtArrear.SelStart = 0
txtArrear.SelLength = Len(txtArrear.Text)
End Sub

Private Sub txtArrear_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'     txtArrear.Text = Format(txtArrear, "#,##0.00")
     txtBonus.SetFocus
End If
End Sub

Private Sub txtAttn_Change()
If Len(Trim(txtAttn)) = 0 Then txtAttn = 0

txtLeave = Val(txtOthersDeduction) - Val(txtAttn)
End Sub
Private Sub txtAttn_LostFocus()
If Val(txtAttn) > Val(txtOthersDeduction.Text) Then
    MsgBox "Invalid Input !", vbInformation, "IT Division, DNMIH"
    txtAttn.SetFocus
    Exit Sub
End If
End Sub

Private Sub txtBasic_Change()
 If localSalaryType = "R" Then
   GetEffects
 End If
End Sub
Private Sub GetEffects()
   txtDA = 0
   If isSDAApplicable = True Then
     txtSDA = Round(Val(txtBasic) * 0.35)
   Else
     txtSDA = 0
   End If
   Get_HouseRent
   Get_Convence
End Sub
Private Sub txtBasic_Click()
  GetEffects
End Sub

Private Sub txtBasic_GotFocus()
txtBasic.SelStart = 0
txtBasic.SelLength = Len(txtBasic.Text)
End Sub

Private Sub txtConv_Change()
 Default_Zero txtOthers_Add
    Calculate
End Sub

Private Sub txtConv_GotFocus()
txtMed.SelStart = 0
txtMed.SelLength = Len(txtMed.Text)
End Sub

Private Sub txtConv_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'     txtConv.Text = Format(txtConv, "#,##0.00")
    txtTiffin.SetFocus
End If
End Sub

Private Sub txtDA_Change()
     Default_Zero txtDA
    Calculate
End Sub

Private Sub txtElec_Bill_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
Public Sub Calculate()
On Error GoTo Errdesc
    txtTotal = Val(txtBasic) + Val(txtHR) + Val(txtConv) + Val(txtMed) _
            + Val(txtDA) + Val(txtTiffin) + Val(txtOthers_Add) _
            + Val(txtArrear) + Val(txtBonus) + Val(txtSDA) _
            + Val(txtDressAllowance) + Val(txtEducationAssisAllow)



    txtTotal_Deduction = Val(txtPF_Contribution) + Val(txtAdvance) _
            + Val(lblRev_Stamp) + Val(txtNationalDisesterFund) + Val(txtPF_Loan) + Val(txtOthersDeduction.Text)
            
    txtNet_Payable = Val(txtTotal) - Val(txtTotal_Deduction)

Exit Sub

Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub txtDA_GotFocus()
txtDA.SelStart = 0
txtDA.SelLength = Len(txtDA.Text)
End Sub

Private Sub txtDA_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtArrear.SetFocus
End If
End Sub

Private Sub txtDressAllowance_Change()
  Default_Zero txtDressAllowance
   Calculate
End Sub

Private Sub txtDressAllowance_GotFocus()
  txtDressAllowance.SelStart = 0
txtDressAllowance.SelLength = Len(txtDressAllowance.Text)
End Sub

Private Sub txtHR_Change()
 Default_Zero txtOthers_Add
    Calculate
End Sub

Private Sub txtHR_GotFocus()
txtHR.SelStart = 0
txtHR.SelLength = Len(txtHR.Text)
End Sub

Private Sub txtHR_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtMed.SetFocus
End If
End Sub

Private Sub txtMed_Change()
 Default_Zero txtOthers_Add
    Calculate
End Sub

Private Sub txtMed_GotFocus()
txtMed.SelStart = 0
txtMed.SelLength = Len(txtMed.Text)
End Sub

Private Sub txtMed_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtConv.SetFocus
End If
End Sub

Private Sub txtOthers_Add_GotFocus()
txtOthers_Add.SelStart = 0
txtOthers_Add.SelLength = Len(txtOthers_Add.Text)
End Sub

Private Sub txtOthers_Add_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtPF_Contribution.SetFocus
End If
End Sub

Private Sub txtNationalDisesterFund_Change()
    Default_Zero txtNationalDisesterFund
    Calculate
End Sub

Private Sub txtNationalDisesterFund_GotFocus()
txtNationalDisesterFund.SelStart = 0
txtNationalDisesterFund.SelLength = Len(txtNationalDisesterFund.Text)
End Sub

Private Sub txtNationalDisesterFund_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    lblRev_Stamp.SetFocus
End If
End Sub

Private Sub txtNationalDisesterFund_KeyPress(KeyAscii As Integer)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
Private Sub txtOther_Add_Change()
   ' Default_Zero txtOther_Add
    Calculate
End Sub
Private Sub txtOther_Add_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
Public Sub Default_Zero(txt As TextBox)
    If Len(txt) < 1 Then txt = 0
End Sub
Private Sub txtWF_Ln_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
Private Sub txtOthers_Add_Change()
    Default_Zero txtOthers_Add
    Calculate
End Sub

Private Sub txtPF_Contribution_Change()
  Default_Zero txtPF_Loan
  Calculate
End Sub

Private Sub txtPF_Contribution_GotFocus()
txtPF_Contribution.SelStart = 0
txtPF_Contribution.SelLength = Len(txtPF_Contribution.Text)
End Sub

Private Sub txtPF_Contribution_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'    txtPF_Contribution.Text = Format(txtPF_Contribution, "#,##0.00")
    txtPF_Loan.SetFocus
End If
End Sub

Private Sub txtPF_Contribution_LostFocus()
  If txtPF_Contribution.Text = "" Then
     txtPF_Contribution.Text = 0
  End If
End Sub

Private Sub txtPF_Loan_Change()
    Default_Zero txtPF_Loan
    Calculate
End Sub

Private Sub txtPF_Loan_GotFocus()
txtPF_Loan.SelStart = 0
txtPF_Loan.SelLength = Len(txtPF_Loan.Text)
End Sub

Private Sub txtPF_Loan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'    txtPF_Loan.Text = Format(txtPF_Loan, "#,##0.00")
    txtAdvance.SetFocus
End If
End Sub

Private Sub txtSDA_Change()
   Default_Zero txtSDA
   Calculate
End Sub

Private Sub txtBonus_Change()
'    Default_Zero txtBonus
    If Mid(CboSalaryType, 1, 1) = "B" Then
       txtTotal = txtBonus
    Else
     Calculate
    End If
End Sub
Private Sub Get_All_Information()
On Error GoTo Errdesc
Dim cmd As New Command
conn1.ConnectionString = strCN.Connection_String
conn1.Open
cmd.ActiveConnection = conn1
cmd.CommandType = adCmdText
cmd.CommandText = "SELECT ST_DEPT.DEPT_NM,EMP_JOB_INFO.SCALE_CODE,EMP_JOB_INFO.BASIC_SAL, " + _
                " St_Desig.DESIGNATION , emp_info.Emp_Nm,EMP_JOB_INFO.MODE_OF_PAYMENT,emp_job_info.PF_MEM,EMP_JOB_INFO.EMP_CLASS,EMP_JOB_INFO.EMP_Position,EMP_JOB_INFO.EMP_Designation_Level " + _
                " From emp_info, ST_DEPT, St_Desig, EMP_JOB_INFO " + _
                " WHERE ((EMP_JOB_INFO.EMP_ID=EMP_INFO.EMP_ID) " + _
                " AND (EMP_JOB_INFO.DESIG=ST_DESIG.DESIG_CODE) " + _
                " AND (EMP_JOB_INFO.DEPT=ST_DEPT.DEPT_CODE)  and emp_info.EMP_ID='" & Combo1(1) & "')"

rs1.CursorLocation = adUseClient
rs1.Open cmd.CommandText, conn1, adOpenDynamic, adLockOptimistic

    If rs1.RecordCount > 0 Then
        txtName = rs1.Fields(4)
        txtDesig = rs1.Fields(3)
        txtDept = rs1.Fields(0)
        BasicSalaryOfEmp = rs1.Fields(2)
        EmpScaleID = rs1.Fields(1)
        PFALLOWORNOT = rs1.Fields(6)
        ClassofEmployee = rs1.Fields(7)
        txtStaffSerial = rs1.Fields(8)
        txtDesignationLevel = rs1.Fields(9)
        rs1.Close
        conn1.Close
        cboYear.SetFocus
    Else
        MsgBox "Invalid Employee No.", vbInformation, organizationInfo
        rs1.Close
        conn1.Close
        Exit Sub
    End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, organizationInfo
End Sub
Private Sub Emp_Employeer_Contribution()
On Error GoTo Errdesc
Dim cmd As New Command
conn2.ConnectionString = strCN.Connection_String
conn2.Open
cmd.ActiveConnection = conn2
cmd.CommandType = adCmdText

cmd.CommandText = " select EMPCONTRPF,EMRCONTRPF,EFFDATE from PARAMETER_MAIN where effdate=(select max(EFFDATE)from PARAMETER_MAIN)"

RS2.CursorLocation = adUseClient
RS2.Open cmd.CommandText, conn2, adOpenDynamic, adLockOptimistic

    If RS2.RecordCount > 0 Then
        EmpContr = RS2.Fields(0)
        EmployeerContr = RS2.Fields(1)
        txtBasic = BasicSalaryOfEmp
        Employee_Contribution = BasicSalaryOfEmp * EmpContr / 100
        Employeer_Contribution = BasicSalaryOfEmp * EmpContr / 100
        txtPF_Contribution = Round(Employee_Contribution)
        
    End If
        RS2.Close
        conn2.Close
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Get_ParameterValue()
On Error GoTo Errdesc
Dim cmd As New Command
conn3.ConnectionString = strCN.Connection_String
conn3.Open
cmd.ActiveConnection = conn3
cmd.CommandType = adCmdText

cmd.CommandText = " select HR,MED,CONV,TFN,STR_BASIC from ST_PAYSCALE where SCALE_CODE='" & EmpScaleID & "'"

rs3.CursorLocation = adUseClient
rs3.Open cmd.CommandText, conn3, adOpenDynamic, adLockOptimistic

    If rs3.RecordCount > 0 Then
        
        HouseRentValue = rs3.Fields(0) * rs3.Fields(4) / 100
       
       If rs3.Fields("STR_BASIC") < 1800 Then
            If HouseRentValue < 850 Then
                HouseRentValue = 850
            Else
                HouseRentValue = rs3.Fields(0) * rs3.Fields(4) / 100
            End If
       ElseIf rs3.Fields("STR_BASIC") > 1800 And rs3.Fields("STR_BASIC") < 3800 Then
            If HouseRentValue < 990 Then
                HouseRentValue = 990
            Else
                HouseRentValue = rs3.Fields(0) * rs3.Fields(4) / 100
            End If
       ElseIf rs3.Fields("STR_BASIC") >= 3801 And rs3.Fields("STR_BASIC") < 9000 Then
            If HouseRentValue < 1900 Then
                HouseRentValue = 1990
            Else
                HouseRentValue = rs3.Fields(0) * rs3.Fields(4) / 100
            End If
       ElseIf rs3.Fields("STR_BASIC") >= 9001 Then
            If HouseRentValue < 4050 Then
                HouseRentValue = 4050
            Else
                HouseRentValue = rs3.Fields(0) * rs3.Fields(4) / 100
            End If
       
       End If
              
       
       
       
        If rs3.Fields(1) = "" Then
            MedicalValue = 0
        Else
             MedicalValue = rs3.Fields(1)
        End If
        
        If rs3.Fields(2) = "" Then
            ConvenceCalue = 0
        Else
            ConvenceCalue = rs3.Fields(2)
        End If
        
        If rs3.Fields(3) = "" Then
            TiffinValue = 0
        Else
            TiffinValue = rs3.Fields(3)
        End If
        
        
'        txtDA = 10 * BasicSalaryOfEmp / 100
'
'
'        If Val(txtDA) < 200 Then
'            txtDA = 200
'        End If
'
        txtDA = 0
        
        
        
        txtHR = HouseRentValue
        txtMed = MedicalValue
        txtConv = ConvenceCalue
        txtTiffin = TiffinValue
        'txtDA = 0
        txtArrear = 0
        txtBonus = 0
        txtOthers_Add = 0
'        txtPF_Loan = 0
        txtAdvance = 0
        txtNationalDisesterFund = 0
        lblRev_Stamp = 0
        txtOthersDeduction.Text = 0
        txtWorkingDay.Text = 0
        'txtMed = 500
        'txtDA.SetFocus
        txtArrear.SetFocus
        
    End If
    
    rs3.Close
        conn3.Close
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub GetPreparedSalary()
'On Error GoTo Errdesc
Dim cmd As New Command
Dim conn005 As New Connection
ChechWhetherSlaryPaidirNot = Empty

conn005.ConnectionString = strCN.Connection_String
conn005.Open
cmd.ActiveConnection = conn005
cmd.CommandType = adCmdText

cmd.CommandText = "select *  from Salary_Preparation where  emp_id='" & Combo1(1) & "' AND PAY_MONTH='" & cboMonth & "' AND PAY_YEAR='" & cboYear & "' and Salary_Type='" & Mid(CboSalaryType, 1, 1) & "'"
rs005.CursorLocation = adUseClient
rs005.Open cmd.CommandText, conn005, adOpenDynamic, adLockOptimistic

If rs005.RecordCount > 0 Then
  
    ChechWhetherSlaryPaidirNot = rs005.Fields(0)
   If Not IsNull(rs005.Fields(0)) Then
         StatusBar1.Panels(1) = "Salary has been Preapared for " & rs005.Fields(1) & "  " & "In this Month"
         
 ''''addition
        txtBasic = rs005.Fields("BASIC")
        
        txtHR = rs005.Fields("H_RENT")
        txtMed = rs005.Fields("MED")
        txtConv = rs005.Fields("CONV")
        txtTiffin = rs005.Fields("TFN")
        txtDA = rs005.Fields("DA")
        txtSDA = rs005.Fields("S_DA")
        txtAttn = rs005.Fields("ATTN")
        txtLeave = "" & rs005.Fields("LEAVE")
        txtArrear = rs005.Fields("ARREAR")
        txtBonus = "" & rs005.Fields("Bonus_Allowance") ''bonus
        txtDressAllowance = "" & rs005.Fields("Dress_Allowance")
        txtOthers_Add = rs005.Fields("OTHERS_ADDITION")
        txtEducationAssisAllow = rs005.Fields("edu_asst_allowance")
  '''deduction
        txtPF_Contribution = rs005.Fields("PF_CONTRI_DEDUCTION")
        txtPF_Loan = rs005.Fields("PF_LN_AMOUNT")
        txtAdvance = rs005.Fields("SALARY_ADVANCE")
        txtNationalDisesterFund = rs005.Fields("ND_FUND_DEDUCT")
        lblRev_Stamp = rs005.Fields("R_STAMP")
        
   ''''payable and remarks
        txtNet_Payable = rs005.Fields("NET_PAYABLE")
        txtRemarks = "" & rs005.Fields("Remarks")
        txtOthersDeduction = rs005.Fields("Others_Deduction") ''''others
        txtWorkingDay = 0
       
       End If
       If Mid(CboSalaryType, 1, 1) = "B" Then
          txtBasic = rs005.Fields("FESTIVAL_BONUS_BASIC")
       End If
    
  Else
         ClearAddition
         ClearDeduction
         StatusBar1.Panels(1) = "Salary has not been Prapared for this Month"
   
End If

    rs005.Close
    conn005.Close

   
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, organizationInfo
End Sub
Private Sub To_Update_Loan_Refund_Table_From_Salary()
On Error GoTo Errdesc
   With Salaray_Pre
        .Connstring = strCN.Connection_String
              .Emp_ID = Combo1(1)
              .LoanRefundedDate = Date
              .NoOfInstallmentPaid = NoofInstallmentPaidByEmp + 1
              .AmountPaid = SlabAmountofLoan
              .Notes = "Loan has Adjusted from Salary"
              .LoanRefundNo = "Ln Refund"
              .EntrDate = Date
              .Loan_Sub_Save
          
           End With
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Check_Whether_LoanhasTaken_By_Employee()
On Error GoTo Errdesc
Dim cmd As New Command
Dim conn44 As New Connection
Dim rs44 As New Recordset
conn44.ConnectionString = strCN.Connection_String
conn44.Open
cmd.ActiveConnection = conn44
cmd.CommandType = adCmdText
cmd.CommandText = " SELECT LOANINFORMATION_MAIN.EMP_ID, " + _
                " LOANINFORMATION_MAIN.NOOFINSTALLMENT,LOANINFORMATION_MAIN.SLABINSTALLMENTAMOUNT," + _
                " COUNT(LOANINFORMATION_SUB.AmountPaid) As TOTAL_PAID,LOANINFORMATION_MAIN.ISSUEDAMOUNT " + _
                " From LOANINFORMATION_MAIN, LOANINFORMATION_SUB " + _
                " Where (LOANINFORMATION_MAIN.EMP_ID = LOANINFORMATION_SUB.EMP_ID) " + _
                " AND (LOANINFORMATION_MAIN.EMP_ID='" & Combo1(1) & "')" + _
                " GROUP BY LOANINFORMATION_MAIN.EMP_ID,LOANINFORMATION_MAIN.NOOFINSTALLMENT, " + _
                " LOANINFORMATION_MAIN.SlabInstallmentAmount,LOANINFORMATION_MAIN.ISSUEDAMOUNT "

rs44.CursorLocation = adUseClient
rs44.Open cmd.CommandText, conn44, adOpenDynamic, adLockOptimistic

If rs44.RecordCount > 0 Then
      If Not IsNull(rs44.Fields(0)) Then
         TotalInstallment = rs44.Fields(4)
         SlabAmountofLoan = rs44.Fields(2)
         NoofInstallmentPaidByEmp = rs44.Fields(3)
         '---Calculate Total Amount of Loan Paid by Employee
         TotalLoanPaidByEmp = SlabAmountofLoan * NoofInstallmentPaidByEmp
         '---Calculate Total Amount of Loan Taken by Employee
          AmountofLoanTaken = TotalInstallment * SlabAmountofLoan
            If TotalLoanPaidByEmp >= AmountofLoanTaken Then
                   MsgBox "There is no Loan Amount dues of this Employee", vbInformation, "IT Division, DNMIH"
                   Exit Sub
            Else
                txtPF_Loan = SlabAmountofLoan
                                        '--- To_Update_Loan_Refund_Table_From_Salary
                txtRemarks = " Total " & NoofInstallmentPaidByEmp & " Installment was paid out of " & TotalInstallment
                
            End If
     End If
    
'LoanTakenFlag = 1
    'as else ------------------ one more validation has to give
Else
        Emp_Exit_In_Main
        Emp_Exit_In_Sub
       ' ---NoofInstallmentPaidByEmp = 0
        Extra_Validastion_For_Loan_Take
'LoanTakenFlag = 0
End If
     rs44.Close
    conn44.Close

    
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Check_Whether_Advance_Paid_ByEmp()
On Error GoTo Errdesc
Dim cmd As New Command
conn007.ConnectionString = strCN.Connection_String
conn007.Open
cmd.ActiveConnection = conn007
cmd.CommandType = adCmdText
cmd.CommandText = "select  SALARY_ADVANCE   from " + _
                " Salary_Preparation where  emp_id='" & Combo1(1) & "' AND PAY_MONTH='" & cboMonth & "' AND PAY_YEAR='" & cboYear & "'"
RS007.CursorLocation = adUseClient
RS007.Open cmd.CommandText, conn007, adOpenDynamic, adLockOptimistic

If RS007.RecordCount > 0 Then
        DedeductionAdvanceEachMonth = RS007.Fields(0)
Else
        DedeductionAdvanceEachMonth = 0
        
End If

RS007.Close
conn007.Close

Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Affect_Save_onAdvance_Table()
On Error GoTo Errdesc
With Advance_Info
            .Connstring = strCN.Connection_String
            .Emp_ID = Combo1(1)
            .Adv_issue_dt = Date$
            .Adv_Amt = AdvanceTakenAmount
            .Num_Inst = NoofIstalltimeAdvance - 1
            .Notes = "From Salary Advance has Adjusted"
            .Balance = AdvanceTakenAmount - DedeductionAdvanceEachMonth
            .Save
        End With
        
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
    
End Sub
Private Sub Check_Whether_Advance_HasTakenAndPAid_ByEmp()
On Error GoTo Errdesc

Dim cmd As New Command
Dim rs8 As New Recordset
Dim conn04 As New Connection
conn04.ConnectionString = strCN.Connection_String
conn04.Open
cmd.ActiveConnection = conn04
cmd.CommandType = adCmdText

cmd.CommandType = adCmdText
cmd.CommandText = "select EMP_ID,ADV_AMT,NUM_INST,PAID_INSTALLED,BALANCE from Advance_Info " + _
                    "where emp_id='" & Combo1(1) & "' and BALANCE<>0 and track_id=(select max(track_id) from Advance_Info where emp_id='" & Combo1(1) & "')"

rs8.CursorLocation = adUseClient
rs8.Open cmd.CommandText, conn04, adOpenDynamic, adLockOptimistic

If rs8.RecordCount > 0 Then
    AdvanceTakenAmount = rs8.Fields(1)
    NoofIstalltimeAdvance = rs8.Fields(2)
    PaidAdvanceInstallment = rs8.Fields(3)
    AdvanceBalanceAmount = rs8.Fields(4)
    Label15 = "Tk.(" & AdvanceBalanceAmount & ")"
    txtAdvance.SetFocus
    AdvanceTakenByEmp = 1
Else
    AdvanceTakenByEmp = 0
    Label15 = ""
End If
    
    rs8.Close
    conn04.Close
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Emp_Exit_In_Main()
On Error GoTo Errdesc

Dim cmd As New Command
Dim rs8 As New Recordset
Dim conn04 As New Connection
conn04.ConnectionString = strCN.Connection_String
conn04.Open
cmd.ActiveConnection = conn04
cmd.CommandType = adCmdText

cmd.CommandType = adCmdText
cmd.CommandText = "select emp_id from loaninformation_main where emp_id='" & Combo1(1) & "'"

rs8.CursorLocation = adUseClient
rs8.Open cmd.CommandText, conn04, adOpenDynamic, adLockOptimistic

If rs8.RecordCount > 0 Then
    EmpIdExitinMain = rs8.Fields(0)
End If
    
    rs8.Close
    conn04.Close
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Emp_Exit_In_Sub()
On Error GoTo Errdesc

Dim cmd As New Command
Dim rs8 As New Recordset
Dim conn04 As New Connection
conn04.ConnectionString = strCN.Connection_String
conn04.Open
cmd.ActiveConnection = conn04
cmd.CommandType = adCmdText

cmd.CommandType = adCmdText
cmd.CommandText = "select emp_id from loaninformation_sub where emp_id='" & EmpIdExitinMain & "'"

rs8.CursorLocation = adUseClient
rs8.Open cmd.CommandText, conn04, adOpenDynamic, adLockOptimistic

If rs8.RecordCount = 0 Then
   Extra_Validastion_For_Loan_Take
    NoofInstallmentPaidByEmp = 0
   '-------------------To_Update_Loan_Refund_Table_From_Salary
End If
    
    rs8.Close
    conn04.Close
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Extra_Validastion_For_Loan_Take()
On Error GoTo Errdesc

Dim cmd As New Command
Dim rs8 As New Recordset
Dim conn04 As New Connection
conn04.ConnectionString = strCN.Connection_String
conn04.Open
cmd.ActiveConnection = conn04
cmd.CommandType = adCmdText

cmd.CommandType = adCmdText
cmd.CommandText = "select SLABINSTALLMENTAMOUNT from loaninformation_main where emp_id='" & Combo1(1) & "'"

rs8.CursorLocation = adUseClient
rs8.Open cmd.CommandText, conn04, adOpenDynamic, adLockOptimistic

If rs8.RecordCount > 0 Then
   SlabAmountofLoan = rs8.Fields(0)
   txtPF_Loan = SlabAmountofLoan
   LoanTakenFlag = 1
   Else
   LoanTakenFlag = 0
End If
    
    rs8.Close
    conn04.Close
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub txtBonus_GotFocus()
txtBonus.SelStart = 0
txtBonus.SelLength = Len(txtBonus.Text)
End Sub

Private Sub txtBonus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtOthers_Add.SetFocus
End If
End Sub

Private Sub txtTiffin_Change()
 Default_Zero txtOthers_Add
 Calculate
End Sub

'Private Sub Check_For_SalaryIncrement_And_Dues()
'On Error GoTo Errdesc
'
'Dim cmd As New Command
'Dim rs8 As New Recordset
'Dim conn04 As New Connection
'conn04.ConnectionString = strCN.Connection_String
'conn04.Open
'cmd.ActiveConnection = conn04
'cmd.CommandType = adCmdText
'
'cmd.CommandType = adCmdText
'
'cmd.CommandText = "select EMP_ID,AMOUNT ,LAST_DT_INCRE from INCREMENT_RECORD where " + _
'              " TRACK_ID=(select max(TRACK_ID) from INCREMENT_RECORD where emp_id='" & Combo1(1) & "')"
'
'rs8.CursorLocation = adUseClient
'rs8.Open cmd.CommandText, conn04, adOpenDynamic, adLockOptimistic
'
'If rs8.RecordCount > 0 Then
'   SlabAmountofLoan = rs8.Fields(0)
'   txtPF_Loan = SlabAmountofLoan
'End If
'
'    rs8.Close
'    conn04.Close
'Exit Sub
'Errdesc:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub

Private Sub Get_HouseRent()
On Error GoTo Errdes
If BasicSalaryOfEmp >= 0 And BasicSalaryOfEmp < 2000 Then
      EmployeeeHousrRent = 0
ElseIf BasicSalaryOfEmp >= 2000 And BasicSalaryOfEmp <= 5000 Then
    EmployeeeHousrRent = BasicSalaryOfEmp * 65 / 100

    If EmployeeeHousrRent < 2800 Then
        EmployeeeHousrRent = 2800
    Else
       EmployeeeHousrRent = BasicSalaryOfEmp * 65 / 100
    End If


ElseIf BasicSalaryOfEmp >= 5001 And BasicSalaryOfEmp <= 10800 Then
    EmployeeeHousrRent = BasicSalaryOfEmp * 60 / 100

    If EmployeeeHousrRent < 3300 Then
        EmployeeeHousrRent = 3300
    Else
       EmployeeeHousrRent = BasicSalaryOfEmp * 60 / 100
    End If
ElseIf BasicSalaryOfEmp >= 10801 And BasicSalaryOfEmp <= 21600 Then
    EmployeeeHousrRent = BasicSalaryOfEmp * 55 / 100

    If EmployeeeHousrRent < 6500 Then
        EmployeeeHousrRent = 6500
    Else
       EmployeeeHousrRent = BasicSalaryOfEmp * 55 / 100
    End If
ElseIf BasicSalaryOfEmp > 21601 Then
    EmployeeeHousrRent = BasicSalaryOfEmp * 50 / 100

    If EmployeeeHousrRent < 11900 Then
        EmployeeeHousrRent = 11900
    Else
       EmployeeeHousrRent = BasicSalaryOfEmp * 50 / 100
    End If

End If

txtHR = Round(EmployeeeHousrRent)

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, organizationInfo
End Sub
Private Sub Get_Convence()
On Error GoTo Errdes
If ClassofEmployee <> "1" Then
    txtConv = 150
Else
    txtConv = 0
End If
       txtMed = 700
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, organizationInfo
End Sub
'Private Sub Check_Salary_Given_Or_not()
'On Error GoTo Errdesc
'Dim cmd As New Command
'Dim conn005 As New Connection
'ChechWhetherSlaryPaidirNot = Empty
'
'conn005.ConnectionString = strCN.Connection_String
'conn005.Open
'cmd.ActiveConnection = conn005
'cmd.CommandType = adCmdText
'
'cmd.CommandText = "select *  from Salary_Preparation where  emp_id='" & Combo1(1) & "' AND PAY_MONTH='" & cboMonth & "' AND PAY_YEAR='" & cboYear & "'"
'rs005.CursorLocation = adUseClient
'rs005.Open cmd.CommandText, conn005, adOpenDynamic, adLockOptimistic
'
'If rs005.RecordCount > 0 Then
'
'    ChechWhetherSlaryPaidirNot = rs005.Fields(0)
'    If Not IsNull(rs005.Fields(0)) Then
'End Sub

Private Sub txtTiffin_GotFocus()
txtTiffin.SelStart = 0
txtTiffin.SelLength = Len(txtTiffin.Text)
End Sub

Private Sub txtTiffin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtDA.SetFocus
End If
End Sub

Private Sub txtTotal_Change()
  If Mid(CboSalaryType, 1, 1) = "B" Then
     txtTotal = txtBonus
     txtNet_Payable = txtTotal
  End If
End Sub
