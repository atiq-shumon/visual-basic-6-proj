VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBonusPreparation 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   $"frmBonusPreparation.frx":0000
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10710
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   67
      Top             =   6480
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   18697
            MinWidth        =   18697
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   5685
      Index           =   0
      Left            =   165
      TabIndex        =   10
      Top             =   0
      Width           =   10455
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
         Left            =   1710
         TabIndex        =   35
         Top             =   1260
         Width           =   915
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
         Left            =   2620
         TabIndex        =   34
         Top             =   1260
         Width           =   1320
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
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2160
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
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2475
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
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3105
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
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3420
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
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   3735
         Width           =   915
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
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   4050
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
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   4680
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
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2790
         Width           =   915
      End
      Begin VB.TextBox txtOthers_Deduct 
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
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3105
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
         TabIndex        =   24
         Top             =   2160
         Width           =   915
      End
      Begin VB.TextBox txtPF_Loan 
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
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2475
         Width           =   915
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
         Locked          =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   22
         Top             =   2790
         Width           =   915
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
         Left            =   7965
         TabIndex        =   21
         Text            =   "0"
         Top             =   1305
         Width           =   465
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   1740
         Left            =   7275
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2160
         Width           =   2940
      End
      Begin VB.TextBox txtTelephone 
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
         TabIndex        =   1
         Top             =   4365
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   5130
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
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   5130
         Width           =   915
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
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   5130
         Width           =   915
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
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   405
         Width           =   5685
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   855
         Width           =   3840
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
         Left            =   7020
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   855
         Width           =   3075
      End
      Begin VB.TextBox Text1 
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
         Left            =   9630
         TabIndex        =   14
         Top             =   1305
         Width           =   465
      End
      Begin VB.CommandButton cmdView 
         Height          =   330
         Index           =   0
         Left            =   3095
         Picture         =   "frmBonusPreparation.frx":00CA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   340
         Width           =   375
      End
      Begin VB.TextBox Text2 
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
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3740
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
         TabIndex        =   2
         Top             =   3420
         Width           =   915
      End
      Begin VB.TextBox Text3 
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
         Left            =   5955
         TabIndex        =   11
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year/ Month"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   405
         TabIndex        =   66
         Top             =   1305
         Width           =   900
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
         TabIndex        =   65
         Top             =   2490
         Width           =   825
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
         TabIndex        =   64
         Top             =   2115
         Width           =   870
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
         TabIndex        =   63
         Top             =   2805
         Width           =   1140
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
         TabIndex        =   62
         Top             =   3105
         Width           =   960
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
         TabIndex        =   61
         Top             =   3420
         Width           =   825
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
         TabIndex        =   60
         Top             =   2160
         Width           =   1080
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
         TabIndex        =   59
         Top             =   2475
         Width           =   600
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus && Allowances"
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
         Left            =   345
         TabIndex        =   58
         Top             =   1710
         Width           =   1950
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
         Left            =   3630
         TabIndex        =   57
         Top             =   1710
         Width           =   1665
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
         Left            =   405
         TabIndex        =   56
         Top             =   4680
         Width           =   690
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
         TabIndex        =   55
         Top             =   2790
         Width           =   1125
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Atten&dance"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6750
         TabIndex        =   54
         Top             =   1305
         Width           =   825
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
         TabIndex        =   53
         Top             =   3405
         Width           =   1155
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
         TabIndex        =   52
         Top             =   5175
         Width           =   1140
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Payable"
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   32
         Left            =   405
         TabIndex        =   51
         Top             =   5100
         Width           =   1215
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
         Left            =   5940
         TabIndex        =   50
         Top             =   5175
         Width           =   870
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
         Left            =   3600
         TabIndex        =   49
         Top             =   4335
         Width           =   4605
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DA"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   405
         TabIndex        =   48
         Top             =   3735
         Width           =   825
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Arrear"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   405
         TabIndex        =   47
         Top             =   4050
         Width           =   825
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N.D Fund (-)"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3240
         TabIndex        =   46
         Top             =   3105
         Width           =   1005
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         Caption         =   "Employee  &ID"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   405
         TabIndex        =   45
         Top             =   405
         Width           =   1140
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
         Left            =   6990
         TabIndex        =   44
         Top             =   1710
         Width           =   1500
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Name"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3645
         TabIndex        =   43
         Top             =   405
         Width           =   420
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
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   0
         Left            =   1710
         Top             =   2115
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   1
         Left            =   1710
         Top             =   2430
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   2
         Left            =   1710
         Top             =   3060
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   3
         Left            =   1710
         Top             =   3375
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   4
         Left            =   1710
         Top             =   3690
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   5
         Left            =   1710
         Top             =   4005
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   6
         Left            =   1710
         Top             =   4635
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   7
         Left            =   1710
         Top             =   2745
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   8
         Left            =   4590
         Top             =   3060
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   9
         Left            =   4590
         Top             =   2115
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   10
         Left            =   4590
         Top             =   2430
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   11
         Left            =   4590
         Top             =   2745
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   14
         Left            =   7920
         Top             =   1260
         Width           =   555
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   1920
         Index           =   15
         Left            =   7110
         Top             =   2115
         Width           =   3150
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   16
         Left            =   7245
         Top             =   5085
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   17
         Left            =   4590
         Top             =   3375
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   18
         Left            =   4590
         Top             =   5085
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   19
         Left            =   1710
         Top             =   5085
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   20
         Left            =   4275
         Top             =   360
         Width           =   5865
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
         TabIndex        =   42
         Top             =   3915
         Width           =   915
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   12
         Left            =   6840
         Top             =   810
         Width           =   3300
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   21
         Left            =   1710
         Top             =   810
         Width           =   3975
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
         Left            =   5850
         TabIndex        =   41
         Top             =   855
         Width           =   825
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
         Left            =   405
         TabIndex        =   40
         Top             =   810
         Width           =   840
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   13
         Left            =   1710
         Top             =   4320
         Width           =   1095
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
         Left            =   405
         TabIndex        =   39
         Top             =   4365
         Width           =   1080
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Left            =   9585
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Leave"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   8955
         TabIndex        =   38
         Top             =   1350
         Width           =   450
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Working Day"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4455
         TabIndex        =   37
         Top             =   1305
         Width           =   930
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   22
         Left            =   5850
         Top             =   1260
         Width           =   555
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   3060
         Top             =   330
         Width           =   465
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Other(-)"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3240
         TabIndex        =   36
         Top             =   3720
         Width           =   525
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   23
         Left            =   4590
         Top             =   3720
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   5415
      Picture         =   "frmBonusPreparation.frx":0994
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5835
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   6690
      Picture         =   "frmBonusPreparation.frx":239E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5835
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   2850
      Picture         =   "frmBonusPreparation.frx":3F88
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5835
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   4125
      Picture         =   "frmBonusPreparation.frx":591A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5835
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   7965
      Picture         =   "frmBonusPreparation.frx":72AC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5835
      Width           =   1185
   End
   Begin VB.ListBox lstTips 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3FEFF&
      ForeColor       =   &H000000C0&
      Height          =   225
      ItemData        =   "frmBonusPreparation.frx":8D2E
      Left            =   180
      List            =   "frmBonusPreparation.frx":8D30
      TabIndex        =   5
      Top             =   5805
      Visible         =   0   'False
      Width           =   1050
   End
End
Attribute VB_Name = "frmBonusPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Advance_Info As New clsAdvance_Info
Private objSalary As New Salary_Monthly
Private Salaray_Pre As New Cls_salary_Preparation
Dim conn1 As New Connection
Dim rs1 As New Recordset
Dim conn2 As New Connection
Dim RS2 As New Recordset
Dim conn3 As New Connection
Dim rs3 As New Recordset
Dim conn4 As New Connection
Dim rs4 As New Recordset
Dim conn5 As New Connection
Dim rs5 As New Recordset
Dim conn6 As New Connection
Dim conn007 As New Connection
Dim RS007 As New Recordset
Dim rs005 As New Recordset
Dim BasicSalaryOfEmp As Integer
Private Sub cboMonth_Click()
On Error GoTo Errdesc
    If Combo1(1) = "" Then Exit Sub
    Get_All_Information
    Get_Value_Into_Grid
      Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub cboYear_Click()
If Combo1(1) = "" Then Exit Sub
    Flash_Data
End Sub
Private Sub cmdClear_Click()
    Clear_Screen
    StatusBar1.Panels(1) = ""
    Combo1(1).SetFocus
End Sub
Private Sub cmdClose_Click()
    Close_Msg Me
End Sub
Private Sub cmdDelete_Click()
On Error GoTo Errdesc
BonusPreparationStatus = 1
With objSalary
    .Connstring = strCN.Connection_String
    .Emp_ID = Combo1(1)
    .Salay_Month_Get = Trim(cboMonth)
    .Salary_YearGet = cboYear
    .Delete
    MsgBox "Data Deleted Successfully", vbInformation, "IT Division, DNMIH"
    
    Clear_Screen
    
    StatusBar1.Panels(1) = ""
    Combo1(1).SetFocus
End With
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub cmdPreview_Click()

End Sub

Private Sub cmdSave_Click()
On Error GoTo Errdesc
    BonusPreparationStatus = 1

    With Salaray_Pre
        .Connstring = strCN.Connection_String
        .Emp_ID = Combo1(1)
        .DEPT_NM = txtDept
        .designation = txtDesig
        .Emp_Nm = txtName
        .PAY_MONTH = cboMonth
        .PAY_YEAR = cboYear
        .BASIC = txtBasic
        .H_RENT = txtHR
        .MED = txtMed
        .CONV = txtConv
        .TFN = txtTiffin
        .DA = txtDA
        .ATTN = txtAttn
        .LEAVE = Text1
        .ARREAR = txtArrear
        .OTHERS_ADDITION = txtOthers_Add
        .PF_CONTRI_DEDUCTION = txtPF_Contribution
        .PF_LN_AMOUNT = txtPF_Loan
        .SALARY_ADVANCE = txtAdvance
        .OTHERS_DEDUCTION = txtOthers_Deduct
        .R_STAMP = lblRev_Stamp
        .NET_PAYABLE = txtNet_Payable
        .CREATE_BY = "Dsl"
        .CREATE_DATE = Date$
        .UPDATE_DATE = Date$
        .Remarks = txtRemarks
        .WORKING_DAY = Text2
        .SALARY_DISBURSE = 0
        .adddeduct_other = Text3
        .Save
        MsgBox "Bonus Has been preparared for this Employee", vbInformation, "IT Division, DNMIH"
    End With

Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub
Private Sub CommandButton1_Click()
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
Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Errdesc
   If KeyCode = 13 Then
        Get_All_Information
    End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
    If (Shift And vbCtrlMask) > 0 Then
        Select Case KeyCode
           
            
            Case vbKeyA: txtArrear.SetFocus
            Case vbKeyT: txtTelephone.SetFocus
            Case vbKeyO: txtOthers_Add.SetFocus
            Case vbKeyH: txtOthers_Deduct.SetFocus
            Case vbKeyR: txtRemarks.SetFocus
            Case vbKeyS: cmdSave.SetFocus
            Case vbKeyI: Combo1(1).SetFocus
            Case vbKeyI: txtAttn.SetFocus
            
        End Select
    End If
    
End Sub
Private Sub Form_Load()
On Error GoTo Errdes
    Screen_Position Me

    Load_Yr Me
    Load_MonthNm Me
    
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
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
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

Private Sub lblRev_Stamp_Change()
 Default_Zero txtOthers_Deduct
    Calculate
End Sub

Private Sub lblRev_Stamp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdSave.SetFocus
End If
End Sub

Private Sub Text2_Change()
 Default_Zero txtAdvance
 Calculate
End Sub
Private Sub txtAdvance_Change()
    Default_Zero txtAdvance
    Calculate
End Sub
Private Sub txtArrear_Change()
    Default_Zero txtArrear
    Calculate
End Sub
Private Sub txtAttn_Change()
If Len(Trim(txtAttn)) = 0 Then txtAttn = 0
Text1 = Val(Text2) - Val(txtAttn)
End Sub

Private Sub txtConv_Change()
 Default_Zero txtOthers_Add
    Calculate
End Sub
Private Sub txtDA_Change()
     Default_Zero txtOthers_Add
    Calculate
End Sub
Private Sub txtElec_Bill_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
Public Sub Flash_Data()
On Error GoTo Errdes
        Dim EmpID As String
        Dim PMonth As String
        Dim PYear As String
        
        
        
        EmpID = Combo1(1)
        PMonth = cboMonth
        PYear = cboYear
        
        Clear_Screen
        
        Combo1(1) = EmpID
        cboMonth = PMonth
        cboYear = PYear
        
    With objSalary
        .Connstring = strCN.Connection_String
        .Emp_ID = Combo1(1)
        .PAY_MONTH = Get_Month_No(cboMonth)
        .PAY_YEAR = cboYear
        .GetX
           
        txtName = .Emp_Nm
        txtDesig = .designation
        txtDept = .Department
                
        txtBasic = .BASIC
        txtHR = .H_RENT
        txtConv = .CONV
        txtTiffin = .TFN
        txtTelephone = .TELEPHONE
        
        txtMed = .MED
        txtPF_Contribution = Round(.PF_Ded)
        txtPF_Loan = .Ln_Amount
        txtAdvance = .ADV_AMOUNT
        txtArrear = .ARREAR
        txtDA = .DA
        
        txtOthers_Add = .OTHERS_ALLOWANCE
        txtOthers_Deduct = .Others_Ded
        lblRev_Stamp = .R_STAMP
        
        txtAttn = .ATTN
        txtRemarks = .Remarks
                
    End With
        
  'Calculate
Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Public Sub Calculate()
On Error GoTo Errdesc
    txtTotal = 0 + Val(txtHR) + Val(txtConv) + Val(txtMed) _
            + Val(txtDA) + Val(txtTiffin) + Val(txtOthers_Add) _
            + Val(txtArrear) + Val(txtTelephone)



    txtTotal_Deduction = Val(txtPF_Contribution) + Val(txtAdvance) _
            + Val(lblRev_Stamp) + Val(txtOthers_Deduct) + Val(txtPF_Loan) + Val(Text2.Text)
            
    txtNet_Payable = Val(txtTotal) - Val(txtTotal_Deduction)

Exit Sub

Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub txtHR_Change()
 Default_Zero txtOthers_Add
    Calculate
End Sub

Private Sub txtMed_Change()
 Default_Zero txtOthers_Add
    Calculate
End Sub

Private Sub txtOthers_Deduct_Change()
    Default_Zero txtOthers_Deduct
    Calculate
End Sub

Private Sub txtOthers_Deduct_KeyPress(KeyAscii As Integer)
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
Private Sub txtPF_Loan_Change()
    Default_Zero txtPF_Loan
    Calculate
End Sub
Private Sub txtTelephone_Change()
    Default_Zero txtTelephone
    Calculate
End Sub
Private Sub Get_All_Information()
On Error GoTo Errdesc
Dim cmd As New Command
conn1.ConnectionString = strCN.Connection_String
conn1.Open
cmd.ActiveConnection = conn1
cmd.CommandType = adCmdText
cmd.CommandText = "SELECT ST_DEPT.DEPT_NM,EMP_JOB_INFO.SCALE_CODE,EMP_JOB_INFO.BASIC_SAL, " + _
                " St_Desig.DESIGNATION , emp_info.Emp_Nm,EMP_JOB_INFO.MODE_OF_PAYMENT,emp_job_info.PF_MEM " + _
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
        
        txtBasic = BasicSalaryOfEmp
        
               
        txtTelephone = BasicSalaryOfEmp

        txtHR = 0
        txtMed = 0
        txtConv = 0
        txtTiffin = 0
        txtDA = 0
        txtAttn = 0
        Text1 = 0
        txtArrear = 0
        txtOthers_Add = 0
        txtPF_Contribution = 0
'        txtPF_Loan = 0
        txtAdvance = 0
        txtOthers_Deduct = 0
        lblRev_Stamp = 0
        txtNet_Payable = 0
        txtRemarks = ""
        Text2 = 0
        Text3 = 0
        rs1.Close
        conn1.Close
        txtTelephone.SetFocus
                

    Else
        MsgBox "Invalid Employee No.", vbInformation, "Warning:IT Division, DNMIH"
        rs1.Close
        conn1.Close
        Exit Sub
    End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub txtTelephone_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    lblRev_Stamp.SetFocus
End If
End Sub

Private Sub txtTiffin_Change()
 Default_Zero txtOthers_Add
 Calculate
End Sub
Private Sub Get_Value_Into_Grid()
On Error GoTo Errdesc
Dim cmd As New Command
Dim conn005 As New Connection


conn005.ConnectionString = strCN.Connection_String
conn005.Open
cmd.ActiveConnection = conn005
cmd.CommandType = adCmdText

cmd.CommandText = "select *  from BONUS_PREAPARATION where  emp_id='" & Combo1(1) & "' AND PAY_MONTH='" & cboMonth & "' AND PAY_YEAR='" & cboYear & "'"
rs005.CursorLocation = adUseClient
rs005.Open cmd.CommandText, conn005, adOpenDynamic, adLockOptimistic

If rs005.RecordCount > 0 Then
  
    
    If Not IsNull(rs005.Fields(0)) Then
         StatusBar1.Panels(1) = "Bonus has been Preapared for " & rs005.Fields(1) & "  " & "In this Month"
         
        txtTelephone = "" & rs005.Fields("TELEPHONE")
        lblRev_Stamp = rs005.Fields("R_STAMP")
        txtNet_Payable = rs005.Fields("NET_PAYABLE")
        txtRemarks = "" & rs005.Fields("Remarks")

         
    Else
         StatusBar1.Panels(1) = ""
    End If
    
    
End If

    rs005.Close
    conn005.Close

   
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

