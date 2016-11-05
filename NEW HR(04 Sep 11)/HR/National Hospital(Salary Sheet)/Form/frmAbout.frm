VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ST_Leave 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Leave Setup"
   ClientHeight    =   5145
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   8175
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3551.171
   ScaleMode       =   0  'User
   ScaleWidth      =   7676.749
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   3465
      Picture         =   "frmAbout.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4545
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   4755
      Picture         =   "frmAbout.frx":22D4
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4545
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   810
      Picture         =   "frmAbout.frx":3EBE
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4545
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   2130
      Picture         =   "frmAbout.frx":5850
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4545
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   6075
      Picture         =   "frmAbout.frx":71E2
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4545
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   4335
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   7890
      Begin VB.TextBox txtleave_name 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   3960
         TabIndex        =   11
         Top             =   360
         Width           =   3660
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000E&
         Caption         =   "No"
         ForeColor       =   &H8000000D&
         Height          =   330
         Index           =   1
         Left            =   4815
         TabIndex        =   10
         Top             =   810
         Width           =   510
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000E&
         Caption         =   "Yes"
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   1
         Left            =   3960
         TabIndex        =   9
         Top             =   855
         Width           =   735
      End
      Begin VB.TextBox txtcarry_max_days 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   6615
         TabIndex        =   7
         Top             =   765
         Width           =   1005
      End
      Begin VB.TextBox txtdays 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1305
         TabIndex        =   5
         Top             =   765
         Width           =   1050
      End
      Begin VB.TextBox txtleave_code 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1305
         TabIndex        =   2
         Top             =   360
         Width           =   1050
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2805
         Left            =   225
         TabIndex        =   17
         Top             =   1260
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   4948
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
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Carry Arrear"
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   1
         Left            =   2520
         TabIndex        =   8
         Top             =   855
         Width           =   1365
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Carry Days"
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   1
         Left            =   5580
         TabIndex        =   6
         Top             =   855
         Width           =   1365
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   " Leave Days"
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   855
         Width           =   915
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Leave Name"
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   3
         Top             =   450
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Leave Code"
         ForeColor       =   &H8000000D&
         Height          =   330
         Index           =   1
         Left            =   225
         TabIndex        =   1
         Top             =   405
         Width           =   1185
      End
   End
End
Attribute VB_Name = "ST_Leave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
