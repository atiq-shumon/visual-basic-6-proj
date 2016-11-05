VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form18 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Provident Fund Income Distribution"
   ClientHeight    =   5040
   ClientLeft      =   1110
   ClientTop       =   1710
   ClientWidth     =   9675
   Icon            =   "frmPF_Income_Distribution.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9675
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   135
      TabIndex        =   4
      Top             =   90
      Width           =   9420
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   8640
         TabIndex        =   13
         Top             =   450
         Width           =   645
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   6480
         TabIndex        =   12
         Top             =   450
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   4005
         TabIndex        =   11
         Top             =   495
         Width           =   1500
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   495
         Width           =   2040
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2490
         Left            =   225
         TabIndex        =   5
         Top             =   1215
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   4392
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   "Name"
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
            Caption         =   "Designation"
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
            Caption         =   "Unit"
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
            Caption         =   "Cost Centre"
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
         BeginProperty Column04 
            DataField       =   ""
            Caption         =   "Unit"
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
               ColumnWidth     =   1844.787
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2729.764
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   900.284
            EndProperty
         EndProperty
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fiscal Year"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   9
         Top             =   495
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Fund"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   3195
         TabIndex        =   8
         Top             =   495
         Width           =   765
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Income"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   5535
         TabIndex        =   7
         Top             =   495
         Width           =   930
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Income %"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   7920
         TabIndex        =   6
         Top             =   495
         Width           =   690
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   480
      Index           =   4
      Left            =   6390
      Picture         =   "frmPF_Income_Distribution.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Height          =   480
      Index           =   2
      Left            =   3645
      Picture         =   "frmPF_Income_Distribution.frx":234C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Height          =   480
      Index           =   0
      Left            =   2160
      Picture         =   "frmPF_Income_Distribution.frx":3D56
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Height          =   480
      Index           =   1
      Left            =   5040
      Picture         =   "frmPF_Income_Distribution.frx":5A78
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   1230
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_Click()
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

