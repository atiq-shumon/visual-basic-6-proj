VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form29 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BalanceSheet"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "Balancesheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   825
      Left            =   -30
      TabIndex        =   2
      Top             =   -120
      Width           =   5985
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balancesheet "
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   3480
         TabIndex        =   3
         Top             =   150
         Width           =   2385
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1005
      Left            =   -60
      TabIndex        =   5
      Top             =   570
      Width           =   6075
      Begin MSComCtl2.DTPicker dtst_dt 
         Height          =   285
         Left            =   210
         TabIndex        =   6
         Top             =   450
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Format          =   22806529
         CurrentDate     =   38518
      End
      Begin MSComCtl2.DTPicker dted_dt 
         Height          =   285
         Left            =   3885
         TabIndex        =   7
         Top             =   450
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   503
         _Version        =   393216
         Format          =   22806529
         CurrentDate     =   36949
      End
      Begin VB.Shape Shape3 
         Height          =   345
         Left            =   3840
         Top             =   420
         Width           =   2055
      End
      Begin VB.Shape Shape2 
         Height          =   345
         Left            =   180
         Top             =   420
         Width           =   1995
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--Date Range--"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   225
         Left            =   2460
         TabIndex        =   10
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   390
         TabIndex        =   9
         Top             =   180
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4320
         TabIndex        =   8
         Top             =   180
         Width           =   270
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   855
      Left            =   -30
      TabIndex        =   4
      Top             =   1440
      Width           =   6045
      Begin VB.CommandButton cmdPREVIEW 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4830
         Picture         =   "Balancesheet.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Preview"
         Top             =   240
         Width           =   510
      End
      Begin VB.CommandButton cmdClose 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5370
         Picture         =   "Balancesheet.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Exit"
         Top             =   225
         Width           =   510
      End
      Begin VB.Shape Shape1 
         Height          =   510
         Index           =   4
         Left            =   4770
         Top             =   180
         Width           =   1170
      End
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1035
      TabIndex        =   1
      Top             =   705
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox txtQuery 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2295
      TabIndex        =   0
      Top             =   705
      Visible         =   0   'False
      Width           =   1230
   End
End
Attribute VB_Name = "Form29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub cmdPREVIEW_Click()
    rptMode = 26
    Screen.MousePointer = vbHourglass
'    Me.txtQuery.Text = "and a.vou_date<=''" & Format(Me.dted_dt.Value, "yyyy-mm-dd") & "''"
    Me.txtTitle = "Trail Balance as on  " & Me.dted_dt.Value
   
    CRViewer1.Show vbModal
    
    
End Sub

Private Sub dted_dt_CloseUp()

'    dted_dt.MaxDate = objectCompSetup.ed_dt
'    dted_dt.MinDate = objectCompSetup.st_dt
'
End Sub

Private Sub dted_dt_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
       SendKeys Chr(9)
    End If
    
End Sub

Private Sub dted_dt_LostFocus()

    dted_dt_CloseUp
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
    
End Sub

Private Sub Form_Load()
    
    rptMode = 10
'    objectCompSetup.Flush_Comp (strcn)
    dted_dt.Value = Date
    
End Sub

