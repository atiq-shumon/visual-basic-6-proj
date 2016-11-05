VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form21 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheque Status"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "rpt_cheque_status.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtst_dt 
      Height          =   285
      Left            =   180
      TabIndex        =   7
      Top             =   1140
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      Format          =   56623105
      CurrentDate     =   38518
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   855
      Left            =   -30
      TabIndex        =   5
      Top             =   -120
      Width           =   5985
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque Status Entry"
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
         Left            =   2040
         TabIndex        =   6
         Top             =   150
         Width           =   3540
      End
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1035
      TabIndex        =   4
      Top             =   1605
      Visible         =   0   'False
      Width           =   1230
   End
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
      Left            =   4740
      Picture         =   "rpt_cheque_status.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Preview"
      Top             =   1650
      Width           =   510
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
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
      Left            =   5280
      Picture         =   "rpt_cheque_status.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exit"
      Top             =   1635
      Width           =   510
   End
   Begin VB.TextBox txtQuery 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2295
      TabIndex        =   3
      Top             =   1605
      Visible         =   0   'False
      Width           =   1230
   End
   Begin MSComCtl2.DTPicker dted_dt 
      Height          =   285
      Left            =   3855
      TabIndex        =   0
      Top             =   1140
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   503
      _Version        =   393216
      Format          =   56623105
      CurrentDate     =   36949
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
      Left            =   4290
      TabIndex        =   10
      Top             =   870
      Width           =   270
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
      Left            =   360
      TabIndex        =   9
      Top             =   870
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Range"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2490
      TabIndex        =   8
      Top             =   1170
      Width           =   945
   End
   Begin VB.Shape Shape1 
      Height          =   510
      Index           =   3
      Left            =   4680
      Top             =   1605
      Width           =   1170
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdPREVIEW_Click()
    rptMode = 14
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

