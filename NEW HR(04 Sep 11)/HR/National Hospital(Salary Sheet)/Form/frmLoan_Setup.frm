VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " PF & Others Loan Setup "
   ClientHeight    =   5190
   ClientLeft      =   1665
   ClientTop       =   2115
   ClientWidth     =   8400
   Icon            =   "frmLoan_Setup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form40"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   6165
      Picture         =   "frmLoan_Setup.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4500
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   2220
      Picture         =   "frmLoan_Setup.frx":234C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4500
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   900
      Picture         =   "frmLoan_Setup.frx":3CDE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4500
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   4845
      Picture         =   "frmLoan_Setup.frx":5670
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4500
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   3555
      Picture         =   "frmLoan_Setup.frx":725A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4500
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4200
      Left            =   135
      TabIndex        =   4
      Top             =   135
      Width           =   8115
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1860
         Left            =   360
         TabIndex        =   19
         Top             =   1935
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   3281
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
      Begin VB.TextBox txtInt_Rate 
         Height          =   285
         Left            =   3690
         TabIndex        =   18
         Top             =   1260
         Width           =   915
      End
      Begin VB.TextBox txtPercent 
         Height          =   285
         Left            =   3690
         TabIndex        =   17
         Top             =   810
         Width           =   915
      End
      Begin VB.TextBox txtLn_Nm 
         Height          =   330
         Left            =   3690
         TabIndex        =   16
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtInstallment 
         Height          =   285
         Left            =   1260
         TabIndex        =   15
         Top             =   1125
         Width           =   1185
      End
      Begin VB.TextBox txtCeiling 
         Height          =   285
         Left            =   1260
         TabIndex        =   14
         Top             =   765
         Width           =   1185
      End
      Begin VB.TextBox txtLnCode 
         Height          =   285
         Left            =   1260
         TabIndex        =   13
         Top             =   405
         Width           =   1185
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "No. of  Installment"
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
         Height          =   435
         Left            =   315
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Taka      Or "
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
         Left            =   2730
         TabIndex        =   11
         Top             =   810
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%  of own contribution (only for PF Loan)"
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
         Left            =   4770
         TabIndex        =   9
         Top             =   810
         Width           =   3000
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Name"
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
         Left            =   2730
         TabIndex        =   8
         Top             =   405
         Width           =   810
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate of Interest"
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
         Height          =   435
         Left            =   2745
         TabIndex        =   7
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ceiling"
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
         Left            =   315
         TabIndex        =   6
         Top             =   810
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Code"
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
         Left            =   315
         TabIndex        =   5
         Top             =   390
         Width           =   780
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Loan As New St_Loan
Private Loan_Rs As New Recordset
Dim Track_Id As Long

Private Sub cmdClear_Click()
    Clear_Screen
    txtLnCode.SetFocus
End Sub

Private Sub cmdClose_Click()
    Close_Msg Me
End Sub

Private Sub cmdDelete_Click()
    Delete_Record "20", Track_Id
    Flash_Into_Grid
    Clear_Screen
    txtLnCode.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

   On Error Resume Next
    Screen_Position Me
    Track_Id = 0
    Flash_Into_Grid
   
End Sub
Private Sub cmdSave_Click()

    With Loan
        .ConnString = strCN.Connection
        .Ln_Code = txtLnCode
        .Ln_Nm = txtLn_Nm
        .Ceil_Amt = txtCeiling
        .Amt_Prcnt = txtPercent
        .Inst_No = txtInstallment
        .Int_Rate = txtInt_Rate
        .Track_Id = Track_Id
        .Save
    End With
    
    Flash_Into_Grid

    Clear_Screen
    Track_Id = 0
    txtLnCode.SetFocus
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
    txtLnCode = Loan_Rs!Ln_Code
    txtLn_Nm = Loan_Rs!Ln_Nm
    txtCeiling = Loan_Rs!Ceil_Amt
    txtPercent = Loan_Rs!Amt_Prcnt
    txtInstallment = Loan_Rs!Inst_No
    txtInt_Rate = Loan_Rs!Int_Rate
    Track_Id = Loan_Rs!Track_Id
txtLn_Nm.SetFocus

End Sub

Public Sub Flash_Into_Grid()
On Error Resume Next

    With Loan
        .ConnString = strCN.Connection
        Set Loan_Rs = .GetAll
    End With
    
     Set DataGrid1.DataSource = Loan_Rs
                    
        With DataGrid1
            .Columns(0).Width = 1125
            '.Columns(0).DataField = Loan_Rs!Fields(0)
            
            .Columns(1).Width = 2860
            '.Columns(1).DataField = Loan_Rs!Fields(1)
            
            .Columns(2).Width = 1065
            '.Columns(2).DataField = Loan_Rs!Fields(2)
            
            .Columns(3).Width = 720
            '.Columns(3).DataField = Loan_Rs!Fields(3)
            
            .Columns(4).Width = 720
            '.Columns(4).DataField = Loan_Rs!Fields(2)
            
            .Columns(5).Width = 720
            '.Columns(5).DataField = Loan_Rs!Fields(3)

        End With
       
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Destroy Me
End Sub

Private Sub txtCeiling_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtInstallment_KeyPress(KeyAscii As MSForms.ReturnInteger)
     KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtInt_Rate_KeyPress(KeyAscii As MSForms.ReturnInteger)
     KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtPercent_Change()
    Default_Zero txtPercent
End Sub

Public Sub Default_Zero(txt As MSForms.TextBox)
    
    If Len(txt) < 1 Then txt = 0
    
End Sub

Private Sub txtPercent_KeyPress(KeyAscii As MSForms.ReturnInteger)
     KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
