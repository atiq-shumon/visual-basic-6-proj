VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Production and Profit Information                              Carew & Company (Bangladesh) Limited"
   ClientHeight    =   5220
   ClientLeft      =   1980
   ClientTop       =   1665
   ClientWidth     =   8310
   Icon            =   "frmProduction_Info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   6075
      Picture         =   "frmProduction_Info.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4545
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   2130
      Picture         =   "frmProduction_Info.frx":234C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4545
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   810
      Picture         =   "frmProduction_Info.frx":3CDE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4545
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   4755
      Picture         =   "frmProduction_Info.frx":5670
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4545
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   3465
      Picture         =   "frmProduction_Info.frx":725A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4545
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4290
      Left            =   135
      TabIndex        =   8
      Top             =   135
      Width           =   8025
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2760
         Left            =   315
         TabIndex        =   17
         Top             =   1260
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   4868
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BorderStyle     =   0
         ColumnHeaders   =   0   'False
         ForeColor       =   10485760
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   6
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
         ColumnCount     =   4
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
               ColumnWidth     =   1725.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1755.213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1814.74
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin MSForms.TextBox txtAct_Prod 
         Height          =   330
         Left            =   5760
         TabIndex        =   3
         Top             =   765
         Width           =   1005
         VariousPropertyBits=   746604571
         ForeColor       =   8388608
         BorderStyle     =   1
         Size            =   "1773;582"
         BorderColor     =   16761024
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAmount 
         Height          =   330
         Left            =   5760
         TabIndex        =   1
         Top             =   315
         Width           =   1005
         VariousPropertyBits=   746604571
         ForeColor       =   8388608
         BorderStyle     =   1
         Size            =   "1773;582"
         BorderColor     =   16761024
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboFiscalYr 
         Height          =   330
         Left            =   1935
         TabIndex        =   0
         Top             =   315
         Width           =   1320
         VariousPropertyBits=   746604571
         ForeColor       =   12582912
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2328;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         BorderColor     =   16761024
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtExp_Prod 
         Height          =   285
         Left            =   1935
         TabIndex        =   2
         Top             =   765
         Width           =   1305
         VariousPropertyBits=   746604571
         ForeColor       =   8388608
         BorderStyle     =   1
         Size            =   "2302;503"
         BorderColor     =   16761024
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lacs Tk."
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
         Index           =   2
         Left            =   6975
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M.Ton"
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
         Index           =   0
         Left            =   6975
         TabIndex        =   15
         Top             =   810
         Width           =   435
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Production"
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
         Index           =   1
         Left            =   4365
         TabIndex        =   14
         Top             =   810
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M.Ton"
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
         Index           =   1
         Left            =   3375
         TabIndex        =   13
         Top             =   810
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fiscal Year"
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
         TabIndex        =   12
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Expected Production"
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
         Height          =   255
         Index           =   0
         Left            =   315
         TabIndex        =   11
         Top             =   810
         Width           =   1665
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Profit Amount"
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
         Left            =   4365
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   2850
         Index           =   6
         Left            =   270
         Top             =   1215
         Width           =   7485
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Prod As New Profit_Prod_Info
Private Prod_Rs As New Recordset
Dim Track_Id As Long

Private Sub cmdClear_Click()
    Clear_Screen
End Sub

Private Sub cmdClose_Click()
    Close_Msg Me
End Sub

Private Sub cmdDelete_Click()
    Delete_Record "3", Track_Id
    Flash_Into_Grid
    Clear_Screen
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Screen_Position Me
    
    Load_FiscalYr Me
    Track_Id = 0
    Flash_Into_Grid
   
End Sub
Private Sub cmdSave_Click()

    With Prod
        .ConnString = strCN.Connection
        .Fscl_year = cboFiscalYr
        .Exp_Prod = txtExp_Prod
        .Act_Prod = txtAct_Prod
        .Profit_Amt = txtAmount
        .Track_Id = Track_Id
        .Save
    End With
    
    Flash_Into_Grid
    Clear_Screen
    Track_Id = 0
    
End Sub

Private Sub DataGrid1_Click()

    cboFiscalYr = Prod_Rs!Fscl_year
    txtExp_Prod = Prod_Rs!Exp_Prod
    txtAct_Prod = Prod_Rs!Act_Prod
    txtAmount = Prod_Rs!Profit_Amt
    
    Track_Id = Prod_Rs!Track_Id
End Sub

Public Sub Flash_Into_Grid()

    With Prod
        .ConnString = strCN.Connection
        Set Prod_Rs = .GetAll
    End With
    
     Set DataGrid1.DataSource = Prod_Rs
                    
        With DataGrid1
            .Columns(0).Width = 1775
            '.Columns(0).DataField = Prod_Rs!Fields(0)

            .Columns(1).Width = 1775
            '.Columns(1).DataField = Prod_Rs!Fields(1)

            .Columns(2).Width = 1775
            '.Columns(2).DataField = Prod_Rs!Fields(2)

            .Columns(3).Width = 1775
            '.Columns(3).DataField = Prod_Rs!Fields(3)


        End With
       
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Destroy Me
End Sub

Private Sub txtAct_Prod_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only

End Sub

Private Sub txtAmount_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only

End Sub

Private Sub txtExp_Prod_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
