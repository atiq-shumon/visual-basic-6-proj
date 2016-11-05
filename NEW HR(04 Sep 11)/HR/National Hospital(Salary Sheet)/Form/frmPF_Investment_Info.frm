VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form17 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Investment Information"
   ClientHeight    =   6030
   ClientLeft      =   1665
   ClientTop       =   2115
   ClientWidth     =   8400
   Icon            =   "frmPF_Investment_Info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form40"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8400
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   6165
      Picture         =   "frmPF_Investment_Info.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5355
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   2220
      Picture         =   "frmPF_Investment_Info.frx":234C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5355
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   900
      Picture         =   "frmPF_Investment_Info.frx":3CDE
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5355
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   4845
      Picture         =   "frmPF_Investment_Info.frx":5670
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5355
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   3555
      Picture         =   "frmPF_Investment_Info.frx":725A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5355
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   8115
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2535
         Left            =   315
         TabIndex        =   19
         Top             =   2250
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BorderStyle     =   0
         ColumnHeaders   =   0   'False
         ForeColor       =   12582912
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
         ColumnCount     =   6
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
         BeginProperty Column04 
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
         BeginProperty Column05 
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
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3330.142
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2204.788
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2280.189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2055.118
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2280.189
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtNotes 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   645
         Left            =   4995
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   1215
         Width           =   2805
      End
      Begin MSComCtl2.DTPicker dtpInvDt 
         Height          =   330
         Left            =   1710
         TabIndex        =   21
         Top             =   1170
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   12582912
         CalendarTrailingForeColor=   16576
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   48955393
         CurrentDate     =   37722
      End
      Begin MSComCtl2.DTPicker dtpMatDt 
         Height          =   330
         Left            =   1710
         TabIndex        =   22
         Top             =   1575
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   12582912
         CalendarTrailingForeColor=   16576
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   48955393
         CurrentDate     =   37722
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   735
         Index           =   1
         Left            =   4860
         Top             =   1170
         Width           =   2985
      End
      Begin MSForms.TextBox txtRate 
         Height          =   285
         Left            =   7335
         TabIndex        =   13
         Top             =   810
         Width           =   510
         VariousPropertyBits=   746604571
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "900;503"
         BorderColor     =   16761024
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAmount 
         Height          =   285
         Left            =   4860
         TabIndex        =   12
         Top             =   810
         Width           =   1140
         VariousPropertyBits=   746604571
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "2011;503"
         BorderColor     =   16761024
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtInvNm 
         Height          =   285
         Left            =   4860
         TabIndex        =   11
         Top             =   450
         Width           =   2985
         VariousPropertyBits=   746604571
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "5265;503"
         BorderColor     =   16761024
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtBankNm 
         Height          =   285
         Left            =   1710
         TabIndex        =   10
         Top             =   810
         Width           =   2265
         VariousPropertyBits=   746604571
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "3995;503"
         BorderColor     =   16761024
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtInvID 
         Height          =   285
         Left            =   1710
         TabIndex        =   9
         Top             =   450
         Width           =   1365
         VariousPropertyBits=   746604571
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "2408;503"
         BorderColor     =   16761024
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Investment ID"
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
         TabIndex        =   8
         Top             =   450
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
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
         Left            =   4140
         TabIndex        =   7
         Top             =   1215
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maturity Date"
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
         Left            =   300
         TabIndex        =   6
         Top             =   1605
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Left            =   4140
         TabIndex        =   5
         Top             =   810
         Width           =   555
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   2625
         Index           =   6
         Left            =   270
         Top             =   2205
         Width           =   7575
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Investment Name"
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
         Left            =   3450
         TabIndex        =   4
         Top             =   450
         Width           =   1365
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   6165
         TabIndex        =   3
         Top             =   810
         Width           =   1110
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
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
         Left            =   300
         TabIndex        =   2
         Top             =   810
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invesment Date"
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
         Left            =   300
         TabIndex        =   1
         Top             =   1200
         Width           =   1110
      End
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Invest As New Inv_Info
Private Invest_Rs As New Recordset
Dim Track_Id As Long

Private Sub cmdClear_Click()
    Clear_Screen
End Sub

Private Sub cmdClose_Click()
    Close_Msg Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
  
    Track_Id = 0
    Flash_Into_Grid
   
End Sub
Private Sub cmdSave_Click()

    With Invest
        .ConnString = strCN.Connection
        .Inv_ID = txtInvID
        .Inv_Dt = dtpInvDt
        .Mat_Dt = dtpMatDt
        .Inv_NM = txtInvNm
        .Amt = txtAmount
        .Bank_Nm = txtBankNm
        .Int_Rate = txtRate
        .Notes = txtNotes
        .Track_Id = Track_Id
        .Save
    End With
    
    Flash_Into_Grid
    
    Track_Id = 0
    
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next

    txtInvID = Invest_Rs!Inv_ID
    txtInvNm = Invest_Rs!Inv_NM
    dtpInvDt = Invest_Rs!Inv_Dt
    dtpMatDt = Invest_Rs!Mat_Dt
    txtAmount = Invest_Rs!Amt
    txtBankNm = Invest_Rs!Bank_Nm
    txtRate = Invest_Rs!Int_Rate
    txtNotes = Invest_Rs!Notes
    Track_Id = Invest_Rs!Track_Id
End Sub

Public Sub Flash_Into_Grid()

    With Invest
        .ConnString = strCN.Connection
        Set Invest_Rs = .GetAll
    End With
    
     Set DataGrid1.DataSource = Invest_Rs
                    
        With DataGrid1
            .Columns(0).Width = 1140
            '.Columns(0).DataField = Invest_Rs!Fields(0)

            .Columns(1).Width = 2430
            .Columns(1).DataField = Invest_Rs!Inv_Dt

            .Columns(2).Width = 2080
            '.Columns(2).DataField = Invest_Rs!Fields(2)

            .Columns(3).Width = 1050
            '.Columns(3).DataField = Invest_Rs!Fields(4)

            .Columns(4).Width = 650
            '.Columns(4).DataField = Invest_Rs!Fields(4)
            
             .Columns(5).Width = 1050
            '.Columns(5).DataField = Invest_Rs!Fields(5)

            .Columns(6).Width = 1050
            '.Columns(6).DataField = Invest_Rs!Fields(6)

        End With
       
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Destroy Me
End Sub
