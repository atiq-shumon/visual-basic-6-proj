VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form15 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Bonus Setup"
   ClientHeight    =   5265
   ClientLeft      =   2385
   ClientTop       =   1755
   ClientWidth     =   6705
   ForeColor       =   &H00800000&
   Icon            =   "frmBonus_Setup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form13"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6705
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   315
      TabIndex        =   5
      Top             =   405
      Width           =   6225
      Begin VB.TextBox txtRate 
         Height          =   465
         Left            =   4005
         TabIndex        =   12
         Top             =   1395
         Width           =   1905
      End
      Begin VB.TextBox txtBonusNm 
         Height          =   465
         Left            =   2025
         TabIndex        =   11
         Top             =   1395
         Width           =   1725
      End
      Begin VB.TextBox txtBonusCode 
         Height          =   465
         Left            =   45
         TabIndex        =   10
         Top             =   1395
         Width           =   1725
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2130
         Left            =   90
         TabIndex        =   6
         Top             =   1935
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   3757
         _Version        =   393216
         ForeColor       =   8388608
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
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Festival Bonus (Eid Bonus, Puja Bonus,X-mas Bonus etc. )"
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   270
         TabIndex        =   13
         Top             =   495
         Width           =   4785
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bonus Rate"
         ForeColor       =   &H00C00000&
         Height          =   510
         Left            =   4095
         TabIndex        =   9
         Top             =   990
         Width           =   1815
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bonus Name"
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   2025
         TabIndex        =   8
         Top             =   1080
         Width           =   1725
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bonus Code"
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   1125
         Width           =   1680
      End
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   2790
      Picture         =   "frmBonus_Setup.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4635
      Width           =   1140
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   5355
      Picture         =   "frmBonus_Setup.frx":22D4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4635
      Width           =   1140
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   1500
      Picture         =   "frmBonus_Setup.frx":3D56
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4635
      Width           =   1140
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   225
      Picture         =   "frmBonus_Setup.frx":56E8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4635
      Width           =   1140
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   4035
      Picture         =   "frmBonus_Setup.frx":707A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4635
      Width           =   1185
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BnsFs As New St_BnsFs
Private BnsFs_Rs As New Recordset

Private BnsPd As New St_BnsPd
Private BnsPd_Rs As New Recordset

Private BnsPr As New St_BnsPr
Private BnsPr_Rs As New Recordset

Dim SSTab_Index As Integer

Dim Track_Id As Long

Private Sub cmdClear_Click()
    Clear_Screen
    txtBonusCode(SSTab_Index).SetFocus
End Sub

Private Sub cmdClose_Click()
    Close_Msg Me
End Sub

Private Sub cmdDelete_Click()

 Select Case SSTab_Index
    Case 0
            Delete_Record "17", Track_Id
    Case 1
            Delete_Record "18", Track_Id
    Case 2
            Delete_Record "19", Track_Id
 End Select

    Flash_Into_Grid
    Clear_Screen
    txtBonusCode(SSTab_Index).SetFocus
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
If SSTab_Index = 0 Then
    
    With BnsFs
        .ConnString = strCN.Connection
        .Bonus_Code = txtBonusCode(SSTab_Index)
        .Bonus_Nm = txtBonusNm
        .Rate = txtRate(SSTab_Index)
        .Track_Id = Track_Id
        .Save
    End With
    
ElseIf SSTab_Index = 1 Then

    With BnsPd
        .ConnString = strCN.Connection
        .Bonus_Code = txtBonusCode(SSTab_Index)
        .Prod_From = txtFrom(SSTab_Index)
        .Prod_To = txtTo(SSTab_Index)
        .Rate = txtRate(SSTab_Index)
        .Track_Id = Track_Id
        .Save
    End With
    
Else

    With BnsPr
        .ConnString = strCN.Connection
        .Bonus_Code = txtBonusCode(SSTab_Index)
        .Prf_From = txtFrom(SSTab_Index)
        .Prf_To = txtTo(SSTab_Index)
        .Rate = txtRate(SSTab_Index)
        .Track_Id = Track_Id
        .Save
    End With


End If
    
    Flash_Into_Grid
    Clear_Screen
    Track_Id = 0
    txtBonusCode(SSTab_Index).SetFocus
End Sub

Private Sub DataGrid1_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then

    txtBonusCode(0) = BnsFs_Rs!Fs_Code
    txtBonusNm = BnsFs_Rs!Bonus_Nm
    txtRate(0) = BnsFs_Rs!Rate
    Track_Id = BnsFs_Rs!Track_Id
    txtBonusNm.SetFocus
ElseIf Index = 1 Then
    
    txtBonusCode(1) = BnsPd_Rs!Pd_Code
    txtFrom(1) = BnsPd_Rs!Prod_From
    txtTo(1) = BnsPd_Rs!Prod_To
    txtRate(1) = BnsPd_Rs!Rate
    Track_Id = BnsPd_Rs!Track_Id
    txtFrom(1).SetFocus
Else

    txtBonusCode(2) = BnsPr_Rs!Pr_Code
    txtFrom(2) = BnsPr_Rs!Prf_From
    txtTo(2) = BnsPr_Rs!Prf_To
    txtRate(2) = BnsPr_Rs!Rate
    Track_Id = BnsPr_Rs!Track_Id
    txtFrom(2).SetFocus

End If

    
        
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    Screen_Position Me
    Track_Id = 0
    SSTab_Index = 0
    Flash_Into_Grid
    Set_TabIndex
End Sub

Public Sub Flash_Into_Grid()
On Error Resume Next

If SSTab_Index = 0 Then

    With BnsFs
        .ConnString = strCN.Connection
        Set BnsFs_Rs = .GetAll
    End With
    
     Set DataGrid1(0).DataSource = BnsFs_Rs
         DataGrid1(0).Refresh

        With DataGrid1(0)
            .Columns(0).Width = 720
            '.Columns(0).DataField = Desig_Rs!Fields(0)

            .Columns(1).Width = 3465
            '.Columns(1).DataField = Desig_Rs!Fields(1)

            .Columns(2).Width = 1215
            '.Columns(2).DataField = Desig_Rs!Fields(2)

        End With

ElseIf SSTab_Index = 1 Then

     With BnsPd
        .ConnString = strCN.Connection
        Set BnsPd_Rs = .GetAll
    End With

    Set DataGrid1(1).DataSource = BnsPd_Rs
         DataGrid1(1).Refresh
        
        With DataGrid1(1)
            .Columns(0).Width = 1215
            '.Columns(0).DataField = Desig_Rs!Fields(0)

            .Columns(1).Width = 1380
            '.Columns(1).DataField = Desig_Rs!Fields(1)

            .Columns(2).Width = 1490
            '.Columns(2).DataField = Desig_Rs!Fields(2)

            .Columns(3).Width = 1350
            '.Columns(3).DataField = Desig_Rs!Fields(3)
        End With

Else

    With BnsPr
        .ConnString = strCN.Connection
        Set BnsPr_Rs = .GetAll
    End With
    
    Set DataGrid1(2).DataSource = BnsPr_Rs
         DataGrid1(2).Refresh
        
        With DataGrid1(2)
            .Columns(0).Width = 1215
            '.Columns(0).DataField = Desig_Rs!Fields(0)
    
            .Columns(1).Width = 1380
            '.Columns(1).DataField = Desig_Rs!Fields(1)
    
            .Columns(2).Width = 1490
            '.Columns(2).DataField = Desig_Rs!Fields(2)
    
            .Columns(3).Width = 1350
            '.Columns(3).DataField = Desig_Rs!Fields(3)
        End With




End If
       
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    SSTab_Index = SSTab1.Tab
    
    Flash_Into_Grid
    Set_TabIndex
End Sub




Public Sub Set_TabIndex()
On Error Resume Next

    Select Case SSTab_Index
    
        Case 0
            txtBonusCode(SSTab_Index).TabIndex = 0
            txtBonusNm.TabIndex = 1
            txtRate(SSTab_Index).TabIndex = 2
            cmdSave.TabIndex = 3
            cmdClear.TabIndex = 4
            cmdClose.TabIndex = 5
            txtBonusCode(SSTab_Index).SetFocus
        Case 1, 2
            txtBonusCode(SSTab_Index).TabIndex = 0
            txtFrom(SSTab_Index).TabIndex = 1
            txtTo(SSTab_Index).TabIndex = 2
            txtRate(SSTab_Index).TabIndex = 3
            cmdSave.TabIndex = 4
            cmdClear.TabIndex = 5
            cmdClose.TabIndex = 6
            txtBonusCode(SSTab_Index).SetFocus
    End Select
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Destroy Me
End Sub

Private Sub txtFrom_KeyPress(Index As Integer, KeyAscii As Integer)
      KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtRate_KeyPress(Index As Integer, KeyAscii As Integer)
      KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub


Private Sub txtTo_KeyPress(Index As Integer, KeyAscii As Integer)
      KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
