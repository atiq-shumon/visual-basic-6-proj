VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "House Rent Allowance Setup"
   ClientHeight    =   5130
   ClientLeft      =   2445
   ClientTop       =   1920
   ClientWidth     =   7215
   Icon            =   "frmHouse_Rent.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   180
      TabIndex        =   11
      Top             =   135
      Width           =   6855
      Begin VB.TextBox txtSlubCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C000C0&
         Height          =   285
         Index           =   3
         Left            =   450
         TabIndex        =   0
         Top             =   525
         Width           =   1245
      End
      Begin VB.TextBox txtBasic_From 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C000C0&
         Height          =   285
         Index           =   3
         Left            =   1845
         TabIndex        =   1
         Top             =   525
         Width           =   1335
      End
      Begin VB.TextBox txtBasic_To 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C000C0&
         Height          =   285
         Index           =   3
         Left            =   3330
         TabIndex        =   2
         Top             =   525
         Width           =   1425
      End
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C000C0&
         Height          =   285
         Index           =   3
         Left            =   4815
         TabIndex        =   3
         Top             =   525
         Width           =   525
      End
      Begin VB.TextBox txtMinTk 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C000C0&
         Height          =   240
         Index           =   3
         Left            =   5535
         TabIndex        =   4
         Top             =   540
         Width           =   750
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2940
         Index           =   3
         Left            =   405
         TabIndex        =   12
         Top             =   855
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   5186
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
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
         ColumnCount     =   5
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1530.142
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   929.764
            EndProperty
         EndProperty
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate (%)"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   4770
         TabIndex        =   17
         Top             =   225
         Width           =   600
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Salary (to)"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   16
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Slab Code"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   450
         TabIndex        =   15
         Top             =   225
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   3030
         Index           =   23
         Left            =   360
         Top             =   810
         Width           =   6090
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Salary (from)"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   1755
         TabIndex        =   14
         Top             =   225
         Width           =   1305
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         Height          =   330
         Index           =   19
         Left            =   360
         Top             =   495
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         Height          =   330
         Index           =   20
         Left            =   1755
         Top             =   495
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         Height          =   330
         Index           =   21
         Left            =   3240
         Top             =   495
         Width           =   1545
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         Height          =   330
         Index           =   22
         Left            =   4770
         Top             =   495
         Width           =   690
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         Height          =   330
         Index           =   24
         Left            =   5445
         Top             =   495
         Width           =   1005
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum (Tk.)"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   5455
         TabIndex        =   13
         Top             =   225
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   5625
      Picture         =   "frmHouse_Rent.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4500
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   1680
      Picture         =   "frmHouse_Rent.frx":234C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4500
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   360
      Picture         =   "frmHouse_Rent.frx":3CDE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4500
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   4305
      Picture         =   "frmHouse_Rent.frx":5670
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4500
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   3015
      Picture         =   "frmHouse_Rent.frx":725A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4500
      Width           =   1185
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   855
      TabIndex        =   5
      Top             =   675
      Visible         =   0   'False
      Width           =   60
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private HR_Allow As New clsHouse_rent_allowance

Dim Value As String

Dim con As New Connection
Dim cmd As New Command
Dim RS As New Recordset

Private Sub cmdClear_Click()
   On Error Resume Next
    Clear_Screen
   ' txtSlubCode(SSTab_Index).SetFocus
End Sub

Private Sub cmdClose_Click()
    Close_Msg Me
End Sub

Private Sub cmdDelete_Click()
        Flash_Into_Grid
        Clear_Screen
       ' txtSlubCode(SSTab_Index).SetFocus
End Sub

Private Sub cmdSave_Click()

   On Error Resume Next

            With HR_Allow
                .Connstring = strCN.Connection_String
                .Slab_code = Me.txtSlubCode(3)
                .Basic_From = txtBasic_From(3)
                .Basic_To = txtBasic_To(3)
                .Rate = txtRate(3)
                .Minimum = txtMinTk(3)
                .Save
            End With

            Flash_Into_Grid
            Clear_Screen
           ' POP_Param
            
        txtSlubCode(3).SetFocus
End Sub

Private Sub DataGrid1_Click(Index As Integer)

   On Error Resume Next

'        txtSlubCode(3) = !Slab_code
'        txtBasic_From(3) = HR_Allow_Rs!Basic_From
'        txtBasic_To(3) = HR_Allow_Rs!Basic_To
'        txtRate(3) = HR_Allow_Rs!Rate
'        txtMinTk(3) = HR_Allow_Rs!Minimum
'        txtBasic_From(3).SetFocus
'        Exit Sub
'
'    txtBasic_From(3).SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
   Screen_Position Me
   Flash_Into_Grid
End Sub

Public Sub Flash_Into_Grid()

        With HR_Allow
           .Connstring = strCN.Connection_String
            
            Set DataGrid1(3).DataSource = .GetAll
            
        End With

        

        With DataGrid1(3)
            .Columns(0).Width = 1050
            '.Columns(0).DataField = Desig_Rs!Fields(0)

            .Columns(1).Width = 1480
            '.Columns(1).DataField = Desig_Rs!Fields(1)

            .Columns(2).Width = 1550
            '.Columns(2).DataField = Desig_Rs!Fields(2)

            .Columns(3).Width = 690
            '.Columns(3).DataField = Desig_Rs!Fields(3)

            .Columns(4).Width = 930
            '.Columns(3).DataField = Desig_Rs!Fields(3)
        End With

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Destroy Me
End Sub
Private Sub txtBasic_From_KeyPress(Index As Integer, KeyAscii As Integer)
      KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
Private Sub txtBasic_To_KeyPress(Index As Integer, KeyAscii As Integer)
      KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtMinTk_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtProp_Rt_KeyPress(KeyAscii As MSForms.ReturnInteger)
      KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
Private Sub txtRate_KeyPress(Index As Integer, KeyAscii As Integer)
      KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub


Private Sub txtSlubCode_KeyPress(Index As Integer, KeyAscii As Integer)

 If KeyAscii = 13 And txtSlubCode(3) <> "" Then
    
     With HR_Allow
            .Connstring = strCN.Connection_String
            .Slab_code = txtSlubCode(3)
            .GetX
            
         txtBasic_From(3) = .Basic_From
         txtBasic_To(3) = .Basic_To
         txtRate(3) = .Rate
         txtMinTk(3) = .Minimum
         
        End With
    End If

End Sub
