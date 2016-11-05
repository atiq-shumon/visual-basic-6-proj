VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOpeningBalance 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10545
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Opening Blance (For PF) Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   1545
      Left            =   90
      TabIndex        =   14
      Top             =   90
      Width           =   10380
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   1280
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   810
         Width           =   1335
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   7845
         TabIndex        =   6
         Top             =   810
         Width           =   1020
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   810
         Width           =   1095
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   8880
         TabIndex        =   7
         Top             =   810
         Width           =   1140
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   6795
         TabIndex        =   5
         Top             =   810
         Width           =   1020
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   5760
         TabIndex        =   4
         Top             =   810
         Width           =   1020
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   4720
         TabIndex        =   3
         Top             =   810
         Width           =   1020
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   3690
         TabIndex        =   2
         Top             =   810
         Width           =   1020
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   2635
         TabIndex        =   1
         Top             =   810
         Width           =   1020
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   24
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Opening Bal."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7800
         TabIndex        =   22
         Top             =   495
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   135
         Top             =   765
         Width           =   9945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Closing Bal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8940
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Int.Amt."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6875
         TabIndex        =   20
         Top             =   495
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Empl.Cont"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5820
         TabIndex        =   19
         Top             =   495
         Width           =   1005
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Emp Cont."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4770
         TabIndex        =   18
         Top             =   495
         Width           =   1005
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "End Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3795
         TabIndex        =   17
         Top             =   495
         Width           =   945
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Beg.Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2670
         TabIndex        =   16
         Top             =   495
         Width           =   930
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Emp ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   495
         Width           =   1020
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00E0E0E0&
         Height          =   1050
         Left            =   90
         Shape           =   4  'Rounded Rectangle
         Top             =   315
         Width           =   10110
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Height          =   1680
      Left            =   90
      TabIndex        =   12
      Top             =   1665
      Width           =   10380
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1455
         Left            =   90
         TabIndex        =   13
         Top             =   135
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   2566
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
               LCID            =   2057
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
               LCID            =   2057
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
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   9195
      Picture         =   "frmOpeningBalance.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3465
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   6765
      Picture         =   "frmOpeningBalance.frx":1A82
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3465
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   5550
      Picture         =   "frmOpeningBalance.frx":3414
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3465
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   7980
      Picture         =   "frmOpeningBalance.frx":4DA6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3465
      Width           =   1185
   End
End
Attribute VB_Name = "frmOpeningBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PF_Open As New Cls_IncrementPro
Private Sub cmdClear_Click()
On Error GoTo Errdes
For i = 0 To 7
    txtfields(i).Text = ""
Next
Combo1.Text = ""
Combo1.SetFocus
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdDelete_Click()
On Error GoTo Errdes
With PF_Open
    .Connstring = strCN.Connection_String
    .Emp_ID = Combo1
    .BEGIN_YEAR = txtfields(0)
    .END_YEAR = txtfields(1)
    .PF_Closing_Delete
     MsgBox "Data has Deleted Successfully", vbInformation, "IT Division, DNMIH."
     cmdClear_Click
     Get_Value_Into_Grid
End With
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub cmdSave_Click()
On Error GoTo Errdes
With PF_Open
    .Connstring = strCN.Connection_String
    .Emp_ID = Combo1
    .BEGIN_YEAR = txtfields(0)
    .END_YEAR = txtfields(1)
    .EMP_CONTRUBUTION = txtfields(2)
    .Employeer_Contribution = txtfields(3)
    .INTEREST_AMOUNT = txtfields(4)
    .OPENING_AMOUNT = txtfields(6)
    .CLOSEING_AMOUNT = txtfields(5)
    .PF_Closing_Save
     MsgBox "Data Saved Successfully", vbInformation, "IT Division, DNMIH"
     Get_Value_Into_Grid
     cmdClear.SetFocus
End With
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    
End If
End Sub

Private Sub Combo1_LostFocus()
Get_Emp_Name
End Sub

Private Sub DataGrid1_Click()
On Error GoTo Errdes
Combo1 = DataGrid1.Columns(0)
txtfields(0).Text = DataGrid1.Columns(1)
txtfields(1).Text = DataGrid1.Columns(2)
txtfields(2).Text = DataGrid1.Columns(3)
txtfields(3).Text = DataGrid1.Columns(4)
txtfields(4).Text = DataGrid1.Columns(5)
txtfields(6).Text = DataGrid1.Columns(6)
txtfields(5).Text = DataGrid1.Columns(7)
Get_Emp_Name
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Form_Load()
On Error GoTo Errdes
Dim cmd As New Command
Dim conn10 As New Connection
Dim rs10 As New Recordset
Dim conn11 As New Connection
Dim rs11 As New Recordset

conn10.ConnectionString = strCN.Connection_String
conn10.Open
cmd.ActiveConnection = conn10
cmd.CommandType = adCmdText
cmd.CommandText = "select emp_id from emp_info order by emp_id "
rs10.CursorLocation = adUseClient
rs10.Open cmd.CommandText, conn10, adOpenDynamic, adLockOptimistic

If rs10.RecordCount > 0 Then
    Do Until rs10.EOF
        Combo1.AddItem rs10.Fields(0)
        rs10.MoveNext
    Loop
    Combo1 = Combo1.List(0)
End If

rs10.Close
conn10.Close
Get_Value_Into_Grid
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Get_Value_Into_Grid()
On Error GoTo Errdes
    With PF_Open
        .Connstring = strCN.Connection_String
     Set DataGrid1.DataSource = .GetAll
    End With

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Get_Emp_Name()
On Error GoTo Errdesc
Dim cmd As New Command
Dim conn2 As New ADODB.Connection
Dim RS2 As New ADODB.Recordset
conn2.ConnectionString = strCN.Connection_String
conn2.Open
cmd.ActiveConnection = conn2
cmd.CommandType = adCmdText

cmd.CommandText = "select EMP_NM from emp_info where Emp_Id='" & Combo1 & "'"

RS2.CursorLocation = adUseClient
RS2.Open cmd.CommandText, conn2, adOpenDynamic, adLockOptimistic

    If RS2.RecordCount > 0 Then
        txtfields(7).Text = RS2.Fields(0)
        RS2.Close
        conn2.Close
    Else
        txtfields(7).Text = ""
        RS2.Close
        conn2.Close
    End If
Exit Sub
Errdesc:
MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub txtfields_LostFocus(Index As Integer)
Select Case Index


Case 0
    
    If Len(Trim(txtfields(0))) = 0 Then
            txtfields(0).Text = 2000
     End If
     
Case 1
    
    If Len(Trim(txtfields(1))) = 0 Then
            txtfields(1).Text = YEAR(Now)
     End If


Case 2
    
    If Len(Trim(txtfields(2))) = 0 Then
            txtfields(2).Text = 0
     End If

Case 3
    'If Len(Trim(txtfields(3))) = 0 Then
            txtfields(3).Text = Val(txtfields(2).Text)
    'Else
      
    'End If
     
Case 4
    If Len(Trim(txtfields(4))) = 0 Then
            txtfields(4).Text = 0
    End If

Case 5

    If Len(Trim(txtfields(5))) = 0 Then
            txtfields(5).Text = 0
     End If
     
Case 6
     If Len(Trim(txtfields(6))) = 0 Then
            txtfields(6).Text = 0
     End If


    txtfields(5) = Val(txtfields(2)) + Val(txtfields(3)) + Val(txtfields(4))
End Select

End Sub
