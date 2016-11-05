VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form12 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Provident fund & other policy setup"
   ClientHeight    =   7155
   ClientLeft      =   2085
   ClientTop       =   2685
   ClientWidth     =   8670
   Icon            =   "frmParam_Setup1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   8670
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   4815
      Picture         =   "frmParam_Setup1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6570
      Width           =   1140
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   3435
      Picture         =   "frmParam_Setup1.frx":234C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6570
      Width           =   1140
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   2115
      Picture         =   "frmParam_Setup1.frx":3CDE
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6570
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6450
      Left            =   90
      TabIndex        =   20
      Top             =   45
      Width           =   8475
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   240
         Left            =   3915
         TabIndex        =   0
         Top             =   270
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Effective Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2205
         TabIndex        =   2
         Top             =   293
         Width           =   1995
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         Height          =   330
         Left            =   2115
         Top             =   225
         Width           =   4065
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Retention (seasonal) staff/worker will get                % of their basic salary during off season."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   630
         TabIndex        =   3
         Top             =   5070
         Width           =   6420
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employer will contribute                  % of the empolyee's contribution."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   630
         TabIndex        =   4
         Top             =   975
         Width           =   4770
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Yearly interest will be calculated  upon total contribution(Employee + Empoyer) at a rate of                %."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   630
         TabIndex        =   5
         Top             =   1290
         Width           =   7185
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Emplyee will contribute                   % of his/her basic salary to the PF trust each month. "
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   21
         Top             =   660
         Width           =   6180
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmParam_Setup1.frx":5670
         ForeColor       =   &H00800000&
         Height          =   555
         Index           =   2
         Left            =   630
         TabIndex        =   6
         Top             =   1605
         Width           =   6990
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum service age limit under pay commission is                   years."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   630
         TabIndex        =   7
         Top             =   2370
         Width           =   4860
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum service age limit under wage commission is                years."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   630
         TabIndex        =   8
         Top             =   2700
         Width           =   4860
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Holiday OT pay will be                % of daily basic pay , assuming                 hours a day."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   630
         TabIndex        =   9
         Top             =   4545
         Width           =   6105
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "OT pay will be                % of hourly basic pay."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   630
         TabIndex        =   10
         Top             =   4230
         Width           =   3255
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum working hour in a month is assumed to be                 hours ."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   630
         TabIndex        =   11
         Top             =   3915
         Width           =   4875
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Service period is assumed to be one year if month part is >=                 months under pay commission."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   630
         TabIndex        =   12
         Top             =   3180
         Width           =   7125
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Service period is assumed to be one year if month part is >=                 months under wage commission."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   630
         TabIndex        =   13
         Top             =   3495
         Width           =   7260
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "House rent allowance under wage commission is                        % of Baisc Salary."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   630
         TabIndex        =   14
         Top             =   5385
         Width           =   5775
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DA (Dearness Allowance) under pay commission is                     % of Basic Salary."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   13
         Left            =   630
         TabIndex        =   15
         Top             =   5700
         Width           =   5775
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DA (Dearness Allowance) under wage commission is                  % of Basic Salary."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   630
         TabIndex        =   16
         Top             =   6030
         Width           =   5775
      End
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New Connection
Dim cmd As New Command
Dim RS As New Recordset
Dim Policy As Integer
Dim Flag As Integer
Dim Value As String
Dim Index As Integer
Dim i As Integer
Dim parameterinfo As New ParameterMain
Private Sub Form_Load()
On Error GoTo Errdes
    Screen_Position Me
    POP_Current_Settings
    Get_Default_Value
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Public Sub Set_Param(Index As Integer)
On Error GoTo Errdes
Select Case Index

Case 0
    Policy = 21
Case 1
    Policy = 22
Case 2
    Policy = 23
Case 3
    Policy = 24
Case 4
    Policy = 25
Case 5
    Policy = 41
Case 6
    Policy = 42
Case 7
    Policy = 43
Case 8
    Policy = 44
Case 9
    Policy = 51
Case 10
    Policy = 52
Case 11
    Policy = 53
Case 12
    Policy = 54
Case 13
    Policy = 31
Case 14
    Policy = 32
Case 15
    Policy = 33
Case 16
    Policy = 34

End Select

Flag = 1
''Value = txtParam(Index)

con.ConnectionString = strCN.Connection_String

con.Open
cmd.CommandText = "exec Set_Param " + CStr(Policy) + "," _
        + CStr(Flag) + ",'" + Value + "'"

cmd.ActiveConnection = con
cmd.Execute
con.Close

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Public Sub POP_Param(Polc As Integer)
On Error Resume Next

    con.ConnectionString = strCN.Connection_String
    con.Open
    cmd.CommandText = "exec POP_Param " + CStr(Polc)

    cmd.ActiveConnection = con
    Set RS = cmd.Execute

    Policy = RS.Fields(0)
    Value = RS.Fields(2)
    con.Close

    Select Case Polc
        Case 21
            Index = 0
        Case 22
            Index = 1
        Case 23
            Index = 2
        Case 24
            Index = 3
        Case 25
            Index = 4
        '-------------------
        Case 31
            Index = 13
        Case 32
            Index = 14
        Case 33
            Index = 15
        Case 34
            Index = 16
        Case 41
            Index = 5
        Case 42
            Index = 6
        Case 43
            Index = 7
        Case 44
            Index = 8
        Case 51
            Index = 9
        Case 52
            Index = 10
        Case 53
            Index = 11
        Case 54
            Index = 12
        
            
    End Select

   '' txtParam(Index) = Value

End Sub
Public Sub POP_Current_Settings()
On Error Resume Next

For i = 21 To 54
   POP_Param (i)
Next

End Sub
Private Sub cmdClear_Click()
    Clear_Screen
    MaskEdBox1 = Format(Date, "dd/mm/yy")
   '' txtParam(0).SetFocus
End Sub
Private Sub cmdClose_Click()
    Close_Msg Me
End Sub
Private Sub cmdSave_Click()
'On Error GoTo Errdes
'With parameterinfo
'
'        .Connstring = strCN.Connection_String
'        .EFFDATE = MaskEdBox1
'        ''.EMPCONTRPF = txtParam(0)
'        '''.EMRCONTRPF = txtParam(1)
'        .YEARLYINTOFPF = txtParam(2)
'        .PFINCOMEINVESTDIST = txtParam(3)
'        .CUYRTOTALCONTRIBUTION = txtParam(4)
'        .MAXSEUPCOM = txtParam(5)
'        .MAXSEUPWOM = txtParam(6)
'        .SPEASSUPCOM = txtParam(7)
'        .SPEASSUWCOM = txtParam(8)
'        .MAXWHINMONTH = txtParam(9)
'        .OTPAYPHR = txtParam(10)
'        .HOLIDAYOT = txtParam(11)
'        .HOLIDAYOTHOUR = txtParam(12)
'        .SEAWRSALINOFFSE = txtParam(13)
'        .HOUSERENTUNWCOM = txtParam(14)
'        .DAUNDERPCOM = txtParam(15)
'        .DAUNDERWCOM = txtParam(16)
'        .Save
''End With

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
   
End Sub
Private Sub txtParam_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
     KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
Private Sub Get_Default_Value()
On Error GoTo Errdes
Dim getconnected As New Connection
Dim cmd As New Command
Dim myrs As New ADODB.Recordset

getconnected.ConnectionString = strCN.Connection_String
getconnected.Open
cmd.ActiveConnection = getconnected
cmd.CommandType = adCmdText
cmd.CommandText = "select * from parameter_main where effdate=(select max(effdate)  from parameter_main)"

cmd.Properties("iRowsetChange") = True
cmd.Properties("updatability") = 7
myrs.CursorLocation = adUseClient

myrs.Open cmd.CommandText, getconnected, adOpenDynamic, adLockOptimistic

If myrs.BOF = False Then
  
'    MaskEdBox1 = Format(myrs.Fields(0), "dd/mm/yy")
'    txtParam(0) = myrs.Fields(1)
'    txtParam(1) = myrs.Fields(2)
'    txtParam(2) = myrs.Fields(3)
'    txtParam(3) = myrs.Fields(4)
'    txtParam(4) = myrs.Fields(5)
'    txtParam(7) = myrs.Fields(6)
'    txtParam(8) = myrs.Fields(7)
'    txtParam(9) = myrs.Fields(8)
'    txtParam(10) = myrs.Fields(9)
'    txtParam(11) = myrs.Fields(10)
'    txtParam(12) = myrs.Fields(11)
'    txtParam(13) = myrs.Fields(12)
'    txtParam(14) = myrs.Fields(13)
'    txtParam(15) = myrs.Fields(14)
'    txtParam(16) = myrs.Fields(15)
'    txtParam(5) = myrs.Fields(16)
'    txtParam(6) = myrs.Fields(17)
End If
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
