VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form12 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedule Report"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   FillStyle       =   3  'Vertical Line
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RptSchedule.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   975
      Left            =   -150
      TabIndex        =   8
      Top             =   -180
      Width           =   6585
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   540
         Left            =   4380
         TabIndex        =   9
         Top             =   150
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2025
      Left            =   -30
      TabIndex        =   13
      Top             =   600
      Width           =   6495
      Begin VB.ComboBox cboHeadName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1950
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   525
         Width           =   4080
      End
      Begin VB.ComboBox cboUserHead 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   195
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   525
         Width           =   1740
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   3960
         TabIndex        =   14
         Top             =   1350
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22806529
         CurrentDate     =   38518
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   210
         TabIndex        =   15
         Top             =   1350
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22806529
         CurrentDate     =   38518
      End
      Begin VB.Shape Shape7 
         Height          =   405
         Left            =   1920
         Top             =   480
         Width           =   4125
      End
      Begin VB.Shape Shape6 
         Height          =   405
         Left            =   150
         Top             =   480
         Width           =   1785
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Control Head"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   270
         TabIndex        =   22
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Control Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1965
         TabIndex        =   21
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   4410
         TabIndex        =   20
         Top             =   960
         Width           =   330
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Range :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2580
         TabIndex        =   19
         Top             =   1410
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   270
         TabIndex        =   18
         Top             =   960
         Width           =   525
      End
      Begin VB.Shape Shape2 
         Height          =   405
         Left            =   180
         Top             =   1320
         Width           =   2145
      End
      Begin VB.Shape Shape5 
         Height          =   405
         Left            =   3930
         Top             =   1320
         Width           =   2145
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   825
      Left            =   -60
      TabIndex        =   10
      Top             =   2460
      Width           =   6645
      Begin VB.CommandButton cmdEXIT 
         BackColor       =   &H00C0C0C0&
         Cancel          =   -1  'True
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5790
         Picture         =   "RptSchedule.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exit"
         Top             =   270
         Width           =   510
      End
      Begin VB.CommandButton cmdPreview 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5280
         Picture         =   "RptSchedule.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Preview"
         Top             =   270
         Width           =   510
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   525
         Left            =   5220
         Top             =   210
         Width           =   1125
      End
   End
   Begin VB.TextBox txtuseracc 
      Height          =   495
      Left            =   3780
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   405
      Left            =   1440
      Top             =   2550
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox nbrDepRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   11100
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   1260
      Width           =   405
   End
   Begin VB.ListBox lstCheckAccName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1200
      Left            =   9780
      Sorted          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2610
      Visible         =   0   'False
      Width           =   4065
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3780
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtAccHead 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   8460
      TabIndex        =   5
      Text            =   "AccHead"
      Top             =   3330
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox nbrAccBudg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9900
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   1260
      Width           =   1185
   End
   Begin VB.TextBox nbrTrack_id 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7860
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "Tack_Id"
      Top             =   3270
      Visible         =   0   'False
      Width           =   705
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   4950
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000000&
      Height          =   945
      Left            =   -30
      Top             =   1620
      Width           =   6480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dep. Rate"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   10755
      TabIndex        =   6
      Top             =   990
      Width           =   735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Budget"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   10080
      TabIndex        =   3
      Top             =   990
      Width           =   510
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      Height          =   735
      Left            =   -30
      Top             =   900
      Width           =   6510
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub SaveAcct()
On Error GoTo err_loop
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter

    
    Dim userid As String
    userid = "Emdad"
    Dim DepAcc As Double
    DepAcc = 100
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 40, txtAccHead.Text)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 40, txtUserAcc.Text)
    cmd.Parameters.Append Param2
'
'    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 40, txtAccName.Text)
'    cmd.Parameters.Append Param3
    
'    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 40, txtBangla_Name.Text)
'    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 9, Val(nbrAccBudg.Text))
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 9, Val(nbrDepRate.Text))
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 9, Val(DepAcc))
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 50, userid)
    cmd.Parameters.Append Param8
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL SaveAcct(?, ?, ?, ?, ?, ?, ?, ?)}"
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub


Private Sub ClearScreen()

    Me.txtUserAcc.Text = ""
'    txtBangla_Name.Text = ""
'    Me.txtAccName.Text = ""
    Me.nbrAccBudg.Text = "0.00"
    Me.nbrDepRate.Text = "0.00"
    Me.nbrTrack_id.Text = ""
    
End Sub

Private Sub cboHeadName_Click()

'    Dim Conn As New Connection
'    Dim cmd As New Command
'    Dim RS As New Recordset
'
'    Conn.Open strcn.Connection_String
'    Set cmd.ActiveConnection = Conn
'
'    cmd.CommandType = adCmdText
'    cmd.CommandText = "Select user_acc,acc_name from Acct where acc_name= '" + Trim(cboHeadName.Text) + "'"
''    cmd.Properties("IRowsetChange") = True
''    cmd.Properties("Updatability") = 7
'
'    RS.CursorLocation = adUseClient
'    RS.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
'
'    If Not (RS.EOF And RS.BOF) Then
'        cboUserHead.Text = RS("User_acc")
'    Else
'        cboUserHead.Text = ""
'    End If
End Sub

Private Sub cboHeadName_GotFocus()
'    Call ShowControl
End Sub

Private Sub cboHeadName_LostFocus()

'    If Len(Trim(cboHeadName.Text)) = 0 Then Exit Sub
'    Call GetControlCode(Me, Trim(Me.cboHeadName.Text))
    
End Sub
Private Sub cboUserHead_Click()
    
'    Call GetControlName(Me, Trim(Me.cboUserHead.Text))
    
    Adodc2.ConnectionString = strcn.Connection_String
    Adodc2.RecordSource = "select acc_code,user_acc,acc_name from Acct where user_acc='" + cboUserHead.Text + "'"
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then
            cboHeadName.Text = Adodc2.Recordset!acc_name
            txtAccHead.Text = Adodc2.Recordset!acc_code
    End If
    Call AutoAccCode
'    Adodc4.ConnectionString = strcn.Connection_String
'    Adodc4.RecordSource = "select max(acc_code)as code from Acct where acc_head='" + txtAccHead.Text + "'"
'    Adodc4.Refresh
'    If Adodc4.Recordset.RecordCount > 0 Then
'
'        txtUserAcc.Text = Adodc4.Recordset!code + 1
'    End If
    'Call GetGrdData
End Sub

Private Sub cboUserHead_GotFocus()
    Call ShowControl
End Sub

Private Sub cmdADD_Click()
    Call ClearScreen
End Sub
Private Sub AutoAccCode()
    On Error GoTo err_loop

    Adodc4.ConnectionString = strcn.Connection_String
    Adodc4.RecordSource = "select max(acc_code) as new_acct from acct where acc_head='" & Trim(txtAccHead.Text) & "'"
    Adodc4.Refresh
    

    If IsNull(Adodc4.Recordset!new_acct) = True Then
       If Len(Trim(txtAccHead.Text)) <= 2 Then
          txtUserAcc.Text = (Val(txtAccHead.Text) * 100) + 1
       Else
          txtUserAcc.Text = (Val(txtAccHead.Text) * 1000) + 1
       End If
    Else
       If Val(txtAccHead.Text) = Val(Adodc4.Recordset!new_acct) Then
          If Len(Trim(txtAccHead.Text)) <= 2 Then
             txtUserAcc.Text = (Val(txtAccHead.Text) * 100) + 1
          Else
             txtUserAcc.Text = (Val(txtAccHead.Text) * 1000) + 1
          End If
       Else
            txtUserAcc.Text = Val(Adodc4.Recordset!new_acct) + 1
       End If
    End If
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub
Private Sub cmdDELETE_Click()

    On Error GoTo err_loop
        If Len(Trim(Me.txtUserAcc.Text)) = 0 Then
           MsgBox "Accounts code required", vbCritical, "Warning..."
           Me.txtUserAcc.Text = ""
           Me.txtUserAcc.SetFocus
           Exit Sub
        End If

        If Val(Me.nbrTrack_id.Text) <= 0 Then
           MsgBox "Item not selected", vbCritical, "Warning..."
           Exit Sub
        End If
        '''Checking control head''''''''''''''
           Adodc2.ConnectionString = strcn.Connection_String
            Adodc2.RecordSource = "select acc_head from acct where acc_head in(select acc_code from acct where track_id=" & Val(nbrTrack_id.Text) & ")"
            Adodc2.Refresh
            If Adodc2.Recordset.RecordCount > 0 Then
                MsgBox ("You can not delete Control Head"), vbCritical, "Warning..."
            Exit Sub
            End If
            
'
'            Adodc3.ConnectionString = strcn.Connection_String
'            Adodc3.RecordSource = "select acc_code from ledger where acc_code in(select acc_code from acct where track_id=" & Val(nbrTrack_id.Text) & ")"
'            Adodc3.Refresh
'            If Adodc3.Recordset.RecordCount > 0 Then
'                MsgBox ("Code in use"), vbCritical, "Warning..."
'            Exit Sub
'            End If
        '''---------------------------------------------------------------
        
    
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
        
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 4, Val(nbrTrack_id.Text))
    cmd.Parameters.Append Param1
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL DeleteAcct(?)}"
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    
    
   ' Call GetGrdData
   ' Call ClearScreen
 
     Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

'Private Sub cmdEdit_Click()
'
'    On Error GoTo err_loop
'
'    If Len(Trim(Me.txtUserAcc.Text)) = 0 Then
'        MsgBox "Accounts code required", vbCritical
'        txtUserAcc.SetFocus
'        Exit Sub
'    End If
'
'    If Len(Trim(Me.txtAccName.Text)) = 0 Then
'        MsgBox "Accounts name required", vbCritical
'        txtAccName.SetFocus
'        Exit Sub
'    End If
'
'    Dim Conn As New ADODB.Connection
'    Dim cmd As New ADODB.Command
'    Dim RS As New ADODB.Recordset
'
'    Dim Param1 As New Parameter
'    Dim Param2 As New Parameter
'    Dim Param3 As New Parameter
'    Dim Param4 As New Parameter
'    Dim Param5 As New Parameter
'    Dim Param6 As New Parameter
'
'
'
'    Dim userid As String
'    userid = "Emdad"
'    Dim DepAcc As Double
'    DepAcc = 100
'    Conn.Open strcn.Connection_String
'
'    Set cmd.ActiveConnection = Conn
'    cmd.CommandType = adCmdText
'
'    '----------------------------------------------------------------------------------
'    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 40, txtUserAcc.Text)
'    cmd.Parameters.Append Param1
'
'    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 40, txtAccName.Text)
'    cmd.Parameters.Append Param2
'
'    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 9, Val(nbrAccBudg.Text))
'    cmd.Parameters.Append Param3
'
'    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 9, Val(nbrDepRate.Text))
'    cmd.Parameters.Append Param4
'
'    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 9, Val(nbrTrack_id.Text))
'    cmd.Parameters.Append Param5
'
'    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 50, userid)
'    cmd.Parameters.Append Param6
'    '----------------------------------------------------------------------------------
'
'    cmd.Properties("PLSQLRSet") = True
'
'    cmd.CommandText = "{CALL EditAcct(?, ?, ?, ?, ?, ?)}"
'    Set RS = cmd.Execute
'
'
'    cmd.Properties("PLSQLRSet") = False
'
'
'    Call GetGrdData
'    Exit Sub
'err_loop:
'    MsgBox Err.Description, vbCritical
'    Resume Next
'
'
'End Sub

Private Sub cmdEXIT_Click()

    Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdPREVIEW_Click()
 Screen.MousePointer = vbHourglass
     rptMode = 8
   CRViewer1.Show vbModal
    
End Sub

'Private Sub cmdSAVE_Click()
'
'     If Len(Trim(Me.cboUserHead.Text)) = 0 Then
'        MsgBox "Control head required", vbCritical
'        cboUserHead.SetFocus
'        Exit Sub
'     End If
'
'     If Len(Trim(Me.cboHeadName.Text)) = 0 Then
'        MsgBox "Control name required", vbCritical
'        cboHeadName.SetFocus
'        Exit Sub
'     End If
'
'     If Len(Trim(Me.txtUserAcc.Text)) = 0 Then
'        MsgBox "Accounts code required", vbCritical
'        txtUserAcc.SetFocus
'        Exit Sub
'     End If
'
'     If Len(Trim(Me.txtAccName.Text)) = 0 Then
'        MsgBox "Accounts name required", vbCritical
'        txtAccName.SetFocus
'        Exit Sub
'     End If
'
'     If Len(Trim(txtBangla_Name.Text)) = 0 Then
'        MsgBox "Bangla Accounts name required", vbCritical
'        txtBangla_Name.SetFocus
'        Exit Sub
'     End If
'     Call SaveAcct
'     MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
'     Call GetGrdData
'     txtUserAcc.Text = ""
'     txtAccName.Text = ""
''    txtBangla_Name.Text = ""
'     txtAccName.SetFocus
''     Call ShowControl
'     Call AutoAccCode
'End Sub
'
'Private Sub DataGrid1_Click()
'
'    Me.txtuseracc.Text = Me.DataGrid1.Columns(0).Text
'   Me.txtAccName.Text = Me.DataGrid1.Columns(1).Text
'    Me.nbrAccBudg.Text = Me.DataGrid1.Columns(2).Text
'    Me.nbrTrack_id.Text = Me.DataGrid1.Columns(4).Text
'    Me.lstCheckAccName.Visible = False
'
'End Sub
'
Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
    
End Sub

Private Sub Form_Load()

'     rptMode = 18
'     Call ClearScreen
     Call ShowControl
    ' Call GetGrdData
     
End Sub

'Private Sub GetGrdData()
''    On Error GoTo err_loop
'        Adodc1.ConnectionString = strcn.Connection_String
'        Adodc1.RecordSource = "select user_acc,acc_name,acc_name_beng,acc_budg,(select case A.acc_group " & _
'                          " WHEN 0 THEN 'Assets'" & _
'                          " WHEN 1 THEN 'Assets'" & _
'                          " WHEN 2 THEN  'Liabilities'" & _
'                          " WHEN 3 THEN 'Equity'" & _
'                          " WHEN 4 THEN  'Liabilities'" & _
'                          " WHEN 5 THEN  'Income'" & _
'                          " WHEN 6 THEN 'Expenses'" & _
'                          " WHEN 7 THEN  'Expenses'" & _
'                          " WHEN 8 THEN  'Expenses'" & _
'                          " WHEN 9 THEN  'Income'" & _
'                          " WHEN 10 THEN  'Expenses' END from Acct A where A.user_acc=Acct.user_acc) as category,track_id from acct " & _
'                          "where acc_head='" & Trim(Me.txtAccHead.Text) & "' and acc_code <>'" & Trim(Me.txtAccHead.Text) & "'"
'        Adodc1.Refresh
'
'        DataGrid1.Columns(0).Width = 1470.047
'        DataGrid1.Columns(0).Caption = "Code"
'        DataGrid1.Columns(0).Locked = True
'
'        DataGrid1.Columns(1).Width = 3950
'        DataGrid1.Columns(1).Caption = "Accounts Name(English)"
'
'        DataGrid1.Columns(2).Width = 3950
'        DataGrid1.Columns(2).Caption = "Accounts Name(Bangla)"
''        DataGrid1.Columns(2).CellText = "Accounts Name(Bangla)"
'
'
'        DataGrid1.Columns(3).Width = 1275.024
'        DataGrid1.Columns(3).Caption = "Budget"
'        DataGrid1.Columns(3).Alignment = dbgRight
'
'        DataGrid1.Columns(4).Width = 2500
'        DataGrid1.Columns(4).Caption = "Category"
'
'       DataGrid1.Columns(5).Width = 100
'        DataGrid1.Columns(5).Visible = False
''
''        DataGrid1.Columns(6).Width = 100
''        DataGrid1.Columns(6).Visible = False
'
''    Exit Sub
''err_loop:
''    MsgBox Err.Description, vbCritical
''    Resume Next
'
'End Sub
Private Sub ShowControl()
    Adodc2.ConnectionString = strcn.Connection_String
            Adodc2.RecordSource = "select acc_code,acc_name from Acct where acc_code  in(select acc_head from acct) and acc_lbl<>0"
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then
        cboUserHead.Clear
        cboHeadName.Clear
        Adodc2.Recordset.MoveFirst
       While Adodc2.Recordset.EOF = False
           cboUserHead.AddItem Adodc2.Recordset!acc_code
            cboHeadName.AddItem Adodc2.Recordset!acc_name
            Adodc2.Recordset.MoveNext

       Wend
    End If
End Sub

Private Sub nbrAccBudg_GotFocus()

    nbrAccBudg.SelLength = Len(nbrAccBudg.Text)
    
End Sub

Private Sub nbrAccBudg_KeyPress(KeyAscii As Integer)

    If KeyAscii > 26 Then
       If InStr("0123456789.+-", Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
    
End Sub

Private Sub nbrDepRate_GotFocus()

    nbrDepRate.SelLength = Len(nbrDepRate.Text)
    
End Sub

Private Sub nbrDepRate_KeyPress(KeyAscii As Integer)

    If KeyAscii > 26 Then
       If InStr("0123456789.+-", Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
    
End Sub

Private Sub txtAccName_Change()

'    On Error GoTo err_loop
'    If Len(Trim(txtAccName.Text)) = 0 Then
'       lstCheckAccName.Visible = False
'       Exit Sub
'    Else
'       Me.lstCheckAccName.Left = Me.txtAccName.Left
'       Me.lstCheckAccName.Top = Me.DataGrid1.Top
'
'       lstCheckAccName.Visible = True
'    End If
'
'    lstCheckAccName.Clear
'    Con.ConnectionString = strcn
'    Con.Open
'    RS.Open "select acc_name from acct where acc_name like '" & Trim(txtAccName.Text) & "%'", Con
'    If RS.EOF = False Then
'        Do Until RS.EOF
'            lstCheckAccName.AddItem RS!acc_name
'            RS.MoveNext
'        Loop
'    End If
'    RS.Close
'    Con.Close
'    Exit Sub
'err_loop:
'    MsgBox Err.Description, vbCritical
'    Resume Next
    
End Sub

Private Sub txtAccName_KeyPress(KeyAscii As Integer)

    If KeyAscii = 39 Then
       KeyAscii = Asc(Chr(96))
    End If
    
End Sub

Private Sub txtAccName_LostFocus()

    lstCheckAccName.Visible = False
    
End Sub
