VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form28 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Budget Variance Report"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   Icon            =   "frm_budget_report.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00B18A2E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   0
      Left            =   -30
      TabIndex        =   4
      Top             =   630
      Width           =   8325
      Begin VB.Frame Frame3 
         BackColor       =   &H00B18A2E&
         Caption         =   "Fiscal Year"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   645
         Left            =   0
         TabIndex        =   9
         Top             =   570
         Width           =   8385
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   210
            Width           =   4515
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   210
            Width           =   1455
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Code "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   15
            Top             =   210
            Width           =   450
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   14
            Top             =   210
            Width           =   60
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Title"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   2850
            TabIndex        =   13
            Top             =   210
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   3330
            TabIndex        =   12
            Top             =   210
            Width           =   120
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   8385
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Summary"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   3120
            TabIndex        =   16
            Top             =   90
            Width           =   1305
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Income"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   7
            Top             =   90
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Expense"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1500
            TabIndex        =   6
            Top             =   90
            Width           =   1305
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00808080&
            Height          =   315
            Index           =   2
            Left            =   3030
            Top             =   60
            Width           =   1605
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00808080&
            Height          =   315
            Index           =   0
            Left            =   90
            Top             =   60
            Width           =   1305
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00808080&
            Height          =   315
            Index           =   1
            Left            =   1410
            Top             =   60
            Width           =   1605
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   345
            Left            =   5340
            TabIndex        =   8
            Top             =   60
            Width           =   2955
         End
      End
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
      Left            =   7140
      Picture         =   "frm_budget_report.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Preview"
      Top             =   2070
      Width           =   510
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   3780
      Top             =   2370
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3810
      Top             =   2370
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3810
      Top             =   2400
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
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
      Caption         =   ""
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
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   885
      Left            =   -30
      TabIndex        =   2
      Top             =   -120
      Width           =   8325
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report On"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3450
         TabIndex        =   17
         Top             =   150
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Yearly Budget Variance"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   2400
         TabIndex        =   3
         Top             =   390
         Width           =   2880
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3810
      Top             =   2400
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
      Left            =   7680
      Picture         =   "frm_budget_report.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   2070
      Width           =   510
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   7080
      Top             =   2040
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BorderColor     =   &H00000000&
      Height          =   750
      Left            =   0
      Top             =   1890
      Width           =   8310
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdADD_Click()
    txtField(0).Text = ""
    txtField(1).Text = ""
    txtField(2).Text = ""
    txtField(3).Text = ""
    Combo1.SetFocus
End Sub

Private Sub cmdDELETE_Click()
   If Len(Trim(Combo1.Text)) = 0 Then
       MsgBox "Fiscal Year Code Required", vbCritical, "IT Division,DNMIH"
       Combo1.SetFocus
       Exit Sub
    End If

    If Len(Trim(txtField(0).Text)) = 0 Then
       MsgBox "Account Code required", vbCritical, "IT Division,DNMIH"
       txtField(0).SetFocus
       Exit Sub
    End If
    If Len(Trim(txtField(3).Text)) = 0 Then
       MsgBox "Amount required", vbCritical, "IT Division,DNMIH"
       txtField(3).SetFocus
       Exit Sub
    End If

    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select * from fiscal_year where code=" & Trim(Combo1) & ""
    Adodc1.Refresh

    If Adodc1.Recordset.RecordCount = 0 Then
        MsgBox "No such code exists", vbCritical, "IT Division,DNMIH"
        Combo1.SetFocus
        Exit Sub
     End If

    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select * from budget where acc_code='" & Trim(txtField(0)) & "' and fiscal_yr_code=" & Trim(Combo1.Text) & ""
    Adodc1.Refresh

    If Adodc1.Recordset.RecordCount = 0 Then
        MsgBox "No such account code in Same fiscal year exists", vbCritical, "IT Division,DNMIH"
        Combo1.SetFocus
        Exit Sub
     End If
    If MsgBox("Are your sure to delete?", vbCritical + vbYesNo, "IT Division,DNMIH") = vbYes Then
             Call deletebudget
       MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
    End If
    Call load_grid
    Call load_fiscal
    cmdADD_Click
End Sub
Private Sub deletebudget()
  Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param0 As New Parameter
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter

    Dim userid As String
    userid = "Emdad"

    Conn.Open strcn.Connection_String

    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText

    '----------------------------------------------------------------------------------
    Set Param0 = cmd.CreateParameter("param0", adInteger, adParamInput, 5, 3)
    cmd.Parameters.Append Param0

    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 50, txtField(0).Text)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adDouble, adParamInput, 300, Val(txtField(3).Text))
    cmd.Parameters.Append Param2

    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, userid)
    cmd.Parameters.Append Param3

    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 30, Combo1.Text)
    cmd.Parameters.Append Param4


    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True

    cmd.CommandText = "{CALL s_u_d_budget(?,?,?,?,?)}"
    Set RS = cmd.Execute


    cmd.Properties("PLSQLRSet") = False

End Sub

   
'Private Sub deletefiscalyr()
'    Dim Conn As New ADODB.Connection
'    Dim cmd As New ADODB.Command
'    Dim RS As New ADODB.Recordset
'
'    Dim Param0 As New Parameter
'    Dim Param1 As New Parameter
'    Dim Param2 As New Parameter
'    Dim Param3 As New Parameter
'    Dim Param4 As New Parameter
'    Dim Param5 As New Parameter
'
'    Dim userid As String
'    userid = "Emdad"
'
'    Conn.Open strcn.Connection_String
'
'    Set cmd.ActiveConnection = Conn
'    cmd.CommandType = adCmdText
'
'    '----------------------------------------------------------------------------------
'    Set Param0 = cmd.CreateParameter("param0", adInteger, adParamInput, 5, 3)
'    cmd.Parameters.Append Param0
'
'    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 50, txtField(0).Text)
'    cmd.Parameters.Append Param1
'
'    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 300, txtField(1).Text)
'    cmd.Parameters.Append Param2
'
'    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, dtpdate(0).Value)
'    cmd.Parameters.Append Param3
'
'    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, dtpdate(1).Value)
'    cmd.Parameters.Append Param4
'
'    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 10, userid)
'    cmd.Parameters.Append Param5
'
'    '----------------------------------------------------------------------------------
'
'    cmd.Properties("PLSQLRSet") = True
'
'    cmd.CommandText = "{CALL save_fiscal_year(?,?,?,?,?,?)}"
'    Set RS = cmd.Execute
'
'
'    cmd.Properties("PLSQLRSet") = False
'
'End Sub
Private Sub cmdEXIT_Click()
    Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdSAVE_Click()
    If Len(Trim(Combo1.Text)) = 0 Then
       MsgBox "Fiscal Year Code Required", vbCritical, "IT Division,DNMIH"
       Combo1.SetFocus
       Exit Sub
    End If

    If Len(Trim(txtField(0).Text)) = 0 Then
       MsgBox "Account Code required", vbCritical, "IT Division,DNMIH"
       txtField(0).SetFocus
       Exit Sub
    End If
    If Len(Trim(txtField(3).Text)) = 0 Then
       MsgBox "Amount required", vbCritical, "IT Division,DNMIH"
       txtField(3).SetFocus
       Exit Sub
    End If

    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select * from fiscal_year where code=" & Trim(Combo1) & ""
    Adodc1.Refresh

    If Adodc1.Recordset.RecordCount = 0 Then
        MsgBox "No such code exists", vbCritical, "IT Division,DNMIH"
        Combo1.SetFocus
        Exit Sub
     End If

    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select * from budget where acc_code='" & Trim(txtField(0)) & "' and fiscal_yr_code=" & Trim(Combo1.Text) & ""
    Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        MsgBox "Same account code in Same fiscal year exists", vbCritical, "IT Division,DNMIH"
        Combo1.SetFocus
        Exit Sub
     End If

    Call savebudget
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
    Call load_grid
    Call load_fiscal
    cmdADD_Click
End Sub


Private Sub cmdPREVIEW_Click()
  If Combo1.Text = "" Then
     MsgBox "Fiscal Year required", vbInformation, "IT Division,DNMIH"
     Combo1.SetFocus
     Exit Sub
  End If
  
'   If Option1(0).Value = True Or Option1(1).Value = True Then
      rptMode = 24
      Screen.MousePointer = vbHourglass
      CRViewer1.Show vbModal
'    End If
      
End Sub

Private Sub Combo1_Change()
   Adodc3.ConnectionString = strcn.Connection_String
   Adodc3.RecordSource = "select  comp_setup from fiscal_year where code=" & Combo1.Text & " "
   Adodc3.Refresh

   If Adodc3.Recordset.RecordCount > 0 Then
      txtField(1) = Adodc3.Recordset!comp_setup
   End If
End Sub

Private Sub Combo1_Click()
 Adodc3.ConnectionString = strcn.Connection_String
   Adodc3.RecordSource = "select  comp_setup from fiscal_year where code=" & Combo1.Text & " "
   Adodc3.Refresh

   If Adodc3.Recordset.RecordCount > 0 Then
      txtField(1) = Adodc3.Recordset!comp_setup
   End If
End Sub

'Private Sub DataGrid1_Click()
'  If Adodc2.Recordset.RecordCount > 0 Then
'        txtField(0).Text = "" & DataGrid1.Columns(0).Text
'        txtField(2).Text = "" & DataGrid1.Columns(1).Text
'        txtField(3).Text = "" & DataGrid1.Columns(2).Text
'        Combo1.Text = "" & DataGrid1.Columns(3).Text
''        dtpdate(0).Value = "" & DataGrid1.Columns(2).Text
''        dtpdate(1).Value = "" & DataGrid1.Columns(3).Text
''
'  End If
'End Sub

Private Sub dtpdate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
        Case 0, 1
            If KeyCode = 13 Then
                SendKeys Chr(9)
            End If
        End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     SendKeys (Chr(9))
  End If
End Sub

Private Sub Form_Load()
    Call load_fiscal
    Call load_grid
    Label1.Caption = "Income Budget Report"
End Sub
Private Sub load_fiscal()

  Adodc1.ConnectionString = strcn.Connection_String
   Adodc1.RecordSource = "select  code  from fiscal_year"
   Adodc1.Refresh

  If Adodc1.Recordset.RecordCount > 0 Then
     Adodc1.Recordset.MoveFirst
     Do Until Adodc1.Recordset.EOF
       Combo1.AddItem Adodc1.Recordset!code
       Adodc1.Recordset.MoveNext
     Loop

   End If


End Sub
Private Sub load_grid()
   Adodc2.ConnectionString = strcn.Connection_String
   Adodc2.RecordSource = "select a.acc_code as code ,(select acc_name from acct  where acc_code=a.acc_code) as Title, a.proposed_amount as Amount,fiscal_yr_code as fiscal_year from budget a"
   Adodc2.Refresh
End Sub

Private Sub savebudget()
  Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param0 As New Parameter
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter

    Dim userid As String
    userid = "Emdad"

    Conn.Open strcn.Connection_String

    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText

    '----------------------------------------------------------------------------------
    Set Param0 = cmd.CreateParameter("param0", adInteger, adParamInput, 5, 1)
    cmd.Parameters.Append Param0

    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 50, txtField(0).Text)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adDouble, adParamInput, 300, Val(txtField(3).Text))
    cmd.Parameters.Append Param2

    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, userid)
    cmd.Parameters.Append Param3

    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 30, Combo1.Text)
    cmd.Parameters.Append Param4


    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True

    cmd.CommandText = "{CALL s_u_d_budget(?,?,?,?,?)}"
    Set RS = cmd.Execute


    cmd.Properties("PLSQLRSet") = False

End Sub

'Private Sub MSFlexGrid2_DblClick()
'   If Len(Trim(MSFlexGrid2.Text)) <> 0 Then
'       txtField(0).Text = MSFlexGrid2.Text
'       txtField_LostFocus (0)
''       nbrDebit.SetFocus
'       'nbrDollar.SetFocus
'    Else
'       txtField(0).SetFocus
'       txtField(2).Text = ""
'       txtField(3).Text = ""
'
'    End If
'    MSFlexGrid2.Visible = False
'
'End Sub

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'       SendKeys Chr(9)
'    End If
End Sub

'Private Sub MSFlexGrid2_LostFocus()
'   Call MSFlexGrid2_DblClick
'End Sub

Private Sub Option1_Click(Index As Integer)
  Select Case Index
         Case 0
              Label1.Caption = "Income Budget Report"
         Case 1
             Label1.Caption = "Expense Budget Report"
         Case 2
             Label1.Caption = "Summary Budget Report"
      End Select
End Sub

Private Sub txtField_Change(Index As Integer)
  Select Case Index
    Case 3
        If Not IsNumeric(txtField(3).Text) Then
                txtField(3).Text = ""
        End If
    End Select
End Sub

Private Sub txtField_Click(Index As Integer)
  Select Case Index
    Case 0
'        Call GetAccName(Me, Trim(txtField(0).Text))
    End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
        Case 0
           txtField(0).BackColor = &H80000018
         Case 1
            txtField(1).BackColor = &H80000018
     End Select
End Sub
'Private Sub getAcc_Code(strAcc_des As String)
'    On Error GoTo err_loop
'    MSFlexGrid2.Clear
'    MSFlexGrid2.Rows = 0
'
'    MSFlexGrid2.ColWidth(0) = "1200"
'    MSFlexGrid2.ColAlignment(0) = 1
'
'    MSFlexGrid2.ColWidth(1) = "5800"
'
'    Adodc3.ConnectionString = strcn.Connection_String
'    Adodc3.RecordSource = "select user_acc,acc_name from acct where acc_code not in(select acc_head from acct) and upper(acc_name) like '" & Trim(UCase(strAcc_des)) & "%'"
'    Adodc3.Refresh
'    If Adodc3.Recordset.RecordCount > 0 Then
'        Do Until Adodc3.Recordset.EOF
'            MSFlexGrid2.AddItem Adodc3.Recordset!user_acc & vbTab & Adodc3.Recordset!acc_name
'            Adodc3.Recordset.MoveNext
'       Loop
'    End If
'
'    MSFlexGrid2.Visible = True
'    MSFlexGrid2.SetFocus
'    Exit Sub
'err_loop:
'    MsgBox Err.Description, vbCritical
'    Resume Next
'End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
' Select Case Index
'        Case 0
'            If KeyAscii = 13 Then
'                If Len(Trim(txtField(0).Text)) = 0 Then
'                    cmdSAVE.SetFocus
'                Else
'                    SendKeys Chr(9)
'       End If
'    End If
'    End Select
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   Select Case Index
        Case 0
           txtField(0).BackColor = vbWhite
            If Len(Trim(txtField(0).Text)) = 0 Then Exit Sub
             Adodc4.ConnectionString = strcn.Connection_String
             Adodc4.RecordSource = "select acc_name from acct where user_acc='" & Trim(txtField(0).Text) & "'"
             Adodc4.Refresh
             If Adodc4.Recordset.RecordCount > 0 Then
                 txtField(2) = Adodc4.Recordset!acc_name
            End If

         Case 1
            txtField(1).BackColor = vbWhite
   End Select
End Sub
