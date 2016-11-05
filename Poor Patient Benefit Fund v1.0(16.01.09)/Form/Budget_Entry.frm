VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form27 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9495
   Icon            =   "Budget_Entry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEXIT 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      ToolTipText     =   "CLOSE"
      Top             =   7380
      Width           =   1215
   End
   Begin VB.CommandButton cmdPREVIEW 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   6930
      TabIndex        =   5
      ToolTipText     =   "VIEW REPORT"
      Top             =   7380
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   5700
      TabIndex        =   6
      ToolTipText     =   "DELETE"
      Top             =   7380
      Width           =   1215
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "NEW"
      Height          =   375
      Left            =   4470
      TabIndex        =   4
      ToolTipText     =   "NEW ENTRY"
      Top             =   7380
      Width           =   1215
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      ToolTipText     =   "SAVE DATA"
      Top             =   7380
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   3210
      Top             =   7470
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
      Left            =   3240
      Top             =   7470
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
      Left            =   3210
      Top             =   7470
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
      Height          =   765
      Left            =   -30
      TabIndex        =   8
      Top             =   -120
      Width           =   10245
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proposed Budget  Entry"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Left            =   2040
         TabIndex        =   9
         Top             =   120
         Width           =   5280
      End
   End
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
      Height          =   6585
      Index           =   0
      Left            =   -30
      TabIndex        =   11
      Top             =   660
      Width           =   9525
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   7890
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1170
         Width           =   1545
      End
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
         Height          =   975
         Left            =   0
         TabIndex        =   10
         Top             =   -30
         Width           =   9585
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00B18A2E&
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
            ForeColor       =   &H00C0FFC0&
            Height          =   225
            Index           =   0
            Left            =   2670
            TabIndex        =   24
            Top             =   660
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00B18A2E&
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
            ForeColor       =   &H00C0FFC0&
            Height          =   225
            Index           =   1
            Left            =   6330
            TabIndex        =   23
            Top             =   660
            Width           =   1305
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   630
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   210
            Width           =   1455
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   210
            Width           =   6795
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0FF&
            FillColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   6300
            Top             =   630
            Width           =   3075
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0FF&
            FillColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   2610
            Top             =   630
            Width           =   2745
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
            Left            =   570
            TabIndex        =   15
            Top             =   840
            Width           =   120
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Title: "
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
            Left            =   2130
            TabIndex        =   14
            Top             =   240
            Width           =   465
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
            TabIndex        =   13
            Top             =   240
            Width           =   60
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
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   12
            Top             =   240
            Width           =   405
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   1410
         Left            =   2010
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2487
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   14737632
         ForeColor       =   8388608
         BackColorFixed  =   14737632
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483635
         BackColorBkg    =   16777215
         FocusRect       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtField 
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   1170
         Width           =   1785
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1860
         TabIndex        =   21
         Top             =   1170
         Width           =   6075
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Budget_Entry.frx":030A
         Height          =   5115
         Left            =   60
         TabIndex        =   22
         Top             =   1440
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9022
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   -2147483624
         BorderStyle     =   0
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
         RowDividerStyle =   4
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
               Format          =   "0"
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proposed Amount"
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
         Index           =   6
         Left            =   7920
         TabIndex        =   18
         Top             =   960
         Width           =   1380
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
         Index           =   5
         Left            =   1980
         TabIndex        =   17
         Top             =   960
         Width           =   360
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
         Index           =   4
         Left            =   180
         TabIndex        =   16
         Top             =   960
         Width           =   450
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3240
      Top             =   7470
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
   Begin VB.Shape Shape1 
      Height          =   465
      Index           =   2
      Left            =   3150
      Top             =   7320
      Width           =   6285
   End
End
Attribute VB_Name = "Form27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public income_exp_indicator As Integer
Dim Income_Exp_var As Integer
Private Sub cmdADD_Click()
    txtField(0).Text = ""
    txtField(1).Text = ""
    txtField(2).Text = ""
    txtField(3).Text = ""
    Combo1.SetFocus
End Sub

Private Sub cmdDELETE_Click()
   If Len(Trim(Combo1.Text)) = 0 Then
       MsgBox "Fiscal Year Code Required", vbCritical, "IT Division, DNMIH"
       Combo1.SetFocus
       Exit Sub
    End If

    If Len(Trim(txtField(0).Text)) = 0 Then
       MsgBox "Account Code required", vbCritical, "IT Division, DNMIH"
       txtField(0).SetFocus
       Exit Sub
    End If
    If Len(Trim(txtField(3).Text)) = 0 Then
       MsgBox "Amount required", vbCritical, "IT Division, DNMIH"
       txtField(3).SetFocus
       Exit Sub
    End If

    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select * from fiscal_year where code=" & Trim(Combo1) & ""
    Adodc1.Refresh

    If Adodc1.Recordset.RecordCount = 0 Then
        MsgBox "No such code exists", vbCritical, "IT Division, DNMIH"
        Combo1.SetFocus
        Exit Sub
     End If

    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select * from budget where acc_code='" & Trim(txtField(0)) & "' and fiscal_yr_code=" & Trim(Combo1.Text) & ""
    Adodc1.Refresh

    If Adodc1.Recordset.RecordCount = 0 Then
        MsgBox "No such account code in Same fiscal year exists", vbCritical, "IT Division, DNMIH"
        Combo1.SetFocus
        Exit Sub
     End If
    If MsgBox("Are your sure to delete?", vbCritical + vbYesNo, "IT Division, DNMIH") = vbYes Then
             Call deletebudget
       MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
    End If
    Call load_grid(Income_Exp_var)
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
    userid = Form1.Label2(2).Caption

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

    Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 3, 0)
    cmd.Parameters.Append Param5

    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True

    cmd.CommandText = "{CALL s_u_d_budget(?,?,?,?,?,?)}"
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

Private Sub cmdPREVIEW_Click()
  Form28.Show vbModal
End Sub

Private Sub cmdSAVE_Click()
    If Len(Trim(Combo1.Text)) = 0 Then
       MsgBox "Fiscal Year Code Required", vbCritical, "IT Division, DNMIH"
       Combo1.SetFocus
       Exit Sub
    End If

    If Len(Trim(txtField(0).Text)) = 0 Then
       MsgBox "Account Code required", vbCritical, "IT Division, DNMIH"
       txtField(0).SetFocus
       Exit Sub
    End If
    If Len(Trim(txtField(3).Text)) = 0 Then
       MsgBox "Amount required", vbCritical, "IT Division, DNMIH"
       txtField(3).SetFocus
       Exit Sub
    End If

'    Adodc1.ConnectionString = strcn.Connection_String
'    Adodc1.RecordSource = "select * from fiscal_year where code=" & Trim(Combo1) & ""
'    Adodc1.Refresh
'
'    If Adodc1.Recordset.RecordCount = 0 Then
'        MsgBox "No such code exists", vbCritical, "IT Division, DNMIH"
'        Combo1.SetFocus
'        Exit Sub
'     End If

    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select * from budget where acc_code='" & Trim(txtField(0)) & "' and fiscal_yr_code=" & Trim(Combo1.Text) & ""
    Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        MsgBox "Same account code in Same fiscal year exists", vbCritical, "IT Division, DNMIH"
        Combo1.SetFocus
        Exit Sub
     End If
    
    
    If Option1(0).Value = True Then
       income_exp_indicator = 1
    Else
       income_exp_indicator = 2
    End If
    
    Call savebudget
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
    Call load_grid(Income_Exp_var)
    
    cmdADD_Click
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
   
   load_fiscal_data
End Sub

Private Sub load_fiscal_data()
   Adodc2.ConnectionString = strcn.Connection_String
   Adodc2.RecordSource = "select a.acc_code as code ,(select acc_name from acct  where acc_code=a.acc_code) as Title, a.proposed_amount as Amount,a.fiscal_yr_code as fiscal_year from budget a where a.fiscal_yr_code='" & Trim(Combo1) & "' and INCOME_EXP_INDICATOR=" & Income_Exp_var & ""
   Adodc2.Refresh

   format_grid
End Sub
Private Sub DataGrid1_Click()
  On Error GoTo err_desc
        txtField(0).Text = "" & DataGrid1.Columns(0).Text
        txtField(2).Text = "" & DataGrid1.Columns(1).Text
        txtField(3).Text = "" & DataGrid1.Columns(2).Text
        Combo1.Text = "" & DataGrid1.Columns(3).Text
Exit Sub
err_desc:
        MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub

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
    Income_Exp_var = 1
    Call load_fiscal
    Call load_grid(Income_Exp_var)
End Sub
Private Sub load_fiscal()

  Adodc1.ConnectionString = strcn.Connection_String
  Adodc1.RecordSource = "select  code  from fiscal_year"
  Adodc1.Refresh

  
  
  If Adodc1.Recordset.RecordCount > 0 Then
     Adodc1.Recordset.MoveFirst
     Combo1.Clear
     Do Until Adodc1.Recordset.EOF
       Combo1.AddItem Adodc1.Recordset!code
       Adodc1.Recordset.MoveNext
     Loop

   End If


End Sub
Private Sub load_grid(mode As Integer)
   
   Select Case mode
          Case 1, 2
              Adodc2.ConnectionString = strcn.Connection_String
              Adodc2.RecordSource = "select a.acc_code as code ,(select acc_name from acct  where acc_code=a.acc_code) as Title, a.proposed_amount as Amount,fiscal_yr_code as fiscal_year from budget a where a.fiscal_yr_code='" & Trim(Combo1) & "' and a.INCOME_EXP_INDICATOR=" & mode & " "
              Adodc2.Refresh
    End Select

   format_grid
End Sub
Private Sub format_grid()
   If Adodc2.Recordset.RecordCount > 0 Then
      With DataGrid1
           .Columns(0).Width = 1500
           .Columns(1).Width = 6040
           .Columns(2).Width = 1800
      End With
      
  End If
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
    userid = Form1.Label2(2).Caption

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
    
    Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 5, income_exp_indicator)
    cmd.Parameters.Append Param5



    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True

    cmd.CommandText = "{CALL s_u_d_budget(?,?,?,?,?,?)}"
    Set RS = cmd.Execute


    cmd.Properties("PLSQLRSet") = False

End Sub

Private Sub MSFlexGrid2_DblClick()
   If Len(Trim(MSFlexGrid2.Text)) <> 0 Then
       txtField(0).Text = MSFlexGrid2.Text
       txtField_LostFocus (0)
'       nbrDebit.SetFocus
       'nbrDollar.SetFocus
    Else
       txtField(0).SetFocus
       txtField(2).Text = ""
       txtField(3).Text = ""

    End If
    MSFlexGrid2.Visible = False

End Sub

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'       SendKeys Chr(9)
'    End If
End Sub

Private Sub MSFlexGrid2_LostFocus()
   Call MSFlexGrid2_DblClick
End Sub

Private Sub Option1_Click(Index As Integer)
  Select Case Index
         Case 0
              Income_Exp_var = 1
              load_grid (Income_Exp_var)
         Case 1
              Income_Exp_var = 2
              load_grid (Income_Exp_var)
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
Private Sub getAcc_Code(strAcc_des As String)
    On Error GoTo err_loop
    MSFlexGrid2.Clear
    MSFlexGrid2.Rows = 0

    MSFlexGrid2.ColWidth(0) = "1200"
    MSFlexGrid2.ColAlignment(0) = 1

    MSFlexGrid2.ColWidth(1) = "5800"

    Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "select user_acc,acc_name from acct WHERE  upper(acc_name) like '" & Trim(UCase(strAcc_des)) & "%'"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
        Do Until Adodc3.Recordset.EOF
            MSFlexGrid2.AddItem Adodc3.Recordset!user_acc & vbTab & Adodc3.Recordset!acc_name
            Adodc3.Recordset.MoveNext
       Loop
    End If

    MSFlexGrid2.Visible = True
    MSFlexGrid2.SetFocus
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

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
             Adodc4.RecordSource = "select acc_name from acct where upper(user_acc)=upper('" & Trim(txtField(0).Text) & "')"
             Adodc4.Refresh
             If Adodc4.Recordset.RecordCount > 0 Then
                 txtField(2) = Adodc4.Recordset!acc_name
             Else
                MSFlexGrid2.Left = txtField(0).Left
                MSFlexGrid2.Top = txtField(0).Top
                MSFlexGrid2.TabIndex = txtField(0).TabIndex + 1
                Call getAcc_Code(Trim(txtField(0)))
            Exit Sub
            End If
            
'             If Option1(0).Value = True Then
'                Adodc4.ConnectionString = strcn.Connection_String
'                Adodc4.RecordSource = "select acc_code from acct where acc_code = '" & Trim(txtField(0).Text) & "' and (acc_code like '61%' or acc_code like '91%')"
'                Adodc4.Refresh
'
'
'                If Adodc4.Recordset.RecordCount = 0 Then
'                    MsgBox "No such Income accounts exists", vbInformation, "IT Division, DNMIH"
'
'                    txtField(0).Text = ""
'                    txtField(1).Text = ""
'                    txtField(2).Text = ""
'                    txtField(0).SetFocus
'                     Exit Sub
'                End If
'
'            Else
'               Adodc4.ConnectionString = strcn.Connection_String
'                Adodc4.RecordSource = "select acc_code from acct where acc_code = '" & Trim(txtField(0).Text) & "' and (acc_code like '81%' or acc_code like '92%'or acc_code like '11%' or acc_code like '2101%')"
'                Adodc4.Refresh
'
'                If Adodc4.Recordset.RecordCount = 0 Then
'                   MsgBox "No such Expense accounts exists", vbInformation, "IT Division, DNMIH"
'
'                   txtField(0).Text = ""
'                   txtField(1).Text = ""
'                    txtField(2).Text = ""
'                   txtField(0).SetFocus
'                   Exit Sub
'                End If
'           End If
'
'
'
'

         Case 1
            txtField(1).BackColor = vbWhite
   End Select
End Sub
