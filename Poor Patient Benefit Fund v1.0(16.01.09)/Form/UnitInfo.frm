VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unit Information"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "UnitInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUnitName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1710
      TabIndex        =   1
      Top             =   405
      Width           =   4245
   End
   Begin VB.TextBox txtUnitCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   405
      Width           =   1500
   End
   Begin VB.CommandButton cmdSAVE 
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
      Left            =   315
      Picture         =   "UnitInfo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save"
      Top             =   2970
      Width           =   510
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
      Left            =   1845
      Picture         =   "UnitInfo.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exit"
      Top             =   2970
      Width           =   510
   End
   Begin VB.CommandButton cmdADD 
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
      Left            =   825
      Picture         =   "UnitInfo.frx":1292
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "New"
      Top             =   2970
      Width           =   510
   End
   Begin VB.CommandButton cmdDelete 
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
      Left            =   1335
      Picture         =   "UnitInfo.frx":18FC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Delete"
      Top             =   2970
      Width           =   510
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4770
      Top             =   -90
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "UnitInfo.frx":2436
      Height          =   1995
      Left            =   180
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3519
      _Version        =   393216
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
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
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4259.906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   135
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Centre"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1710
      TabIndex        =   6
      Top             =   135
      Width           =   825
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   690
      Index           =   0
      Left            =   180
      Top             =   2835
      Width           =   5775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdClear_Click()
    txtUnitCode.Text = ""
    txtUnitName.Text = ""
    txtUnitCode.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub SaveProject()
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter

    
    Dim userid As String
    userid = "Emdad"
    
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 50, txtUnitCode.Text)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 300, txtUnitName.Text)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 50, userid)
    cmd.Parameters.Append Param3
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL SaveProject(?, ?, ?)}"
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

End Sub

Private Sub cmdADD_Click()
    txtUnitCode.Text = ""
    txtUnitName.Text = ""
    txtUnitCode.SetFocus
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo err_loop
       Dim Conn As New Connection
       Dim cmd As New Command
       Dim RS As New Recordset
       Dim unitcode As String
       unitcode = Trim(txtUnitCode.Text)
       
       Conn.Open strcn.Connection_String
       
       Set cmd.ActiveConnection = Conn
       
       cmd.CommandType = adCmdText
       cmd.CommandText = "Delete from Project where prj_code='" + unitcode + "'"
       RS.CursorLocation = adUseClient
       RS.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
       
        Call GrdData
        
       Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

Private Sub cmdEXIT_Click()
    Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdSAVE_Click()
    If Len(Trim(txtUnitCode.Text)) = 0 Then
       MsgBox "Unit code required", vbCritical
       txtUnitCode.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txtUnitName.Text)) = 0 Then
       MsgBox "Unit name required", vbCritical
       txtUnitName.SetFocus
       Exit Sub
    End If
    
    Call SaveProject
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
    Call GrdData
    Call cmdADD_Click
End Sub

Private Sub GrdData()
    On Error GoTo err_loop
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select prj_code as Code,prj_name as Name from project"
    Adodc1.Refresh
    
    DataGrid1.Columns(0).Width = 1154.835
    DataGrid1.Columns(1).Width = 4259.906
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

Private Sub DataGrid1_Click()
    txtUnitCode.Text = DataGrid1.Columns(0).Text
    txtUnitName.Text = DataGrid1.Columns(1).Text
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub Form_Load()
    Call GrdData
End Sub

Private Sub txtUnitCode_LostFocus()

    On Error Resume Next

    Dim Conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset
    Dim prj_co As String
    prj_code = Trim(txtUnitCode.Text)
    
    Conn.Open strcn.Connection_String
    Set cmd.ActiveConnection = Conn
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "Select prj_code,prj_name from project where prj_code='" + prj_code + "'"
    cmd.Properties("IRowsetChange") = True
    cmd.Properties("Updatability") = 7
                    
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
  
    If Not (RS.EOF And RS.BOF) Then
        txtUnitName.Text = RS("Prj_name")

    Else
        txtUnitName.Text = ""
    End If
End Sub
