VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7050
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Last Name"
      Height          =   285
      Index           =   2
      Left            =   3990
      TabIndex        =   5
      Top             =   30
      Width           =   1065
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Mid Name"
      Height          =   285
      Index           =   1
      Left            =   2910
      TabIndex        =   4
      Top             =   30
      Width           =   1065
   End
   Begin VB.OptionButton Option1 
      Caption         =   "First Name"
      Height          =   285
      Index           =   0
      Left            =   1830
      TabIndex        =   3
      Top             =   30
      Value           =   -1  'True
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   630
      Top             =   1890
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
      Bindings        =   "frmFind.frx":0000
      Height          =   5805
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   10239
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      Enabled         =   -1  'True
      ForeColor       =   12582912
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
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   105
      TabIndex        =   1
      Top             =   390
      Width           =   6915
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Find"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   300
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim objCom As New DSLComFram.clsCommon

Public objFindRS As New ADODB.Recordset
Public strfrmCaption As String
Public intInputsel As Integer
Public SQLString As String
'Public intInputsel1 As Integer
'Public intInputsel2 As Integer
'Public intInputsel3 As Integer
'Public intInputsel4 As Integer
'Public intInputsel5 As Integer

Public OwnerForm As Form

Private Sub DataGrid1_DblClick()
  MsgBox "Please Press on Enter to make sure selection", vbInformation, cmp
  Exit Sub
'On Error GoTo errdes
'OwnerForm.txtfields(intInputsel).SetFocus
'If Not IsEmpty(intInputsel1) Then OwnerForm.txtFields(intInputsel1) = DataGrid1.Columns(1).Text
'If Not IsEmpty(intInputsel2) Then OwnerForm.txtFields(intInputsel2) = DataGrid1.Columns(2).Text
'If Not IsEmpty(intInputsel3) Then OwnerForm.txtFields(intInputsel1) = DataGrid1.Columns(3).Text
'If Not IsEmpty(intInputsel4) Then OwnerForm.txtFields(intInputsel2) = DataGrid1.Columns(4).Text
'If Not IsEmpty(intInputsel5) Then OwnerForm.txtFields(intInputsel1) = DataGrid1.Columns(5).Text
'Exit Sub
'errdes:

End Sub

'Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'    DataGrid1_DblClick
'End If
'End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    OwnerForm.txtfields(intInputsel) = DataGrid1.Columns(0).Text
'   OwnerForm.Label12(intInputsel) = DataGrid1.Columns(1).Text
   Unload Me
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
         
End Sub

Private Sub Form_Load()
Adodc1.connectionstring = GConnString
Adodc1.RecordSource = SQLString ' "select Studentid as [Student ID],Studentname as [Student name] from Studentinfo"
Adodc1.Refresh

DataGrid1.Columns(0).Width = 1500
DataGrid1.Columns(1).Width = 4500

'Dim rs As New adodb.Recordset
'Set rs = objCom.Get_RS("select Studentid as [Student ID],Studentname as [Student name] from Studentinfo", objmyCon)
'Set rs = GetData("select Studentid as [Student ID],Studentname as [Student name] from Studentinfo")
'Set DataGrid1.DataSource = rs
'DataGrid1.Refresh
End Sub

Private Sub Option1_LostFocus(Index As Integer)
  Select Case Index
          Case 0, 1, 2
            txtFind.SetFocus
  End Select
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Set frmFind = Nothing
'End Sub
'
'Private Sub txtFind_Change()
'Dim objrs As New adodb.Recordset
''Set objrs = objCom.Get_RS("select Studentid as [Student ID],Studentname as [Student name] from Studentinfo", objmyCon)
'Set objrs = GetData("select Studentid as [Student ID],Studentname as [Student name] from Studentinfo")
'Set DataGrid1.DataSource = objrs
'DataGrid1.Refresh
'Set objrs = Nothing
'
'End Sub
'
'Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'    DataGrid1.SetFocus
'End If
'End Sub
'Public Function GetData(sqlstring As String) As adodb.Recordset
'Dim cmd As New adodb.command
'Dim con As New adodb.connection
'Dim rs As New adodb.Recordset
'con.Open GConnString
'Set cmd.ActiveConnection = con
'    cmd.CommandType = adCmdText
'    cmd.CommandText = sqlstring
'
' Set rs = cmd.Execute
'Set GetData = rs
'End Function
Private Sub txtFind_Change()
If Option1(0).Value = True Then
  Adodc1.connectionstring = GConnString
  Adodc1.RecordSource = "select Studentid as [Student ID],Studentname as [Student name] from Studentinfo where studentname like '" & Trim(txtFind) & "%'"
  Adodc1.Refresh
ElseIf Option1(1).Value = True Then
  Adodc1.connectionstring = GConnString
  Adodc1.RecordSource = "select Studentid as [Student ID],Studentname as [Student name] from Studentinfo where studentname like '%" & Trim(txtFind) & "%'"
  Adodc1.Refresh

ElseIf Option1(2).Value = True Then
   Adodc1.connectionstring = GConnString
   Adodc1.RecordSource = "select Studentid as [Student ID],Studentname as [Student name] from Studentinfo where studentname like '%" & Trim(txtFind) & "%'"
   Adodc1.Refresh
End If
DataGrid1.Columns(0).Width = 1500
DataGrid1.Columns(1).Width = 4500


  
  
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DataGrid1.SetFocus
End If
End Sub
