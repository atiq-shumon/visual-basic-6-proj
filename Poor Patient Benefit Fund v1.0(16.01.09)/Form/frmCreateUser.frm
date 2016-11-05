VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New User "
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmCreateUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H80000001&
      Height          =   705
      Left            =   -90
      TabIndex        =   19
      Top             =   -120
      Width           =   6555
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add New User"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4470
         TabIndex        =   20
         Top             =   180
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1725
      Left            =   -90
      TabIndex        =   14
      Top             =   90
      Width           =   6765
      Begin VB.TextBox txtPass 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   3615
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1110
         Width           =   2730
      End
      Begin VB.TextBox txtu_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   930
         MaxLength       =   50
         TabIndex        =   0
         Top             =   690
         Width           =   1455
      End
      Begin VB.ComboBox cboType 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   930
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1110
         Width           =   1455
      End
      Begin VB.TextBox txtUser_name 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   1
         Top             =   690
         Width           =   2730
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   225
         Left            =   2745
         TabIndex        =   18
         Top             =   1140
         Width           =   705
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   225
         Left            =   180
         TabIndex        =   17
         Top             =   690
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   240
         Left            =   390
         TabIndex        =   16
         Top             =   1170
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   225
         Left            =   2640
         TabIndex        =   15
         Top             =   720
         Width           =   840
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         Height          =   315
         Left            =   885
         Top             =   660
         Width           =   1515
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         Height          =   345
         Left            =   885
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFC0C0&
         Height          =   285
         Left            =   3585
         Top             =   660
         Width           =   2805
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFC0C0&
         Height          =   285
         Left            =   3585
         Top             =   1080
         Width           =   2805
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   915
      Left            =   -60
      TabIndex        =   13
      Top             =   1710
      Width           =   6555
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   5130
         Picture         =   "frmCreateUser.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdInsert 
         Height          =   495
         Left            =   3930
         Picture         =   "frmCreateUser.frx":1EC4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1185
      End
      Begin VB.Shape Shape5 
         Height          =   555
         Left            =   3900
         Top             =   210
         Width           =   2475
      End
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   465
      Index           =   2
      Left            =   4170
      Picture         =   "frmCreateUser.frx":3856
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4590
      Width           =   1125
   End
   Begin VB.CommandButton cmdEXIT 
      Height          =   465
      Left            =   5340
      Picture         =   "frmCreateUser.frx":5260
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4590
      Width           =   1215
   End
   Begin VB.CommandButton cmdADD 
      Height          =   465
      Index           =   3
      Left            =   2910
      Picture         =   "frmCreateUser.frx":6CE2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4590
      Width           =   1185
   End
   Begin VB.CommandButton cmdSAVE 
      Height          =   465
      Left            =   1710
      Picture         =   "frmCreateUser.frx":8674
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4590
      Width           =   1185
   End
   Begin VB.CommandButton cmdPermit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Permission"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2430
      Width           =   1140
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3570
      Top             =   2430
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
      Height          =   2280
      Left            =   135
      TabIndex        =   7
      Top             =   2235
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4022
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3600
      Top             =   2310
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
      Left            =   3570
      Top             =   2160
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Access Information "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   330
      Left            =   120
      TabIndex        =   8
      Top             =   2130
      Width           =   3270
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim Temp_Tab As ADODB.Recordset
'Dim Temp_Tab_Helper As New ADODB.Recordset
Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub
Private Sub cmdADD_Click(Index As Integer)
  
    txtUser_name.Text = ""
    txtPass.Text = ""
'    Temp_rst
    txtu_id.SetFocus

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDELETE_Click(Index As Integer)
     If Len(Trim(txtu_id.Text)) = 0 Then Exit Sub
    If Trim(txtu_id.Text) = UCase(strUid) Then
       MsgBox "You can't delete yourself", vbCritical
       Exit Sub
    End If
End Sub
Private Sub cmdInsert_Click()
    If Len(Trim(txtu_id.Text)) = 0 Then Exit Sub
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "Select user_id,user_name,pass_word from security where user_id='" & Trim(txtu_id.Text) & "'"
    Adodc1.Refresh
    
    If Adodc1.Recordset.RecordCount > 0 Then
       MsgBox "Duplicate Id not allowed", vbCritical, "Warning..."
       txtu_id.SetFocus
       Exit Sub
    End If
    Call SaveSecurity
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
    txtUser_name.Text = ""
    txtPass.Text = ""
    txtu_id.Text = ""
    txtu_id.SetFocus
End Sub
Private Sub cmdPermit_Click()
'    Flushx
'    DataGrid1.SetFocus
End Sub
Private Sub Form_Load()
    cboType.AddItem "ADMIN"
    cboType.AddItem "USER"
    cboType.AddItem "SUPER USER"
    cboType.Text = "USER"
   
End Sub
Private Sub SaveSecurity()
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim mode As String
     mode = "2"
     
    Conn.Open strcn.Connection_String
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, "2")
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 15, Trim(txtu_id.Text))
    cmd.Parameters.Append Param2

    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 20, txtUser_name.Text)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 15, txtPass.Text)
    cmd.Parameters.Append Param4
    
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL SaveSecurity(?, ?, ?,?)}"
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
End Sub
Private Sub txtPass_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub
Private Sub txtu_id_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     If KeyAscii = 13 Then
        SendKeys Chr(9)
     End If
End Sub

Private Sub txtu_id_LostFocus()
    If Len(Trim(txtu_id.Text)) = 0 Then Exit Sub
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "Select user_id,user_name,pass_word from security where user_id='" & Trim(txtu_id.Text) & "'"
    Adodc1.Refresh
    
    If Adodc1.Recordset.RecordCount > 0 Then
       txtUser_name.Text = "" & Adodc1.Recordset!user_name
'       txtPass.Text = " " & Adodc1.Recordset!Pass_word
       
    Else
       txtUser_name.Text = ""
       txtPass.Text = ""
       txtUser_name.SetFocus
       
    End If
End Sub
Private Sub txtUser_name_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
     If KeyAscii = 13 Then
        SendKeys Chr(9)
     End If
End Sub

