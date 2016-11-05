VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSecurity 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   10440
      TabIndex        =   21
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   9210
      TabIndex        =   20
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "NEW"
      Height          =   375
      Left            =   7980
      TabIndex        =   19
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   6750
      TabIndex        =   18
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5595
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11745
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   705
         Left            =   0
         TabIndex        =   16
         Top             =   -90
         Width           =   11745
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "NEW  USER  INFORMATION SETUP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0FF&
            Height          =   315
            Left            =   3660
            TabIndex        =   17
            Top             =   240
            Width           =   4755
         End
         Begin VB.Image Image1 
            Height          =   855
            Left            =   -120
            Picture         =   "frmSecurity.frx":0000
            Stretch         =   -1  'True
            Top             =   -60
            Width           =   12330
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmSecurity.frx":5982
         Height          =   4095
         Left            =   0
         TabIndex        =   9
         Top             =   1470
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   -2147483624
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
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
         Left            =   4890
         Top             =   4410
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
         Caption         =   "Adodc2"
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
      Begin VB.ComboBox cboShift 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc2"
         Height          =   315
         ItemData        =   "frmSecurity.frx":5997
         Left            =   9690
         List            =   "frmSecurity.frx":5999
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1140
         Width           =   1080
      End
      Begin VB.ComboBox cboUsrType_ 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSecurity.frx":599B
         Left            =   8550
         List            =   "frmSecurity.frx":59A8
         TabIndex        =   4
         Text            =   "Admin"
         Top             =   1140
         Width           =   1170
      End
      Begin VB.TextBox txtuser_c_pass 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6570
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1140
         Width           =   1995
      End
      Begin VB.TextBox txtuser_pass 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4470
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1140
         Width           =   2115
      End
      Begin VB.TextBox txtuser_id 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   510
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1140
         Width           =   765
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   3120
         Top             =   4350
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
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
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1140
         Width           =   3225
      End
      Begin MSComCtl2.DTPicker effective_date 
         Height          =   285
         Left            =   10770
         TabIndex        =   8
         Top             =   1140
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   38049
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9810
         TabIndex        =   15
         Top             =   900
         Width           =   390
      End
      Begin VB.Label cboUsrType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8640
         TabIndex        =   14
         Top             =   900
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4890
         TabIndex        =   13
         Top             =   900
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6720
         TabIndex        =   12
         Top             =   900
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User id"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   11
         Top             =   900
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   10470
         TabIndex        =   10
         Top             =   900
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   15
         Top             =   780
         Width           =   11715
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2490
         TabIndex        =   7
         Top             =   900
         Width           =   480
      End
   End
   Begin VB.Shape Shape3 
      Height          =   465
      Left            =   6690
      Top             =   5580
      Width           =   4995
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Con As New MyConnection
Dim Conn2 As New Connection
Dim rs2 As New Recordset
Dim cmd As New Command

Private Sub CMDEXIT_Click()

    Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If

End Sub

Private Sub flush_grid()

    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select user_id,user_name,user_type,shift_name,dt from Security"
    Adodc1.Refresh
    DataGrid1.Columns(0).Width = 800
    DataGrid1.Columns(1).Width = 5000
    DataGrid1.Columns(2).Width = 2000
    DataGrid1.Columns(3).Width = 2000
    DataGrid1.Columns(4).Width = 2000
'    DataGrid1.Columns(5).Width = 100
    

End Sub

Private Sub cmdSave_Click()
If Me.txtuser_id = "" Then
MsgBox "User id Required", vbInformation, " IT, DNMIH"
txtuser_id.SetFocus
Exit Sub
End If
If Me.txtUsername = "" Then
MsgBox "User Name Required", vbInformation, " IT, DNMIH"
txtUsername.SetFocus
Exit Sub
End If
If Me.txtuser_pass = "" Then
MsgBox "User password Required", vbInformation, " IT, DNMIH"
txtuser_pass.SetFocus
Exit Sub
End If
If Me.txtuser_c_pass = "" Then
MsgBox "User confirm  password Required", vbInformation, " IT, DNMIH"
txtuser_c_pass.SetFocus
Exit Sub
End If

Call save_security_setup
Call flush_grid
MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."

End Sub

Private Sub save_security_setup()
    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
 If conn.State = 0 Then
    conn.Open strcn.Connection_String
 End If
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 30, Me.txtuser_id.Text)
    cmd.Parameters.Append Param1 'user_id name
    
   
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 50, Me.txtUsername.Text)
    cmd.Parameters.Append Param2 'user_name
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 30, Me.txtuser_pass.Text)
    cmd.Parameters.Append Param3 'pass name
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 30, Me.txtuser_c_pass)
    cmd.Parameters.Append Param4 'user_c pass
    


   Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 20, Me.cboUsrType_.Text)
    cmd.Parameters.Append Param5 'type
    
   Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 20, cboShift.Text)
    cmd.Parameters.Append Param6 'type
    
    
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL save_security(?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set rs = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    
    If conn.State = 1 Then
       conn.Close
    End If
End Sub

Private Sub DataGrid1_Click()
 If Adodc1.Recordset.RecordCount = 0 Then
 MsgBox "Nothing to show", vbInformation, "Warning: IT, DNMIH"
 Else
Me.txtuser_id = DataGrid1.Columns(0)
Me.txtUsername = DataGrid1.Columns(1)
cboUsrType_ = DataGrid1.Columns(2)
'''Me.cboShift = DataGrid1.Columns(3)
End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    SendKeys Chr(9)
End If


End Sub

Private Sub Form_Load()

'   Adodc1.ConnectionString = strcn.Connection_String
'    Adodc1.RecordSource = "select user_id,user_name,user_type,dt  from Security"
'    Adodc1.Refresh

    
    
          Adodc2.ConnectionString = strcn.Connection_String
      Adodc2.RecordSource = "select distinct(Shift_name) from Shift_setup"
      Adodc2.Refresh

    If Adodc2.Recordset.RecordCount > 0 Then
        Adodc2.Recordset.MoveFirst
        While Adodc2.Recordset.EOF = False
       cboShift.AddItem Adodc2.Recordset!shift_name
            Adodc2.Recordset.MoveNext
        Wend
     End If
     Call flush_grid
     
      
End Sub

