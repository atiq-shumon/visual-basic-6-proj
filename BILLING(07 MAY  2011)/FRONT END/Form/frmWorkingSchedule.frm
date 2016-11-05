VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmWorkingSchedule 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   7860
      Width           =   13365
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Software Programmer, IT Division, DNMIH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2940
         TabIndex        =   20
         Top             =   60
         Width           =   4725
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Developed && Maintenanced by:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   150
         TabIndex        =   19
         ToolTipText     =   "Developed && Maintenanced by:A.N.M. Atiqur Rahman,Software Programmer, IT Division, DNMIH"
         Top             =   90
         Width           =   2715
      End
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   3810
      TabIndex        =   3
      ToolTipText     =   "CLICK TO SAVE "
      Top             =   7350
      Width           =   1215
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "NEW"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      ToolTipText     =   "CLICK TO CLEAR"
      Top             =   7350
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   6270
      TabIndex        =   6
      ToolTipText     =   "CLICK TO DELETE"
      Top             =   7350
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   7500
      TabIndex        =   7
      ToolTipText     =   "CLICK TO VIEW REPORT"
      Top             =   7350
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   8730
      TabIndex        =   5
      ToolTipText     =   "CLICK TO CLOSE"
      Top             =   7350
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   7335
      Left            =   0
      TabIndex        =   8
      Top             =   -90
      Width           =   10095
      Begin VB.TextBox TxtName 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   1830
         TabIndex        =   16
         Top             =   1230
         Width           =   4335
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   915
         Left            =   0
         TabIndex        =   14
         Top             =   -30
         Width           =   10125
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Roster Duty  Setup"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0FF&
            Height          =   465
            Left            =   2910
            TabIndex        =   15
            Top             =   240
            Width           =   3330
         End
         Begin VB.Image Image1 
            Height          =   855
            Left            =   -120
            Picture         =   "frmWorkingSchedule.frx":0000
            Stretch         =   -1  'True
            Top             =   30
            Width           =   10710
         End
      End
      Begin VB.ComboBox cboUser_id 
         Height          =   315
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1230
         Width           =   1155
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   2550
         Top             =   7020
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
      Begin VB.ComboBox Cbo_Shift_name 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1230
         Width           =   2475
      End
      Begin MSComCtl2.DTPicker effective_date 
         Height          =   315
         Left            =   8610
         TabIndex        =   2
         Top             =   1230
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         OLEDropMode     =   1
         Format          =   58916865
         CurrentDate     =   38049
      End
      Begin VB.Frame Frame2 
         Height          =   5835
         Left            =   270
         TabIndex        =   12
         Top             =   1470
         Width           =   9855
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmWorkingSchedule.frx":5982
            Height          =   5655
            Left            =   60
            TabIndex        =   13
            Top             =   150
            Width           =   9645
            _ExtentX        =   17013
            _ExtentY        =   9975
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   -2147483624
            ForeColor       =   12582912
            HeadLines       =   1
            RowHeight       =   17
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
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USER NAME"
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
         Left            =   3210
         TabIndex        =   17
         Top             =   1020
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EFFECTIVE DATE"
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
         Left            =   8550
         TabIndex        =   11
         Top             =   1020
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ASSIGNED SHIFT"
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
         Left            =   6600
         TabIndex        =   10
         Top             =   1020
         Width           =   1380
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "USER ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   780
         TabIndex        =   9
         Top             =   1020
         Width           =   735
      End
   End
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   3750
      Top             =   7290
      Width           =   6225
   End
End
Attribute VB_Name = "frmWorkingSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboUser_id_Click()
  flush_grid (3)
  load_name
End Sub
Private Sub load_name()
  Dim Con As New ADODB.Connection
  Dim RS As New ADODB.Recordset
  Dim cmd As New ADODB.Command
  
  Con.ConnectionString = strcn.Connection_String
  If Con.State = 0 Then
   Con.Open
  End If
  Set cmd.ActiveConnection = Con
  cmd.CommandType = adCmdText
  cmd.CommandText = "select user_name  from SECURITY  where to_char(user_id)='" & cboUser_id & "'"
  Set RS = cmd.Execute
  If Not RS.EOF Then
    TxtName = RS(0)
  End If
  
  Set Con = Nothing
  Set RS = Nothing
  Set cmd = Nothing

End Sub

Private Sub CMDDELETE_Click()
If cboUser_id = "" Then
     MsgBox "User ID Requied"
        cboUser_id.SetFocus
        Exit Sub
   End If
   
If Cbo_Shift_name = "" Then
     MsgBox "Shift Name Requied"
        Cbo_Shift_name.SetFocus
        Exit Sub
   End If
       
    Call DELETEWorkingSchedule
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Delete..."
'    Call FlushCompSetup
Call flush_grid(1)
'Call clear


End Sub
Private Sub DELETEWorkingSchedule()
Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
       
   If Conn.State = 0 Then
       Conn.Open strcn.Connection_String
   End If
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 20, cboUser_id.Text)
    cmd.Parameters.Append Param1 'user_id


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 40, Cbo_Shift_name.Text)
    cmd.Parameters.Append Param2 'shift_name
    
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 12, effective_date.Value)
    cmd.Parameters.Append Param3 'Tmp Date default sysdate


    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL DELETE_working_schedule(?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    

End Sub
Private Sub cmdExit_Click()
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdSave_Click()
If cboUser_id = "" Then
     MsgBox "User ID Requied"
        cboUser_id.SetFocus
        Exit Sub
   End If
   
If Cbo_Shift_name = "" Then
     MsgBox "Shift Name Requied"
        Cbo_Shift_name.SetFocus
        Exit Sub
   End If
       
    Call SaveWorkdingSchedule
    MsgBox "Operation successful", vbInformation + vbOKOnly, "Save..."
'    Call FlushCompSetup
Call flush_grid(1)
'Call clear

End Sub
Private Sub SaveWorkdingSchedule()
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
       
  If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
  End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 20, cboUser_id.Text)
    cmd.Parameters.Append Param1 'user_id


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 40, Cbo_Shift_name.Text)
    cmd.Parameters.Append Param2 'shift_name
    
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 12, effective_date.Value)
    cmd.Parameters.Append Param3 'Tmp Date default sysdate


    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL save_working_schedule(?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
       Conn.Close
    End If

End Sub
Private Sub flush_grid(MODE As Integer)
  If MODE = 1 Then
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select A.user_id ID ,  B.user_name   Name,A.shift_name  Shift ,A.dt Working_Date  from working_schedule A,SECURITY B WHERE TO_CHAR(A.DT,'DD-MON-YYYY')=TO_CHAR(SYSDATE,'DD-MON-YYYY') AND TO_CHAR(A.user_id)=TO_CHAR(B.user_id) order by A.dt,A.shift_name desc "
    Adodc1.Refresh
 ElseIf MODE = 2 Then
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select A.user_id ID ,  B.user_name Name,A.shift_name Shift,A.dt Working_Date from working_schedule A,SECURITY B WHERE TO_CHAR(A.DT,'DD/MM/YYYY')='" & effective_date & "' AND to_char(A.USER_ID)=to_char(B.USER_ID) order by A.dt,A.shift_name desc "
    Adodc1.Refresh
  ElseIf MODE = 3 Then
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select A.user_id ID ,  B.user_name Name,A.shift_name Shift,A.dt Working_Date from working_schedule A,SECURITY B WHERE TO_CHAR(A.DT,'DD/MM/YYYY')='" & effective_date & "' AND TO_CHAR(A.user_id)='" & cboUser_id & "' AND TO_CHAR(A.user_id)=TO_CHAR(B.user_id) order by A.dt,A.shift_name desc "
    Adodc1.Refresh
   
  End If
        
    DataGrid1.Columns(0).Width = 1250
    DataGrid1.Columns(1).Width = 4300
    DataGrid1.Columns(2).Width = 2500
   
End Sub

Private Sub DataGrid1_Click()

 If Adodc1.Recordset.RecordCount = 0 Then
 MsgBox "Nothing to show", vbInformation, "Warning: IT, DNMIH"
 
 Else
'cboUser_id = DataGrid1.Columns(0)
Cbo_Shift_name = DataGrid1.Columns(2)
effective_date = DataGrid1.Columns(3)
End If

End Sub

Private Sub effective_date_Change()
    flush_grid (2)
End Sub

Private Sub effective_date_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     cmdSAVE.SetFocus
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys Chr(9)
End If
End Sub

Private Sub Form_Load()
             Adodc1.ConnectionString = strcn.Connection_String
            Adodc1.RecordSource = "select TO_NUMBER(user_id) user_id from security ORDER BY USER_ID"
            Adodc1.Refresh

        If Adodc1.Recordset.RecordCount > 0 Then
                Adodc1.Recordset.MoveFirst
            While Adodc1.Recordset.EOF = False
                    cboUser_id.AddItem Adodc1.Recordset!user_id
                   Adodc1.Recordset.MoveNext
                Wend
            cboUser_id.Text = cboUser_id.List(0)
            
        End If
  

            
   
            Adodc1.ConnectionString = strcn.Connection_String
            Adodc1.RecordSource = "select distinct(Shift_name) from Shift_setup"
            Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        While Adodc1.Recordset.EOF = False
        Cbo_Shift_name.AddItem Adodc1.Recordset!shift_name
            Adodc1.Recordset.MoveNext
        Wend
        Cbo_Shift_name.Text = Cbo_Shift_name.List(0)
    End If
    effective_date = Date
   flush_grid (1)
End Sub

