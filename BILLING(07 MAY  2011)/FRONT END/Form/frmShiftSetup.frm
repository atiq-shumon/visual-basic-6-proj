VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmShiftSetup 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   6330
      TabIndex        =   13
      Top             =   3510
      Width           =   1215
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   5100
      TabIndex        =   12
      Top             =   3510
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3465
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7815
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   0
         TabIndex        =   10
         Top             =   -120
         Width           =   7815
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SHIFT  INFORMATION SETUP"
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
            Left            =   1740
            TabIndex        =   11
            Top             =   180
            Width           =   4755
         End
         Begin VB.Image Image1 
            Height          =   855
            Left            =   -3030
            Picture         =   "frmShiftSetup.frx":0000
            Top             =   30
            Width           =   11820
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   3120
         Top             =   4110
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
      Begin VB.TextBox txtShiftName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   225
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1050
         Width           =   2325
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmShiftSetup.frx":5982
         Height          =   2085
         Left            =   225
         TabIndex        =   6
         Top             =   1320
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3678
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   -2147483624
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   19
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
      Begin MSComCtl2.DTPicker effective_date 
         Height          =   285
         Left            =   6030
         TabIndex        =   3
         Top             =   1050
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   38049
      End
      Begin MSComCtl2.DTPicker start_time 
         Height          =   285
         Left            =   2550
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   49152
         Format          =   60882946
         CurrentDate     =   37114
      End
      Begin MSComCtl2.DTPicker end_time 
         Height          =   285
         Left            =   4350
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   49152
         Format          =   60882946
         CurrentDate     =   37114
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Time"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2670
         TabIndex        =   9
         Top             =   810
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Time"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4350
         TabIndex        =   8
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6030
         TabIndex        =   7
         Top             =   810
         Width           =   1065
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   270
         TabIndex        =   5
         Top             =   810
         Width           =   885
      End
   End
   Begin VB.Shape Shape3 
      Height          =   465
      Left            =   5040
      Top             =   3450
      Width           =   2535
   End
End
Attribute VB_Name = "frmShiftSetup"
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
    Adodc1.RecordSource = "select shift_name,shift_start_time,shift_end_time,effective_date from Shift_setup"
    Adodc1.Refresh

End Sub

Private Sub cmdPreview_Click()

End Sub

Private Sub cmdSave_Click()
If Me.txtShiftName = "" Then
MsgBox "Shift Name Required", vbInformation, " IT, DNMIH"
txtShiftName.SetFocus
Exit Sub
End If
Call save_shift_setup
Call flush_grid
MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
txtShiftName.Text = ""

End Sub
Private Sub save_shift_setup()
Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
 If conn.State = 0 Then
    conn.Open strcn.Connection_String
 End If
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 30, txtShiftName.Text)
    cmd.Parameters.Append Param1 'Shift name
    
   
    Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, start_time)
    cmd.Parameters.Append Param2 'Start_time
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, end_time)
    cmd.Parameters.Append Param3 'End_time

    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 12, effective_date.Value)
    cmd.Parameters.Append Param4 'Effective DATE
    


   Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 10, "bo")
    cmd.Parameters.Append Param5 'u_id
    
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL save_shift_setup(?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set rs = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    If conn.State = 1 Then
       conn.Close
    End If
    
End Sub

Private Sub DTPicker3_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub DataGrid1_Click()
If Adodc1.Recordset.RecordCount = 0 Then
 MsgBox "Nothing to show", vbInformation, "Warning: IT, DNMIH"
 
 Else
 
        txtShiftName.Text = DataGrid1.Columns(0)
        start_time.Value = DataGrid1.Columns(1)
        end_time.Value = DataGrid1.Columns(2)
        effective_date.Value = DataGrid1.Columns(3)
    End If
End Sub

Private Sub Form_Load()
' txtInregExtraBed = frmExtraBed.txtRegNoExtraBed.Text
   Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select shift_name,shift_start_time,shift_end_time,effective_date from Shift_setup"
    Adodc1.Refresh
'            cmd.Properties("iRowsetChange") = True
'        cmd.Properties("updatability") = 7
'        rs2.CursorLocation = adUseClient

'        rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
'        If rs2.RecordCount > 0 Then
'         txtNameRelease = rs2!pat_name
''         TxtAgeRelease = rs2!age
'         comSexRelease.Text = rs2!sex
'         comDepartmentRelease = rs2!doc_dept
         
'rs2.Close
'Conn2.Close
'Else
' MsgBox "Invalid Registration No", vbInformation, "Warning: IT, DNMIH"
' rs2.Close
' Conn2.Close
' Exit Sub
' Unload Me

'End If

End Sub

