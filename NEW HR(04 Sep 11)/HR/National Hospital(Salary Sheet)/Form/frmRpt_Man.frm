VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form27 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Manager"
   ClientHeight    =   3630
   ClientLeft      =   2745
   ClientTop       =   3195
   ClientWidth     =   6690
   ForeColor       =   &H00000080&
   Icon            =   "frmRpt_Man.frx":0000
   LinkTopic       =   "Form27"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6690
   Begin MSDataGridLib.DataGrid DataGrid0 
      Height          =   75
      Left            =   585
      TabIndex        =   7
      Top             =   765
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   132
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   17
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
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
         Name            =   "Arial"
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   0
            ColumnWidth     =   2204.788
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbRpt_Month 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      ItemData        =   "frmRpt_Man.frx":08CA
      Left            =   1575
      List            =   "frmRpt_Man.frx":08F2
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2250
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cmbRpt_Year 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      ItemData        =   "frmRpt_Man.frx":0958
      Left            =   2880
      List            =   "frmRpt_Man.frx":0983
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2250
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   675
      Top             =   2970
   End
   Begin VB.TextBox txtRpt_ID 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   1575
      TabIndex        =   1
      Top             =   855
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.ListBox lstParam 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1290
      Left            =   1575
      TabIndex        =   2
      Top             =   855
      Width           =   2130
   End
   Begin VB.ComboBox cmbChooser_1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      ItemData        =   "frmRpt_Man.frx":09D5
      Left            =   1575
      List            =   "frmRpt_Man.frx":09E8
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   405
      Width           =   2130
   End
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   1575
      ScaleHeight     =   570
      ScaleWidth      =   2145
      TabIndex        =   8
      Top             =   2775
      Width           =   2175
      Begin VB.CommandButton Command1 
         Height          =   645
         Index           =   1
         Left            =   1440
         Picture         =   "frmRpt_Man.frx":0A4D
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "   Close   "
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Height          =   645
         Index           =   2
         Left            =   720
         Picture         =   "frmRpt_Man.frx":0D57
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "  Choose Printer  "
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Height          =   645
         Index           =   0
         Left            =   0
         Picture         =   "frmRpt_Man.frx":1061
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "   View Report   "
         Top             =   0
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   405
      Top             =   3105
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1125
      Top             =   3105
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
      Left            =   630
      Top             =   3150
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   135
      Top             =   3150
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
   Begin VB.Image Image1 
      Height          =   375
      Left            =   5715
      Picture         =   "frmRpt_Man.frx":136B
      Stretch         =   -1  'True
      ToolTipText     =   "  Payroll Report  "
      Top             =   2970
      Width           =   420
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFE2D7&
      BorderColor     =   &H00FF8080&
      Height          =   1725
      Left            =   4050
      Top             =   405
      Width           =   2355
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   4185
      Picture         =   "frmRpt_Man.frx":1C35
      Stretch         =   -1  'True
      ToolTipText     =   "  Employee Information Report  "
      Top             =   2970
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   4950
      Picture         =   "frmRpt_Man.frx":24FF
      Stretch         =   -1  'True
      ToolTipText     =   "  Payroll Report  "
      Top             =   3015
      Width           =   420
   End
   Begin VB.Label lblComp_Nm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ayman Textile && Hosiery Ltd."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4050
      TabIndex        =   12
      Top             =   2475
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label lblMon_Yr 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Month / Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   540
      TabIndex        =   11
      Top             =   2295
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblParam 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   540
      TabIndex        =   10
      Top             =   870
      Width           =   45
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Report Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   540
      TabIndex        =   9
      Top             =   450
      Width           =   915
   End
End
Attribute VB_Name = "Form27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Dim Rpt_sno As String
''
''Private Sub cmbChooser_1_Change()
''Change_Chooser
''End Sub
''
''Private Sub cmbChooser_1_Click()
''    Change_Chooser
''End Sub
''
''Private Sub cmbChooser_1_KeyPress(KeyAscii As Integer)
''    If KeyAscii <> 13 Then Exit Sub
''    Change_Chooser
''End Sub
''
''Private Sub Command1_Click(Index As Integer)
''On Error Resume Next
''
''    Dim res
''Select Case Index
''    Case 0        ''View Report Button
''
''    ''--------------------------------
''        Rpt_Mode = Rpt_No + Rpt_sno
''    ''--------------------------------
''
''        Rpt_ID = Trim(txtRpt_ID)
''        Rpt_Desig = lstParam
''        Rpt_Dept = lstParam
''        Rpt_Sec = lstParam
''        Rpt_Month = cmbRpt_Month
''        Rpt_Year = cmbRpt_Year
''        Rpt_BldGr = lstParam
''
''        If Rpt_ID <> Empty Then Photo_Path = Get_Photo(Rpt_ID)
''
''        Form26.Show vbModal
''
''    Case 1          ''Close Button
''
''        res = MsgBox("Do you really want to close it?", vbQuestion + vbYesNo, "Report Manager")
''        If res = vbNo Then
''            Exit Sub
''        Else
''            Unload Me
''        End If
''    Case 2
''        Form1.CommonDialog1.Action = 5
''End Select
''End Sub
''
''Private Sub Command2_Click()
''Form1.CommonDialog1.Action = 5
''End Sub
''
''Private Sub DataGrid0_Click()
''On Error Resume Next
''    txtRpt_ID = Adodc3.Recordset!Emp_id
''End Sub
''
''Public Sub Append_Desig_Dept_ID()
''
''Select Case cmbChooser_1
''
''    Case "Employee specific", "Pay slip", "Salary Statement(July-June)"
''            ''Brings Employee name and Id from Emp_Per_info table
''
''            Adodc3.ConnectionString = strCN.Connection
''            Adodc3.RecordSource = "exec POP_Emp_detail_for_Payroll"
''            Adodc3.Refresh
''
''            If Adodc3.Recordset.RecordCount > 0 Then
''                Adodc3.Recordset.MoveLast
''            End If
''
''            Set DataGrid0.DataSource = Adodc3
''
''            DataGrid0.Columns(0).DataField = "emp_id"
''            DataGrid0.Columns(1).DataField = "NM"
''
''            DataGrid0.ReBind
''            DataGrid0.Refresh
''
''    Case "Designation specific"
''            ''Brings designation from Job title table
''            lstParam.Clear
''            Adodc1.ConnectionString = strCN.Connection
''            Adodc1.RecordSource = "select title from job_title "
''            Adodc1.Refresh
''            If Adodc1.Recordset.RecordCount > 0 Then
''                Do Until Adodc1.Recordset.EOF
''                    lstParam.AddItem Adodc1.Recordset!Title
''                    Adodc1.Recordset.MoveNext
''                Loop
''            End If
''
''
''
''    Case "Department specific", "Salary Sheet (Dept)", _
''    "Pay slip (Dept)", "Salary Information (Dept)", "Bank Statement (Dept)"
''
''            ''Brings Department name from Department_Info table
''            lstParam.Clear
''            Adodc2.ConnectionString = strCN.Connection
''            Adodc2.RecordSource = "select title from Department_Info"
''            Adodc2.Refresh
''
''            If Adodc2.Recordset.RecordCount > 0 Then
''
''              Do Until Adodc2.Recordset.EOF
''                    lstParam.AddItem Adodc2.Recordset!Title
''                    Adodc2.Recordset.MoveNext
''              Loop
''            End If
''    Case "Section specific", "Salary Sheet (Sec)"
''
''            ''Brings Department name from Department_Info table
''            lstParam.Clear
''            Adodc2.ConnectionString = strCN.Connection
''            Adodc2.RecordSource = "Select Title from sec_Info"
''            Adodc2.Refresh
''
''            If Adodc2.Recordset.RecordCount > 0 Then
''
''              Do Until Adodc2.Recordset.EOF
''                    lstParam.AddItem Adodc2.Recordset!Title
''                    Adodc2.Recordset.MoveNext
''              Loop
''            End If
''
''    Case "Job Type"
''
''            ''Brings Job type from Job_Type table
''            lstParam.Clear
''
''            Adodc4.ConnectionString = strCN.Connection
''            Adodc4.RecordSource = "select title from Job_Type"
''            Adodc4.Refresh
''
''            If Adodc4.Recordset.RecordCount > 0 Then
''
''              Do Until Adodc4.Recordset.EOF
''                    lstParam.AddItem Adodc4.Recordset!Title
''                    Adodc4.Recordset.MoveNext
''              Loop
''            End If
''
''    Case "Bloodgroup specific"
''            lstParam.Clear
''
''            Const LItm As Integer = 7
''            Dim Bld_gr(LItm) As String
''            Dim L As Integer
''
''            Bld_gr(0) = "A +ve": Bld_gr(1) = "A -ve": Bld_gr(2) = "B +ve": Bld_gr(3) = "B -ve"
''            Bld_gr(4) = "AB +ve": Bld_gr(5) = "AB -ve": Bld_gr(6) = "O +ve": Bld_gr(7) = "O -ve"
''
''            For L = 0 To LItm
''                lstParam.AddItem Bld_gr(L)
''            Next
''
''
''Case "Group specific"
''
''            ''Brings Group Specific from Group table
''            lstParam.Clear
''
''            Adodc4.ConnectionString = strCN.Connection
''            Adodc4.RecordSource = "select Gr_Name from Group_NM"
''            Adodc4.Refresh
''
''            If Adodc4.Recordset.RecordCount > 0 Then
''
''              Do Until Adodc4.Recordset.EOF
''                    lstParam.AddItem Adodc4.Recordset!gr_name
''                    Adodc4.Recordset.MoveNext
''              Loop
''            End If
''    Case "Shift specific"
''
''            ''Brings Group Specific from Group table
''            lstParam.Clear
''
''            Adodc2.ConnectionString = strCN.Connection
''            Adodc2.RecordSource = "select Shift_Name from Shift"
''            Adodc2.Refresh
''
''            If Adodc2.Recordset.RecordCount > 0 Then
''
''              Do Until Adodc2.Recordset.EOF
''                    lstParam.AddItem Adodc2.Recordset!Shift_Name
''                    Adodc2.Recordset.MoveNext
''              Loop
''            End If
''
''End Select
''
''    If lstParam.Visible = True Then
''        If Not lstParam.ListCount = 0 Then
''            lstParam.Selected(0) = True
''            lstParam.SetFocus
''        End If
''    End If
''End Sub
''
''Private Sub Form_Load()
''cmbRpt_Year = Year(Now)
''cmbRpt_Month = MonthName(Month(Now))
''cmbChooser_1.Clear
''''-------------------------------------------------------------------
''''  Add available Report items to cmbChooser_1 for any speific type
''''-------------------------------------------------------------------
''
''    Dim RItm As Integer
''    Dim Rpt(6) As String    ''-----------must put the max. number (i.e. 5)
''    Dim R As Integer        ''-----------of item added to the cmbChooser_1
''
''Select Case Rpt_No
''     Case "Per_Detail_"     ''-------------"Personal Information Report"
''
''            RItm = 5
''
''            Rpt(0) = "Employee specific"
''            Rpt(1) = "Designation specific"
''            Rpt(2) = "Department specific"
''            Rpt(3) = "Bloodgroup specific"
''            Rpt(4) = "All Employee"
''            Rpt(5) = "Job Type"
''
''    Case "Att_"           ''-------------"Attendance Report"
''            RItm = 3
''
''            Rpt(0) = "Employee specific"
''            Rpt(1) = "Department specific"
''            Rpt(2) = "Section specific"
''            Rpt(3) = "All Employee"
''
''    Case "Leave_"         ''--------------"Leave Report"
''
''            RItm = 0
''
''            Rpt(0) = "Employee specific"
''            Rpt(1) = "Leave List"
''
''    Case "Move_"
''            RItm = 2
''
''            Rpt(0) = "Employee specific"
''            Rpt(1) = "Department specific"
''            Rpt(2) = "All Employee"
''
''    Case "OT_"                      ''--------------"Overtime Report"
''            RItm = 2
''            Rpt(0) = "Employee specific"
''            Rpt(1) = "Section specific"
''            Rpt(2) = "All Employee"
''
''    Case "Pay_"
''            RItm = 6
''
''            Rpt(0) = "Pay slip (Dept)"
''            Rpt(1) = "Pay slip (All)"
''            Rpt(2) = "Salary Sheet (Dept)"
''            Rpt(3) = "Salary Sheet (All)"
''            Rpt(4) = "Salary Information (Dept)"
''            Rpt(5) = "Bank Statement (Dept)"
''            Rpt(6) = "Salary Statement(July-June)"
''
''
''
''   Case "Perform_"                 ''--------------"Performance Report"
''            RItm = 0
''            Rpt(0) = "Employee specific"
''
''
''   Case "Duty_"                     ''--------------"Duty Plan Report"
''
''            RItm = 2
''
''            Rpt(0) = "Employee specific"
''            Rpt(1) = "Shift specific"
''            Rpt(2) = "Group specific"
''
''
''End Select
''''-----------------------------------------------------
''cmbChooser_1.Clear
''For R = 0 To RItm
''    cmbChooser_1.AddItem Rpt(R)
''Next
''''-----------------------------------------------------
''
''Timer1.Enabled = True
''
''End Sub
''
''Private Sub Form_Unload(Cancel As Integer)
''Rpt_No = ""
''End Sub
''
''
''
''
''
''
''Private Sub Image3_Click()
''
''carry = "Report5"
''Call link_svr
''If permitted = 1 Then
''    permitted = 0
''    Unload Me
''
''    Rpt_No = "Move_"
''    Form27.lblRpt_Name = "Movement Report"
''    Form27.Show vbModal
''End If
''End Sub
''
''Private Sub Image4_Click()
''
''carry = "Report7"
''Call link_svr
''If permitted = 1 Then
''    permitted = 0
''
''    Unload Me
''    Rpt_No = "Pay_"
''
''    Form27.lblRpt_Name = "Payroll Report"
''    Form27.Show vbModal
''End If
''End Sub
''
''
''
''Private Sub Image6_Click()
''carry = "Report1"
''Call link_svr
''If permitted = 1 Then
''    permitted = 0
''    Unload Me
''
''    Rpt_No = "Per_Detail_"
''
''
''    Form27.Show vbModal
''End If
''End Sub
''
''Private Sub Timer1_Timer()
'''cmbChooser_1 = cmbChooser_1.List(0)
'''cmbChooser_1.SelStart = 0
'''cmbChooser_1.SelLength = Len(cmbChooser_1)
'''cmbChooser_1.SetFocus
''Timer1.Enabled = False
''
''
''End Sub
''
''
''Private Sub txtRpt_ID_LostFocus()
'''If txtRpt_ID = "" Then
'''    MsgBox "Must enter Employee ID", vbCritical + vbOKOnly, "Report Manager"
'''    txtRpt_ID.SetFocus
'''    Exit Sub
'''End If
''Rpt_ID = txtRpt_ID
''
''End Sub
''
''Public Sub Screen_Rearrange()
'''On Error Resume Next
''
''Select Case Rpt_No
''
''    Case "Per_Detail_"   ''-------------"Personal Information Report"
''
''            Yr_Mon_Position (3)
''
''    Case "Pay_"
''            Select Case cmbChooser_1
''                Case "Pay slip"
''                    Yr_Mon_Position (2)
''                Case "Salary Information (Dept)"
''                    Yr_Mon_Position (3)
''                Case "Bank Statement (Dept)"
''                    Yr_Mon_Position (2)
''                Case Else
''                    Yr_Mon_Position (2)
''            End Select
''
''    Case "OT_"
''
''            Select Case cmbChooser_1
''
''                Case "Employee specific"
''                    Yr_Mon_Position (2)
''                Case Else
''                    Yr_Mon_Position (2)
''            End Select
''
''    Case "Perform_"      ''--------------"Performance Report"
''        Yr_Mon_Position (2)
''
''    Case "Duty_"
''
''            Select Case cmbChooser_1
''
''                Case "Employee specific"
''                    Yr_Mon_Position (2)
''                Case Else
''                    Yr_Mon_Position (2)
''            End Select
''
''    Case Else                    ''-------------"Attendance/Leav/Movement/OT/Payroll
''            Select Case cmbChooser_1
''                Case "Employee specific"
''                    Yr_Mon_Position (2)
''                Case Else
''                    Yr_Mon_Position (2)
''            End Select
''
''End Select
''''-----------------------------------------
''If cmbChooser_1 = "Employee specific" Or cmbChooser_1 = "Pay slip" Then
''        With picButton
''            .Top = 3315
''            .Left = 1485
''        End With
''Else
''        With picButton
''            .Top = 3315
''            .Left = 1485
''        End With
''End If
''
''
''End Sub
''
''Public Sub Yr_Mon_Position(Num As Integer)
''
''Select Case Num
''
''    Case 1              ''Other than Emp. Sp.
''        With cmbRpt_Month
''            .Top = 2790
''            .Left = 1485
''            .Visible = True
''        End With
''
''        With cmbRpt_Year
''            .Top = 2790
''            .Left = 2790
''            .Visible = True
''        End With
''
''        With lblMon_Yr
''            .Top = 2850
''            .Visible = True
''        End With
''
''    Case 2              '' Emp. Sp.
''        With cmbRpt_Month
''            .Top = 2890
''            .Left = 1485
''            .Visible = True
''        End With
''
''        With cmbRpt_Year
''            .Top = 2890
''            .Left = 2790
''            .Visible = True
''        End With
''
''        With lblMon_Yr
''            .Top = 2950
''            .Visible = True
''        End With
''    Case 3
''            With cmbRpt_Month
''                .Visible = False
''            End With
''
''            With cmbRpt_Year
''                .Visible = False
''            End With
''            lblMon_Yr.Visible = False
''
''End Select
''
''End Sub
''
''Public Sub Change_Chooser()
''
''
''Select Case Trim(cmbChooser_1.Text)
''
''    Case "Employee specific", "Pay slip"
''
''        Rpt_sno = "ID"
''
''        With DataGrid0
''            .Height = 930
''            .Width = 3150
''            .Left = 450
''            .Top = 1795
''            .Visible = True
''        End With
''
''        txtRpt_ID.Visible = True
''        txtRpt_ID.SetFocus
''        lblParam = "Employee ID"
''        lstParam.Visible = False
''
''    Case "Designation specific"
''
''        Rpt_sno = "Desig"
''
''        lblParam = "Designation"
''        lstParam.Visible = True
''        lstParam.SetFocus
''        txtRpt_ID.Visible = False
''        DataGrid0.Visible = False
''
''
''    Case "Department specific", "Salary Sheet (Dept)"
''
''        Rpt_sno = "Dept"
''
''        lblParam = "Departmant"
''        lstParam.Visible = True
''        lstParam.SetFocus
''        txtRpt_ID.Visible = False
''        DataGrid0.Visible = False
''
''    Case "Pay slip (Dept)"
''
''        Rpt_sno = "PSlip_Dept"
''
''        lblParam = "Departmant"
''        lstParam.Visible = True
''        lstParam.SetFocus
''        txtRpt_ID.Visible = False
''        DataGrid0.Visible = False
''
''    Case "Section specific", "Salary Sheet (Sec)"
''
''        Rpt_sno = "Sec"
''
''        lblParam = "Section"
''        lstParam.Visible = True
''        lstParam.SetFocus
''        txtRpt_ID.Visible = False
''        DataGrid0.Visible = False
''
''
''    Case "Bloodgroup specific"
''
''        Rpt_sno = "Bld_Gr"
''
''        lblParam = "Bloodgroup"
''        lstParam.Visible = True
''        lstParam.SetFocus
''        txtRpt_ID.Visible = False
''        DataGrid0.Visible = False
''
''
''    Case "All Employee", "Salary Sheet (All)"
''
''        Rpt_sno = "All"
''
''        lblParam = ""
''        lstParam.Clear
''        lstParam.Visible = True
''        txtRpt_ID.Visible = False
''        DataGrid0.Visible = False
''
''    Case "Pay slip (All)"
''
''        Rpt_sno = "PSlip_All"
''
''        lblParam = ""
''        lstParam.Clear
''        lstParam.Visible = True
''        txtRpt_ID.Visible = False
''        DataGrid0.Visible = False
''
''    Case "Salary Statement(July-June)"
''
''         Rpt_sno = "Fiscal_Stat"
''
''         With DataGrid0
''            .Height = 930
''            .Width = 3150
''            .Left = 450
''            .Top = 1795
''            .Visible = True
''        End With
''
''        txtRpt_ID.Visible = True
''        txtRpt_ID.SetFocus
''        lblParam = "Employee ID"
''        lstParam.Visible = False
''
''
''    Case "Job Type"
''
''        Rpt_sno = "JbTp"
''
''        lblParam = "Job Type"
''        lstParam.Clear
''        lstParam.Visible = True
''        txtRpt_ID.Visible = False
''        DataGrid0.Visible = False
''
''
''    Case "Group specific"
''
''        Rpt_sno = "Group"
''
''        lblParam = "Duty Group"
''        lstParam.Visible = True
''        lstParam.SetFocus
''        txtRpt_ID.Visible = False
''        DataGrid0.Visible = False
''
''    Case "Shift specific"
''
''        Rpt_sno = "Shift"
''
''        lblParam = "Duty Shift"
''        lstParam.Visible = True
''        lstParam.SetFocus
''        txtRpt_ID.Visible = False
''        DataGrid0.Visible = False
''
''    Case "Salary Information (Dept)"
''
''        Rpt_sno = "Sal_Info"
''
''        lblParam = "Departmant"
''        lstParam.Visible = True
''        lstParam.SetFocus
''        txtRpt_ID.Visible = False
''        DataGrid0.Visible = False
''
''    Case "Bank Statement (Dept)"
''
''        Rpt_sno = "Bank_Stat"
''
''        lblParam = "Departmant"
''        lstParam.Visible = True
''        lstParam.SetFocus
''        txtRpt_ID.Visible = False
''        DataGrid0.Visible = False
''
''
''End Select
''
''    Screen_Rearrange
''
''    Append_Desig_Dept_ID
''
''
''End Sub
''
''
