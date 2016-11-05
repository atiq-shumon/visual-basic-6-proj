VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTowhomitConcern 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6765
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Show in Two Page"
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   3780
      Width           =   2235
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5445
      Picture         =   "frmTowhomitConcern.frx":0000
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&View Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4215
      Picture         =   "frmTowhomitConcern.frx":08CA
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Income Statment Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   3525
      Left            =   30
      TabIndex        =   0
      Top             =   120
      Width           =   6630
      Begin VB.TextBox ScaleText 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1620
         TabIndex        =   18
         Top             =   1680
         Width           =   4965
      End
      Begin VB.Frame Frame2 
         Height          =   435
         Left            =   390
         TabIndex        =   15
         Top             =   210
         Width           =   3285
         Begin VB.OptionButton optnFormat 
            Appearance      =   0  'Flat
            Caption         =   "New Format"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   17
            Top             =   150
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optnFormat 
            Appearance      =   0  'Flat
            Caption         =   "Old Format"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1650
            TabIndex        =   16
            Top             =   150
            Width           =   1545
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Works as Head"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00915411&
         Height          =   240
         Index           =   1
         Left            =   3750
         TabIndex        =   14
         Top             =   2430
         Width           =   2445
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Work's under Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00915411&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   13
         Top             =   2430
         Width           =   2955
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H008080FF&
         Height          =   330
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1170
         Width           =   4875
      End
      Begin VB.TextBox txtGender 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3690
         TabIndex        =   8
         Top             =   750
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Top             =   750
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker BeginDate 
         Height          =   315
         Left            =   1620
         TabIndex        =   6
         Top             =   2970
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   8388608
         CalendarTitleForeColor=   8388608
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   61603843
         CurrentDate     =   36998
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   4440
         TabIndex        =   7
         Top             =   2955
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   8388608
         CalendarTitleForeColor=   8388608
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   61603843
         CurrentDate     =   36998
      End
      Begin VB.Label Label5 
         Caption         =   "Scale:"
         Height          =   285
         Left            =   330
         TabIndex        =   19
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Shape Shape6 
         Height          =   435
         Left            =   330
         Top             =   2340
         Width           =   6135
      End
      Begin VB.Label Label4 
         Caption         =   "To "
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3990
         TabIndex        =   12
         Top             =   3030
         Width           =   525
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   360
         TabIndex        =   11
         Top             =   1170
         Width           =   1185
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   4415
         Top             =   1155
         Width           =   2075
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   1590
         Top             =   1155
         Width           =   2100
      End
      Begin VB.Label Label3 
         Caption         =   "From "
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   390
         TabIndex        =   5
         Top             =   3000
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Emp ID"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   360
         TabIndex        =   2
         Top             =   795
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   1290
         Top             =   1155
         Width           =   5355
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   1590
         Top             =   2.00150e5
         Width           =   1680
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   1590
         Top             =   705
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmTowhomitConcern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BeginDate_Change()
On Error GoTo Errdes
Dim Myconn1 As New Connection
Dim cmd1 As New Command
Dim myrs1 As New ADODB.Recordset
Myconn1.ConnectionString = strCN.Connection_String
Myconn1.Open
cmd1.ActiveConnection = Myconn1
cmd1.CommandType = adCmdText
cmd1.CommandText = "select add_months(to_date('" & BeginDate & "','dd-mm-yy'), 12) from dual"
cmd1.Properties("iRowsetChange") = True
cmd1.Properties("updatability") = 7
myrs1.CursorLocation = adUseClient
myrs1.Open cmd1.CommandText, Myconn1, adOpenDynamic, adLockOptimistic

If Not myrs1.EOF Then
     DTPicker1 = Format(myrs1.Fields(0), "dd-mmm-yyyy")
    
End If
myrs1.Close
Myconn1.Close

Exit Sub
Errdes:
    If Err.Number = 13 Then
        MsgBox "Date is not in Proper Format"
    Else
        MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
    End If
End Sub


Private Sub BeginDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    DTPicker1.SetFocus
End If
End Sub

Private Sub Check1_Click()
  If Check1.Value = 1 Then
     twoPage = 1
  Else
     twoPage = 0
  End If
  
End Sub

Private Sub Combo1_Click()
  Call Combo1_KeyDown(13, 10)
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Get_Emp_Name
    BeginDate.SetFocus
End If
End Sub

Private Sub Combo1_LostFocus()
  Call Combo1_KeyDown(13, 10)
End Sub

Private Sub Command1_Click()
   
   If Len(Combo1.Text) = 0 Then
      MsgBox "Pls Put Employee ID..", vbInformation, "Software Programmer,IT,DNMIH"
      Combo1.SetFocus
      Exit Sub
   End If
    rptmode = 19
    GetEmployee_info
    EmployeeName = Text1.Text
    EmpIDForTowhom = Trim(Combo1)
    BEGINYEARFORWHOM = BeginDate
    ENDDATEFORWHOM = DTPicker1
    GetFromMonthtoWhom = Format(BeginDate, "MMMM") & " '" & Format(BEGINYEARFORWHOM, "YYYY")
    GetToMonthtoWhom = Format(DTPicker1, "MMMM") & " '" & Format(ENDDATEFORWHOM, "YYYY")
    If Option1(0).Value = True Then
       underDepartmentorNot = 0
    Else
       underDepartmentorNot = 1
    End If
    If optnFormat(0).Value = True Then
       currentFormat = 0
    Else
       currentFormat = 1
    End If
    If Len(ScaleText) > 0 Then
       pScale = ScaleText.Text
    Else
      pScale = ""
    End If
    Form20.Show vbModal
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub Form_Load()
On Error GoTo Errdes
twoPage = 0
Dim cmd As New Command
Dim conn10 As New Connection
Dim rs10 As New Recordset
Dim conn11 As New Connection
Dim rs11 As New Recordset

conn10.ConnectionString = strCN.Connection_String
conn10.Open
cmd.ActiveConnection = conn10
cmd.CommandType = adCmdText

cmd.CommandText = "select emp_id from emp_info order by emp_id"


rs10.CursorLocation = adUseClient
rs10.Open cmd.CommandText, conn10, adOpenDynamic, adLockOptimistic

If rs10.RecordCount > 0 Then
    Do Until rs10.EOF
        Combo1.AddItem rs10.Fields(0)
        rs10.MoveNext
    Loop
End If

rs10.Close
conn10.Close


Dim BeginFiscalYear, BeginDateYear


BeginFiscalYear = "30 Jul 2006"

YearPartOftheDate = Format(Date, "YYYY") - 1
BeginDateYear = Format(BeginFiscalYear, "DD MMM ")
BeginDate = Format(BeginDateYear & YearPartOftheDate, "dd/mm/yyyy")
BeginDate_Change

Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub
Private Sub Get_Emp_Name()
'On Error GoTo Errdesc
Dim cmd As New Command
Dim conn2 As New ADODB.Connection
Dim RS2 As New ADODB.Recordset
conn2.ConnectionString = strCN.Connection_String
conn2.Open
cmd.ActiveConnection = conn2
cmd.CommandType = adCmdText

cmd.CommandText = "select EMP_NM,gender from emp_info where Emp_Id='" & Combo1 & "'"

RS2.CursorLocation = adUseClient
RS2.Open cmd.CommandText, conn2, adOpenDynamic, adLockOptimistic

    If RS2.RecordCount > 0 Then
        Text1.Text = RS2.Fields(0)
        EmployeeName = RS2.Fields(0)
        sGender = IIf(RS2.Fields(1) = 0, "F", "M")
        txtGender = sGender
        RS2.Close
        conn2.Close
    Else
        Text1.Text = ""
        RS2.Close
        conn2.Close
    End If
Exit Sub
Errdesc:
MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub GetEmployee_info()
On Error GoTo Errdes
Dim DateofRetirementYearPArt As String
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New ADODB.Recordset
conn.ConnectionString = strCN.Connection_String
conn.Open
cmd.ActiveConnection = conn
cmd.CommandType = adCmdText

cmd.CommandText = " SELECT EMP_INFO.EMP_NM,ST_DEPT.DEPT_NM,st_desig.DESIGNATION,EMP_JOB_INFO.JDate,EMP_INFO.DOB,(select jtype_nm from st_jbtype where JTYPE_CODE=EMP_JOB_INFO.jtype) Job_Type  From Emp_info, EMP_JOB_INFO," + _
                    " ST_DEPT,st_desig WHERE ((EMP_JOB_INFO.EMP_ID=EMP_INFO.EMP_ID) " + _
                    " AND (EMP_JOB_INFO.DEPT=ST_DEPT.DEPT_CODE) " + _
                    " and (emp_job_info.DESIG=st_desig.DESIG_CODE) AND EMP_INFO.EMP_ID='" & Trim(Combo1) & "')"

cmd.Properties("iRowsetChange") = True
cmd.Properties("updatability") = 7
RS.CursorLocation = adUseClient

RS.Open cmd.CommandText, conn, adOpenDynamic, adLockOptimistic

If Not RS.EOF Then
    EmployeeName = RS.Fields(0)
    DesignationOfEmp = RS.Fields(2)
    DepartemntOfEmp = RS.Fields(1)
    DatofJoin = RS.Fields(3)
    DateofRetirementYearPArt = Format(RS.Fields(4), "YYYY") + 57
    DateofRetirement = Format(RS.Fields(4), "DD MMM ") + DateofRetirementYearPArt
    JobType = LCase(RS.Fields(5))
    
Else
    MsgBox "Invalid Emp ID !", vbCritical, "IT,DNMIH"
    Combo1.SetFocus
    Exit Sub
End If
   
RS.Close
conn.Close
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub


