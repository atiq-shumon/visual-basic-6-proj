VERSION 5.00
Begin VB.Form CurrentPatientUI 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Current Patient Statements"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8325
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox DaysAboveCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4350
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   4560
      Width           =   3045
   End
   Begin VB.ComboBox CabinOrWardCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4350
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   3690
      Width           =   2985
   End
   Begin VB.ComboBox BedTypeCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4350
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   2790
      Width           =   2985
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   -180
      TabIndex        =   17
      Top             =   6360
      Width           =   8625
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Software Programmer,IT Division,DNMIH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2760
         TabIndex        =   19
         Top             =   180
         Width           =   4890
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Developed && Maintenanced By: "
         ForeColor       =   &H008080FF&
         Height          =   195
         Left            =   390
         TabIndex        =   18
         Top             =   210
         Width           =   2295
      End
   End
   Begin VB.ComboBox FiscalYearsCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   540
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2820
      Width           =   2505
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Options"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   -60
      TabIndex        =   10
      Top             =   930
      Width           =   8415
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Staying Days Wise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   5940
         TabIndex        =   25
         Top             =   840
         Width           =   2205
      End
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Staff Wise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   4710
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bed Type && Department Wise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   630
         TabIndex        =   13
         Top             =   840
         Width           =   2835
      End
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bed Type && Cabin/Ward Wise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   4710
         TabIndex        =   12
         Top             =   330
         Width           =   3255
      End
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bed Type Wise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   630
         TabIndex        =   11
         Top             =   330
         Value           =   -1  'True
         Width           =   1665
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   -60
      TabIndex        =   7
      Top             =   5490
      Width           =   8715
      Begin VB.CommandButton ShowButton 
         Caption         =   "Show"
         Height          =   405
         Left            =   4530
         TabIndex        =   9
         Top             =   300
         Width           =   1605
      End
      Begin VB.CommandButton CloseButton 
         Caption         =   "Close"
         Height          =   405
         Left            =   6210
         TabIndex        =   8
         Top             =   300
         Width           =   1605
      End
      Begin VB.Shape Shape1 
         Height          =   525
         Left            =   4470
         Top             =   240
         Width           =   3405
      End
      Begin VB.Image Image2 
         Height          =   855
         Left            =   -2970
         Picture         =   "CurrentPatientUI.frx":0000
         Top             =   0
         Width           =   11820
      End
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   -120
      TabIndex        =   6
      Top             =   -90
      Width           =   8715
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Patient Statement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   2040
         TabIndex        =   16
         Top             =   270
         Width           =   4485
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -300
         Picture         =   "CurrentPatientUI.frx":86C4
         Top             =   60
         Width           =   11820
      End
   End
   Begin VB.ComboBox PatientStatusCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   540
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4560
      Width           =   2505
   End
   Begin VB.ComboBox DepartmentCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   540
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3690
      Width           =   2505
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   540
      TabIndex        =   21
      Top             =   4200
      Width           =   1515
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fiscal Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   540
      TabIndex        =   20
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Staying Days Above"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4350
      TabIndex        =   5
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   540
      TabIndex        =   4
      Top             =   3360
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cabin / Ward "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4350
      TabIndex        =   3
      Top             =   3360
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bed Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4350
      TabIndex        =   2
      Top             =   2400
      Width           =   1140
   End
End
Attribute VB_Name = "CurrentPatientUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UTILITY As New clsUtility
Private Sub BedTypeCombo_Click()
   GetCabinWardLists (BedTypeCombo.Text)
End Sub

Private Sub CloseButton_Click()
  Unload Me
End Sub
Private Sub Form_Load()
     PopulateDays
     PopulateBed
     BedTypeCombo.ListIndex = 0
     DaysAboveCombo.ListIndex = 0
     GetCabinWardLists (BedTypeCombo.Text)
     GetDepartment
     PopulateFiscalYears
     ContentEnableDisable (0)
     PopulatePatientStatus
  End Sub
Private Sub PopulateDays()
    Dim i As Integer
    For i = 1 To 20
        DaysAboveCombo.AddItem i
    Next i
End Sub
Private Sub PopulatePatientStatus()
    Dim patientStatusList() As String
    Dim i As Integer
    patientStatusList = UTILITY.GetPatientStatus()
    For i = LBound(patientStatusList) To UBound(patientStatusList)
         PatientStatusCombo.AddItem patientStatusList(i)
    Next i
    PatientStatusCombo.ListIndex = 0
End Sub
Private Sub PopulateFiscalYears()
 Dim yearList() As String

 Dim i As Integer
    yearList = UTILITY.GetFiscalYears()
    For i = LBound(yearList) To UBound(yearList)
       FiscalYearsCombo.AddItem yearList(i)
    Next i
    FiscalYearsCombo.ListIndex = 0
End Sub
Private Sub PopulateBed()
Dim bedTypeList() As String
Dim i As Integer
bedTypeList = UTILITY.GetBedType()
For i = LBound(bedTypeList) To UBound(bedTypeList)
    BedTypeCombo.AddItem bedTypeList(i)
Next i

BedTypeCombo.ListIndex = 0
   
End Sub

Private Sub Option_Click(index As Integer)
   ContentEnableDisable (index)
   Select Case index
          Case 0
               paramMode = 0
          Case 1
               paramMode = 1
          Case 2
               paramMode = 2
          Case 3
               paramMode = 3
          Case 4
               paramMode = 4
   End Select
End Sub
Private Sub ContentEnableDisable(index As String)
   Select Case index
          Case 0, 3
             Label2.Enabled = False
             Label3.Enabled = False
             Label4.Enabled = False
             CabinOrWardCombo.Enabled = False
             DepartmentCombo.Enabled = False
             DaysAboveCombo.Enabled = False
          Case 1
             Label2.Enabled = True
             Label4.Enabled = False
             CabinOrWardCombo.Enabled = True
             DepartmentCombo.Enabled = False
             DaysAboveCombo.Enabled = False
          Case 2
            Label2.Enabled = False
            Label3.Enabled = True
            Label4.Enabled = False
            CabinOrWardCombo.Enabled = False
            DepartmentCombo.Enabled = True
            DaysAboveCombo.Enabled = False
          Case 4
            Label2.Enabled = False
            Label3.Enabled = False
            Label4.Enabled = True
            CabinOrWardCombo.Enabled = False
            DepartmentCombo.Enabled = False
            DaysAboveCombo.Enabled = True
   End Select
   
End Sub
Private Sub ShowButton_Click()
   PatientStatus = UTILITY.GetPatientStatusNo(PatientStatusCombo.Text)
   rptMode = 508
   Viewer.Show 1
End Sub
Private Sub GetCabinWardLists(bedType As String)
     Dim Conn As New ADODB.Connection
     Dim cmd As New ADODB.Command
     Dim RS As New ADODB.Recordset
     CabinOrWardCombo.clear
     If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
      cmd.ActiveConnection = Conn
     cmd.CommandType = adCmdText
     cmd.CommandText = "select distinct bed_ext_col from bed_info WHERE BED_TYPE='" & bedType & "'"
   
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
   
     If RS.RecordCount > 0 Then
         RS.MoveFirst
         Do Until RS.EOF = True
           CabinOrWardCombo.AddItem RS!bed_ext_col
           RS.MoveNext
        Loop
       CabinOrWardCombo.Text = CabinOrWardCombo.List(0)
    End If
    cmd.Properties("PLSQLRSet") = False
     
End Sub
Private Sub GetDepartment()
   Dim Conn As New ADODB.Connection
   Dim cmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   DepartmentCombo.clear
   
   If Conn.State = 0 Then
      Conn.Open strcn.Connection_String
   End If
   cmd.ActiveConnection = Conn
   cmd.CommandType = adCmdText
   cmd.CommandText = "select distinct doc_department from bed_info order by doc_department desc"
   cmd.Properties("iRowsetChange") = True
   cmd.Properties("updatability") = 7
   RS.CursorLocation = adUseClient
   RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
   
   If RS.RecordCount > 0 Then
      RS.MoveFirst
      
      Do Until RS.EOF = True
         DepartmentCombo.AddItem RS!doc_department
         RS.MoveNext
      Loop
      DepartmentCombo.Text = DepartmentCombo.List(0)
   End If
   cmd.Properties("PLSQLRSet") = False
End Sub
