VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form emp_info 
   BackColor       =   &H00C9AD8F&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Employee Information Form"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7635
   Icon            =   "emp_info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5970
      Top             =   5400
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C9AD8F&
      Height          =   1755
      Left            =   90
      TabIndex        =   32
      Top             =   3480
      Width           =   7365
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "emp_info.frx":08CA
         Height          =   1335
         Left            =   180
         TabIndex        =   33
         Top             =   210
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   2355
         _Version        =   393216
         BackColor       =   16777215
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
   Begin VB.CommandButton cmdExit 
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
      Left            =   2010
      Picture         =   "emp_info.frx":08DF
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Exit"
      Top             =   5355
      Width           =   510
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
      Left            =   -15
      Picture         =   "emp_info.frx":11FD
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Save"
      Top             =   5355
      Width           =   495
   End
   Begin VB.CommandButton cmdPreview 
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
      Left            =   1500
      Picture         =   "emp_info.frx":1867
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Preview"
      Top             =   5355
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
      Left            =   480
      Picture         =   "emp_info.frx":1ED1
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "New"
      Top             =   5355
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
      Left            =   990
      Picture         =   "emp_info.frx":253B
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Delete"
      Top             =   5355
      Width           =   510
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C9AD8F&
      Caption         =   "Employee Info Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   7380
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   5850
         TabIndex        =   11
         Top             =   1665
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   609
         _Version        =   393216
         Format          =   24772609
         CurrentDate     =   38008
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "emp_info.frx":3075
         Left            =   5850
         List            =   "emp_info.frx":307F
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1305
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "emp_info.frx":308F
         Left            =   5850
         List            =   "emp_info.frx":309F
         TabIndex        =   13
         Top             =   2430
         Width           =   1470
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "emp_info.frx":30C8
         Left            =   5850
         List            =   "emp_info.frx":30D2
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   2070
         Width           =   1515
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   8
         Left            =   1800
         TabIndex        =   9
         Top             =   2745
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   7
         Left            =   1800
         TabIndex        =   8
         Top             =   2385
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   6
         Left            =   1800
         TabIndex        =   7
         Top             =   2025
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   5
         Left            =   1800
         TabIndex        =   6
         Top             =   1665
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   4
         Left            =   1800
         TabIndex        =   5
         Top             =   1305
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   3
         Left            =   1830
         TabIndex        =   4
         Top             =   945
         Width           =   5550
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   2
         Left            =   1800
         TabIndex        =   3
         Top             =   585
         Width           =   5550
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   4770
         TabIndex        =   2
         Top             =   180
         Width           =   2580
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   1
         Top             =   225
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         Height          =   195
         Index           =   9
         Left            =   4905
         TabIndex        =   31
         Top             =   2115
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Join Date"
         Height          =   195
         Index           =   12
         Left            =   4905
         TabIndex        =   30
         Top             =   1710
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Index           =   11
         Left            =   4905
         TabIndex        =   29
         Top             =   1350
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         Height          =   195
         Index           =   10
         Left            =   4905
         TabIndex        =   28
         Top             =   2520
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary"
         Height          =   195
         Index           =   8
         Left            =   270
         TabIndex        =   27
         Top             =   2790
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         Height          =   195
         Index           =   7
         Left            =   270
         TabIndex        =   26
         Top             =   2430
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   195
         Index           =   6
         Left            =   270
         TabIndex        =   25
         Top             =   2025
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail Address"
         Height          =   195
         Index           =   5
         Left            =   270
         TabIndex        =   24
         Top             =   1665
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   23
         Top             =   1305
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permanent Address"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   22
         Top             =   945
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Present Address"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   21
         Top             =   630
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         Height          =   195
         Index           =   1
         Left            =   3420
         TabIndex        =   20
         Top             =   225
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Id"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   19
         Top             =   270
         Width           =   870
      End
   End
End
Attribute VB_Name = "emp_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Con As New MyConnection
Dim Conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset
'Public rptMode As Integer
Public strUid As String
Public strcn        As New MyConnection


Private Sub cmdADD_Click()


Call clear
    


End Sub


Private Sub clear()
Dim i
'On Error Resume Next
  

For i = 0 To 8


    Text1(i).Text = ""
    Text1(i).SetFocus

   
Next


 Text1(0).SetFocus

End Sub

Private Sub cmdDelete_Click()

Flush_Grid

Dim reply As String
    reply = MsgBox("Do you want to Delete?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
    

Dim i As Integer

Conn.ConnectionString = strcn.Connection_String
Conn.Open
cmd.ActiveConnection = Conn
cmd.CommandType = adCmdText
cmd.CommandText = "delete from  emp_info where emp_id='" & Trim(Text1(0).Text) & "'"
cmd.Execute
Conn.Close
End If
On Error Resume Next
 For i = 0 To Text1.Count
     Text1(i) = ""
 Next
 
 Call Flush_Grid
   
Call clear

End Sub

Private Sub cmdExit_Click()
Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdSave_Click()
Dim i
'On Error Resume Next
  

For i = 0 To 9

  Select Case i
  
  Case 0, 1, 2, 3
  
    If Text1(i) = Empty Then
        MsgBox Label1(i) + " Requied"
        Text1(i).SetFocus
        Exit Sub
   End If
   
 End Select
   
Next


'   If Len(Trim(txtfields(4).Text)) = 0 Then
'       MsgBox "Patient Name Required ", vbCritical
'       txtPatinetName.SetFocus
'       Exit Sub
'    End If
'
'    If Len(Trim(txtCompAddress.Text)) = 0 Then
'       MsgBox "Company address required", vbCritical
'       txtCompAddress.SetFocus
'       Exit Sub
'    End If
    
    Call SaveDoctorInfo
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
'    Call FlushCompSetup
    Call Flush_Grid
    Call clear
End Sub

Private Sub SaveDoctorInfo()

    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    Dim Param10 As New Parameter
    Dim Param11 As New Parameter
    Dim Param12 As New Parameter
    Dim Param13 As New Parameter
    Dim Param14 As New Parameter
    
'    Dim temp As Date
'    temp = DTPicker1.Value
'    Dim Param11 As New Parameter
'    Dim Param12 As New Parameter
'    Dim Param13 As New Parameter
'    Dim Param14 As New Parameter
'    Dim Param15 As New Parameter
'    Dim Param16 As New Parameter
    
    
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    If Text1(8).Text = "" Then
        Text1(8).Text = 0
     End If
     
      If Text1(6).Text = "" Then
        Text1(6).Text = 0
     End If
     
      If Text1(7).Text = "" Then
        Text1(7).Text = "No Title"
     End If
     
     
     
    
'
''
''    p_refer_code ,
''p_doc_name,
''p_addr  ,p_phone  ,p_fax,
''p_email , birth_date,
''p_marriage_status ,
''p_marriage_date,p_u_id,sysdate);
'commit;
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 20, Text1(0).Text)
    cmd.Parameters.Append Param1 'emp_id

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 45, Text1(1).Text)
    cmd.Parameters.Append Param2 'emp_name
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, DTPicker1.Value)
    cmd.Parameters.Append Param3 'join_date
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 12, Combo2.Text)
    
    cmd.Parameters.Append Param4 'Emp_designation
    
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 25, Text1(7).Text)
    cmd.Parameters.Append Param5 'Title
    
    Set Param6 = cmd.CreateParameter("param6", adInteger, adParamInput, 30, Text1(8).Text)
    
    cmd.Parameters.Append Param6 'salary
    
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 6, Trim(Combo1.Text))
    cmd.Parameters.Append Param7 'sex
    
    
    Set Param8 = cmd.CreateParameter("Param8", adVarChar, adParamInput, 3, Text1(6).Text)
    
    cmd.Parameters.Append Param8 'Age
    
    
    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 100, Text1(2).Text)
    cmd.Parameters.Append Param9 'Present address
    
    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 100, Text1(3).Text)
    cmd.Parameters.Append Param10 'permanent Address
    
  
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 15, Text1(4).Text)
    
    cmd.Parameters.Append Param11 'Phone
    
    
    Set Param12 = cmd.CreateParameter("param12", adVarChar, adParamInput, 25, Text1(5).Text)
    
    cmd.Parameters.Append Param12 'Email
    
    
    
    Set Param13 = cmd.CreateParameter("param13", adVarChar, adParamInput, 20, Combo3.ListIndex)
    
    cmd.Parameters.Append Param13 'Shift;
    
    
    Set Param14 = cmd.CreateParameter("param14", adVarChar, adParamInput, 10, "nh")
    cmd.Parameters.Append Param14 'u_id now defautl as "nh
    
 
'
'    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, dtpEnd.Value)
'    cmd.Parameters.Append Param4
'
'    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 10, userid)
'    cmd.Parameters.Append Param5
'
'    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL SaveEmp_info(?,?,?,?, ?,?,?,?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

End Sub

Private Sub Flush_Grid()
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select emp_id,emp_name,emp_desig as Desigantion ,emp_pre_add as Present_Address ,emp_per_add as Permanent_Address,emp_phone as Phone,age,salary,emp_email,join_date,sex,join_date from emp_info order by emp_id desc"
    Adodc1.Refresh
    
End Sub


Private Sub DataGrid1_Click()

'On Error Resume Next
'Call Flush_Grid
'If DataGrid1.Row > 0 Then
Text1(0).Text = DataGrid1.Columns(0).Text
Text1(1).Text = DataGrid1.Columns(1).Text
Combo2.Text = DataGrid1.Columns(2).Text
Text1(2).Text = DataGrid1.Columns(3).Text
Text1(3).Text = DataGrid1.Columns(4).Text
Text1(4).Text = DataGrid1.Columns(5).Text
Text1(6).Text = DataGrid1.Columns(6).Text
Text1(8).Text = DataGrid1.Columns(7).Text
Text1(5).Text = DataGrid1.Columns(8).Text
DTPicker1.Value = DataGrid1.Columns(11)
Combo1.Text = DataGrid1.Columns(10).Text
'End If

End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
End If
End Sub
Private Sub Form_Load()
Combo1 = Combo1.List(0)
Combo2 = Combo2.List(0)
Combo3 = Combo3.List(0)
Call Flush_Grid


End Sub



