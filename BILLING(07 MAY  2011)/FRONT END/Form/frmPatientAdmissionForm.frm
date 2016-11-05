VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPatientAdmissionForm 
   BackColor       =   &H00C9AD8F&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Patient  Admission  Form"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2610
      Top             =   6030
      Visible         =   0   'False
      Width           =   1545
      _ExtentX        =   2725
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00C9AD8F&
      Height          =   555
      Left            =   135
      TabIndex        =   42
      Top             =   855
      Width           =   7980
      Begin VB.ComboBox Combo 
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   4
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   165
         Width           =   6165
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Dr. Code and Name"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   44
         Top             =   180
         Width           =   1815
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   255
      Left            =   6390
      TabIndex        =   35
      Top             =   4995
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   24444929
      CurrentDate     =   36161
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   2610
      TabIndex        =   34
      Top             =   4995
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   24444929
      CurrentDate     =   36161
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
      Left            =   1050
      Picture         =   "frmPatientAdmissionForm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Delete"
      Top             =   5940
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
      Left            =   540
      Picture         =   "frmPatientAdmissionForm.frx":0B3A
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "New"
      Top             =   5940
      Width           =   510
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
      Left            =   1560
      Picture         =   "frmPatientAdmissionForm.frx":11A4
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Preview"
      Top             =   5940
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
      Left            =   45
      Picture         =   "frmPatientAdmissionForm.frx":180E
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Save"
      Top             =   5940
      Width           =   495
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
      Left            =   2070
      Picture         =   "frmPatientAdmissionForm.frx":1E78
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Exit"
      Top             =   5940
      Width           =   510
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C9AD8F&
      Height          =   1155
      Left            =   90
      TabIndex        =   6
      Top             =   4770
      Width           =   8040
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   12
         Left            =   1575
         TabIndex        =   26
         Top             =   855
         Width           =   4935
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   11
         Left            =   1575
         TabIndex        =   25
         Top             =   540
         Width           =   4935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnosis date"
         Height          =   195
         Index           =   16
         Left            =   5085
         TabIndex        =   24
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Clinical Diagnosis"
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Relativer referance"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1350
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Admission Date"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C9AD8F&
      Height          =   3285
      Left            =   90
      TabIndex        =   5
      Top             =   1440
      Width           =   7965
      Begin VB.ComboBox Combo 
         Height          =   315
         Index           =   1
         Left            =   4995
         TabIndex        =   45
         Text            =   "Combo1"
         Top             =   135
         Width           =   1275
      End
      Begin VB.TextBox txtfields 
         Height          =   315
         Index           =   9
         Left            =   5895
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   1845
         Width           =   1305
      End
      Begin VB.TextBox txtfields 
         Height          =   330
         Index           =   8
         Left            =   1350
         TabIndex        =   38
         Top             =   1440
         Width           =   5805
      End
      Begin VB.ComboBox Combo 
         Height          =   315
         Index           =   0
         ItemData        =   "frmPatientAdmissionForm.frx":2796
         Left            =   1350
         List            =   "frmPatientAdmissionForm.frx":27A3
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   135
         Width           =   1365
      End
      Begin VB.ComboBox Combo 
         Height          =   315
         Index           =   3
         ItemData        =   "frmPatientAdmissionForm.frx":27C4
         Left            =   6030
         List            =   "frmPatientAdmissionForm.frx":27D7
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   2340
         Width           =   1215
      End
      Begin VB.ComboBox Combo 
         Height          =   315
         Index           =   2
         ItemData        =   "frmPatientAdmissionForm.frx":2806
         Left            =   1350
         List            =   "frmPatientAdmissionForm.frx":2810
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   4
         Left            =   1350
         TabIndex        =   18
         Top             =   495
         Width           =   4485
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   10
         Left            =   1260
         TabIndex        =   17
         Top             =   2340
         Width           =   2775
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   7
         Left            =   3600
         TabIndex        =   16
         Top             =   1845
         Width           =   1215
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   6
         Left            =   1350
         TabIndex        =   15
         Top             =   1125
         Width           =   5820
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   5
         Left            =   1350
         TabIndex        =   14
         Top             =   810
         Width           =   4485
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   345
         Left            =   5355
         TabIndex        =   41
         Top             =   1845
         Width           =   465
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Emergency Addr"
         Height          =   255
         Index           =   13
         Left            =   45
         TabIndex        =   39
         Top             =   1395
         Width           =   1275
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
         Height          =   195
         Index           =   9
         Left            =   5355
         TabIndex        =   20
         Top             =   2340
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         Height          =   195
         Index           =   17
         Left            =   180
         TabIndex        =   19
         Top             =   1845
         Width           =   270
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   13
         Top             =   2340
         Width           =   825
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   195
         Index           =   8
         Left            =   3060
         TabIndex        =   12
         Top             =   1845
         Width           =   285
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   11
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Guardgian Name:"
         Height          =   195
         Index           =   6
         Left            =   45
         TabIndex        =   10
         Top             =   855
         Width           =   1245
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   9
         Top             =   585
         Width           =   960
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Type"
         Height          =   195
         Index           =   3
         Left            =   540
         TabIndex        =   8
         Top             =   180
         Width           =   690
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed  No:"
         Height          =   195
         Index           =   2
         Left            =   4320
         TabIndex        =   7
         Top             =   180
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C9AD8F&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   7995
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   1
         Left            =   6525
         TabIndex        =   4
         Top             =   225
         Width           =   1215
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   0
         Left            =   1305
         TabIndex        =   3
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         Height          =   195
         Index           =   10
         Left            =   3360
         TabIndex        =   36
         Top             =   240
         Width           =   345
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Id:"
         Height          =   195
         Index           =   1
         Left            =   5490
         TabIndex        =   2
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Booth No:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmPatientAdmissionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Con As New MyConnection
Dim Conn As New Connection
Dim Conn1 As New Connection
Dim cmd As New Command
Dim RS As New Recordset
'Public rptMode As Integer
Public strUid As String
Public strcn        As New MyConnection


Private Sub cmdADD_Click()

Dim i
'On Error Resume Next
  

For i = 0 To 12

'  Select Case i
  
'  Case 0, 1, 2
  
'    If Text1(i) = Empty Then
          txtfields(i).Text = ""
          txtfields(i).SetFocus
'        Exit Sub
'   End If
'
' End Select
   
Next


 txtfields(0).SetFocus

    
End Sub


Private Sub cmdDelete_Click()
Dim reply As String
    reply = MsgBox("Do you want to Delete?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
    End If
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
''On Error Resume Next
  

For i = 0 To 12

  Select Case i
  
  Case 0, 1, 2, 4, 5, 6, 12
  
    If txtfields(i) = Empty Then
        MsgBox lbl(i) + " Requied"
       txtfields(i).SetFocus
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
    
    Call SavePatientInfo
'    Call update_flag
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
'    Call FlushCompSetup
    
End Sub

Private Sub SavePatientInfo()

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
    Dim Param15 As New Parameter
    Dim Param16 As New Parameter
    
    
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, txtfields(2).Text)
    cmd.Parameters.Append Param1 'bed_no

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 40, txtfields(4).Text)
    cmd.Parameters.Append Param2 'patient_name

    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 50, txtfields(5).Text)
    cmd.Parameters.Append Param3 'guardgian_name
    
    Set Param4 = cmd.CreateParameter("param4", adInteger, adParamInput, 10, Combo(0).Index)
    
    cmd.Parameters.Append Param4 'bed type
    
    Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 6, Combo(2).Index)
    
    cmd.Parameters.Append Param5 'Sex
    
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 3, txtfields(7).Text)
    cmd.Parameters.Append Param6 'age
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 30, txtfields(10).Text)
    cmd.Parameters.Append Param7 'occupation
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 200, txtfields(6).Text)
    cmd.Parameters.Append Param8 'pat_addr
    
    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 70, txtfields(8).Text)
    cmd.Parameters.Append Param9 'imergency addr
    
    Set Param10 = cmd.CreateParameter("param10", adInteger, adParamInput, 10, Get_Segment(Combo(4), False))
    cmd.Parameters.Append Param10 'refer code
    
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 15, txtfields(9).Text)
    cmd.Parameters.Append Param11 'phone
    
    Set Param12 = cmd.CreateParameter("param12", adVarChar, adParamInput, 10, "sumon")
    cmd.Parameters.Append Param12 'u_id
    
   
    
    Set Param13 = cmd.CreateParameter("param13", adVarChar, adParamInput, 5, txtfields(0).Text)
    cmd.Parameters.Append Param13 'booth
    
    Set Param14 = cmd.CreateParameter("param14", adDate, adParamInput, 10, DTPicker2.Value)
    cmd.Parameters.Append Param14
    
    Set Param15 = cmd.CreateParameter("param15", adVarChar, adParamInput, 5, txtfields(2).Text)
    cmd.Parameters.Append Param15 'booth
    
    

    
'
'    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, dtpEnd.Value)
'    cmd.Parameters.Append Param4
'
'    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 10, userid)
'    cmd.Parameters.Append Param5
'
'    '----------------------------------------------------------------------------------
'Get_Segment(0,1)
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL SavePatient_info( ?,?,?,?, ?,?,?,?, ?,?,?,?, ?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

End Sub

Private Sub Combo_Click(Index As Integer)
If Combo(0).ListIndex = 0 Then
'    Call Load_bed_no
End If
End Sub

'Private Sub Flush_Grid()
'    Adodc1.ConnectionString = strcn.Connection_String
'    Adodc1.RecordSource = "select refer_code,doc_name from doctor_info"
'    Adodc1.Refresh
'
'End Sub


'Private Sub Combo_Click(Index As Integer)
'
''On Error Resume Next
'
'
'Label3 = Get_Segment(Combo(4))
'
'End Sub

'Private Sub Command1_Click()
'
'Label3 = Get_Segment("1000199-Rafiqul Islam")
'
'End Sub



Private Sub Form_Load()
'On Error Resume Next
    
    Combo(0) = Combo(0).List(0)
    Combo(2) = Combo(2).List(0)
    Combo(3) = Combo(3).List(0)
    
    Call Load_Doctor
    Call Load_bed_no
   ' Call Flush_Grid
    
    
    '    'Combo(3).ListIndex = 2  ' to retrieve data from data base say 2 indicate bhuddist



End Sub


Private Sub Load_Doctor()

    Conn.ConnectionString = strcn.Connection_String
    Conn.Open
    cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "select refer_code,doc_name from doctor_info"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    RS.CursorLocation = adUseClient
    
    RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
    
    
    If RS.RecordCount > 0 Then
            
         RS.MoveFirst
         
        Do Until RS.EOF = True
        
            Combo(4).AddItem RS.Fields(0) + "-" + RS.Fields(1)
    
            RS.MoveNext
        Loop
    
    End If
    
    RS.Close
    
End Sub

Private Sub Load_bed_no()
'On Error Resume Next

    Conn1.ConnectionString = strcn.Connection_String
    Conn1.Open
    cmd.ActiveConnection = Conn1
    cmd.CommandType = adCmdText
   cmd.CommandText = "select bed_no from bed_info  where occupy_flag='0'and  bed_type= 'Cabin'"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    RS.CursorLocation = adUseClient
    
    RS.Open cmd.CommandText, Conn1, adOpenDynamic, adLockOptimistic
    
    
    If RS.RecordCount > 0 Then
            
         
         RS.MoveFirst
         
        Do Until RS.EOF = True
        
            Combo(1).AddItem RS.Fields(0)
    
            RS.MoveNext
        Loop
    
    End If
    
    RS.Close
    
End Sub


