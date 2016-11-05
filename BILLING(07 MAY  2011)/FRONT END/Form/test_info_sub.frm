VERSION 5.00
Begin VB.Form test_info_sub 
   BackColor       =   &H00C9AD8F&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Test details"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C9AD8F&
      Height          =   1050
      Left            =   135
      TabIndex        =   18
      Top             =   135
      Width           =   5730
      Begin VB.TextBox Text1 
         Height          =   330
         Index           =   1
         Left            =   1395
         TabIndex        =   20
         Top             =   540
         Width           =   4005
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Index           =   0
         Left            =   1395
         TabIndex        =   19
         Top             =   180
         Width           =   2025
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         Caption         =   "Sub Code"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   270
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         Caption         =   "Sub Name"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   675
         Width           =   750
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
      Left            =   2250
      Picture         =   "test_info_sub.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Exit"
      Top             =   3870
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
      Left            =   270
      Picture         =   "test_info_sub.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Save"
      Top             =   3870
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
      Left            =   1755
      Picture         =   "test_info_sub.frx":0F88
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Preview"
      Top             =   3870
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
      Left            =   735
      Picture         =   "test_info_sub.frx":15F2
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "New"
      Top             =   3870
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
      Left            =   1245
      Picture         =   "test_info_sub.frx":1C5C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Delete"
      Top             =   3870
      Width           =   510
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C9AD8F&
      Height          =   2535
      Left            =   90
      TabIndex        =   0
      Top             =   1260
      Width           =   5775
      Begin VB.TextBox Text7 
         Height          =   330
         Left            =   1350
         TabIndex        =   12
         Top             =   2025
         Width           =   2040
      End
      Begin VB.TextBox Text6 
         Height          =   330
         Left            =   1350
         TabIndex        =   11
         Top             =   1665
         Width           =   2040
      End
      Begin VB.TextBox Text4 
         Height          =   330
         Left            =   1350
         TabIndex        =   10
         Top             =   1305
         Width           =   2040
      End
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   1350
         TabIndex        =   6
         Top             =   540
         Width           =   2040
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   1350
         TabIndex        =   5
         Top             =   180
         Width           =   1995
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   1350
         TabIndex        =   1
         Top             =   900
         Width           =   2025
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         Caption         =   "Unique Id"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   2115
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         Caption         =   "DT"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   1710
         Width           =   225
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         Caption         =   "User Id"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   1395
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         Caption         =   "Type"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   1035
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         Caption         =   "Member Code"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   630
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         Caption         =   "Sub Code"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   270
         Width           =   705
      End
   End
End
Attribute VB_Name = "test_info_sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub


Private Sub cmdSAVE_Click()
Dim i


For i = 0 To 10

  Select Case i
  
  Case 0, 1, 2, 3, 4, 5
  
    If txtField(i) = Empty Then
        MsgBox Label1(i) + " Requied"
        txtField(i).SetFocus
        Exit Sub
   End If
   
 End Select
   
Next


    Call SaveDoctorInfo
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."

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
    

    
    
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    

    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, txtField(0).Text)
    cmd.Parameters.Append Param1 'refer_code

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 45, txtField(1).Text)
    cmd.Parameters.Append Param2 'doc_name
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 200, txtField(2).Text)
    cmd.Parameters.Append Param3 'addr
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 25, txtField(3).Text)
    
    cmd.Parameters.Append Param4 'phone
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 25, txtField(4).Text)
    
    cmd.Parameters.Append Param5 'Fax
    
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 25, txtField(5).Text)
    cmd.Parameters.Append Param6 'E-mail
    
    Set Param7 = cmd.CreateParameter("param7", adDate, adParamInput, 10, DTPicker1.Value)
    cmd.Parameters.Append Param7 'birth_date
    
  
    
    Set Param8 = cmd.CreateParameter("param8", adInteger, adParamInput, 12, Combo1.ListIndex)
    
    cmd.Parameters.Append Param8 'marriage_status
    
    Set Param9 = cmd.CreateParameter("param9", adDate, adParamInput, 10, DTPicker2.Value)
    cmd.Parameters.Append Param9 'marriage_date
    
    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 10, "sumon")
    cmd.Parameters.Append Param10 'u_id
 
'----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL SaveDoctor_info( ?,?,?,?, ?,?,?,?, ?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

End Sub

Private Sub Form_Load()
Combo1 = Combo1.List(0)

End Sub


Private Sub txtField_Change(Index As Integer)

End Sub

End Sub

