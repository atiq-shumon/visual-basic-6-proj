VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSTATISTICS 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   4170
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   2940
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000001&
      ForeColor       =   &H80000001&
      Height          =   855
      Left            =   -30
      TabIndex        =   14
      Top             =   -90
      Width           =   5535
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PATIENT STATISTICS"
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
         Left            =   390
         TabIndex        =   15
         Top             =   300
         Width           =   4755
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -1980
         Picture         =   "frmPatient_STATISTICS.frx":0000
         Top             =   0
         Width           =   11820
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1800
      Top             =   2490
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
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   1800
      TabIndex        =   5
      Top             =   780
      Width           =   3705
      Begin VB.ComboBox dept 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   630
         Width           =   3165
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         Caption         =   "Department Wise"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   1545
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "All"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   1020
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   38195
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   38195
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1500
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   38195
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1020
         Width           =   3135
      End
      Begin VB.Shape Shape2 
         Height          =   285
         Left            =   90
         Top             =   180
         Width           =   3225
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   60
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   6
         Top             =   6330
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Option"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   780
      Width           =   1785
      Begin VB.CheckBox Chk_date 
         Appearance      =   0  'Flat
         Caption         =   "Date Specific"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   390
         Width           =   1335
      End
      Begin VB.CheckBox Chk_date_to_date 
         Appearance      =   0  'Flat
         Caption         =   "Date to date"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   930
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Name Specific"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1470
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmPatient_STATISTICS.frx":5982
         Left            =   2280
         List            =   "frmPatient_STATISTICS.frx":598C
         TabIndex        =   1
         Top             =   2160
         Width           =   2085
      End
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   2880
      Top             =   3060
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   0
      Picture         =   "frmPatient_STATISTICS.frx":59A2
      Top             =   2940
      Width           =   11820
   End
End
Attribute VB_Name = "frmSTATISTICS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Check1_Click()

    frmpatient_search.Chk_date.Value = 0
    frmpatient_search.Chk_date_to_date.Value = 0
    'Check1.Value = 1
     Label1.Caption = "Enter a Name"
    txtName.Visible = True
    DTPicker1.Visible = False
    DTPicker2.Visible = False
    DTPicker3.Visible = False
End Sub

Private Sub Chk_date_Click()
    Check1.Value = 0
    Chk_date_to_date.Value = 0
    Label1.Caption = "Select a Date "
    txtName.Visible = False
    DTPicker2.Visible = False
    DTPicker3.Visible = False
    DTPicker1.Visible = True
    
    DTPicker1.Height = DTPicker2.Height
    DTPicker1.Top = DTPicker2.Top
    'frmpatient_search.Chk_date.Value = 1
    'Chk_date_to_date.Value = 0
'     Check1.Value = 0
 ' Else
   'Label1.Caption = ""
  'End If
End Sub

Private Sub Chk_date_to_date_Click()
         Label1.Caption = ""
       Check1.Value = 0
       Chk_date.Value = 0
      txtName.Visible = False
      DTPicker1.Visible = False
      'Chk_date.Visible = False
     DTPicker2.Visible = True
    DTPicker3.Visible = True
    Label1.Caption = "Select from date To date "

   
  
End Sub

Private Sub CMDEXIT_Click()

Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdPreview_Click()
Screen.MousePointer = vbHourglass
  
  If Option2.Value = True And Chk_date.Value = 1 Then
      optionbuttonval = 3
  End If
  If Option2.Value = True And Chk_date_to_date.Value = 1 Then
     optionbuttonval = 4
  End If
 If txtName.Visible = True Then
   If txtName.Text = "" Then
'      txtname.SetFocus
      MsgBox "Please Enter A Name", vbInformation, " IT, DNMIH."
      Check1.SetFocus
      
   Exit Sub
  End If
 End If

 If Check1.Value = 0 And Chk_date.Value = 0 And Me.Chk_date_to_date.Value = 0 Then
  
   MsgBox "Please Select an Search Opition", vbInformation, " IT, DNMIH."
      
   Exit Sub
  End If
 
'   If frmpatient_history.Check1.Value = 1 Then
'      optionbuttonval = 1
'    End If
'    If frmpatient_history.Chk_date.Value = 1 Then
'        optionbuttonval = 2
'           End If
'    If frmpatient_history.Chk_date_to_date.Value = 1 Then
'       optionbuttonval = 3
'    End If
'
         If Option1.Value = True Then
           rptMode = 21
           Viewer.Show vbModal
          
        Else
           rptMode = 19
           Viewer.Show vbModal
        End If
 End Sub
    
    
    
Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    DTPicker1.Visible = False
    DTPicker2.Visible = False
    DTPicker3.Visible = False
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker3.Value = Date
 'rptMode = 1
' Check1.Value = 1
' Me.Chk_date.Value = 0
' Me.Chk_date_to_date.Value = 0
' 'Combo1.Text = "Medicine"
End Sub

Private Sub Option1_Click()
 If Option1.Value = True Then
   dept.Visible = False
 End If
End Sub

Private Sub Option2_Click()
 If Option2.Value = True Then
   dept.Visible = True

   Adodc1.ConnectionString = strcn.Connection_String
      Adodc1.RecordSource = "select distinct(doc_dept) from doctor_info"
      Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        While Adodc1.Recordset.EOF = False
       dept.AddItem Adodc1.Recordset!doc_dept
            Adodc1.Recordset.MoveNext
        Wend
  dept.Text = dept.List(0)
        
    End If
 Else
    dept.Visible = False
    
 End If
End Sub

'Private Sub Option1_Click(Index As Integer)
'Select Case Index
'    Case 0
'        If Option1(0).Value = True Then
'              IntOption = 1
'
''            Option1(1).Enabled = False
'            Combo1.Enabled = False
'        Else
''            Option1(1).Enabled = True
'            Combo1.Enabled = True
'
'        End If
'    Case 1
'        If Option1(1).Value = True Then
'             IntOption = 2
'
''            Option1(1).Enabled = True
'            Combo1.Enabled = True
'        Else
''            Option1(1).Enabled = False
'            Combo1.Enabled = False
'
'        End If
'End Select
'End Sub
Private Sub txtName_Change()
  
End Sub

Private Sub txtname_GotFocus()
    txtName.BackColor = vbCyan
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
           cmdPreview_Click
      End If
End Sub

Private Sub txtname_LostFocus()
    txtName.BackColor = vbWhite
End Sub
