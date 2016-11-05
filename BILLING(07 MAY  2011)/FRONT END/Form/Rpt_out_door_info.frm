VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Rpt_out_door_info 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pathological Test Report Detail"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5085
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3063.859
   ScaleMode       =   0  'User
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   60
      Picture         =   "Rpt_out_door_info.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Preview"
      Top             =   2580
      Width           =   510
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
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
      Left            =   600
      Picture         =   "Rpt_out_door_info.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exit"
      Top             =   2580
      Width           =   510
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5070
      Begin VB.ComboBox cboShift 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Rpt_out_door_info.frx":0F88
         Left            =   2580
         List            =   "Rpt_out_door_info.frx":0F8A
         TabIndex        =   17
         Top             =   1590
         Width           =   1965
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "Shift Specific"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1620
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "Employee "
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   3510
         TabIndex        =   15
         Top             =   540
         Width           =   1155
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Date Specific"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1260
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "Booth"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   1920
         TabIndex        =   12
         Top             =   120
         Width           =   765
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "Department"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   1920
         TabIndex        =   11
         Top             =   540
         Width           =   1200
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "Test Head"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   3510
         TabIndex        =   10
         Top             =   120
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "Shift"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   540
         Width           =   645
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "All"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   120
         Width           =   600
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   2130
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   38040
      End
      Begin VB.ComboBox rptOutCombo 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Rpt_out_door_info.frx":0F8C
         Left            =   360
         List            =   "Rpt_out_door_info.frx":0F8E
         TabIndex        =   1
         Top             =   870
         Width           =   4185
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   2580
         TabIndex        =   5
         Top             =   2130
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   38040
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   315
         Left            =   2580
         TabIndex        =   13
         Top             =   1230
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   38040
      End
      Begin VB.Label Label1 
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2550
         TabIndex        =   7
         Top             =   1950
         Width           =   960
      End
      Begin VB.Label lblDate 
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   6
         Top             =   1920
         Width           =   915
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1380
      Top             =   2820
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
   Begin VB.Shape Shape1 
      Height          =   585
      Left            =   -30
      Top             =   2490
      Width           =   1335
   End
End
Attribute VB_Name = "Rpt_out_door_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
DTPicker3.Enabled = True
If Check1.Value = 1 Then
DTPicker1.Enabled = False
DTPicker2.Enabled = False
Else
DTPicker1.Enabled = True
DTPicker2.Enabled = True
DTPicker3.Enabled = False
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
cboShift.Enabled = True
    Adodc1.ConnectionString = strcn.Connection_String
      Adodc1.RecordSource = "select distinct(Shift_name) from Shift_setup"
      Adodc1.Refresh
      cboShift.clear
    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        While Adodc1.Recordset.EOF = False
        cboShift.AddItem Adodc1.Recordset!shift_name
            Adodc1.Recordset.MoveNext
        Wend
        cboShift.Text = cboShift.List(0)
    End If
      
Else
cboShift.Enabled = False
End If

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
 On Error Resume Next
        rptMode = 3
       Viewer.Show vbModal
       
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload Me
End If

End Sub

Private Sub Form_Load()
 rptMode = 3 ''out door test info
 Option1(0).Value = True
rptOutCombo.clear
rptOutCombo.Text = rptOutCombo.List(0)
DTPicker1.Enabled = False
DTPicker2.Enabled = False
DTPicker3.Enabled = False
cboShift.Enabled = False



End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
    Case 0 '''all
         rptOutCombo.clear
        If Option1(0).Value = True Then
              IntOption = 1

'            Option1(1).Enabled = False
        rptOutCombo.Enabled = False
        Else
'            Option1(1).Enabled = True
        rptOutCombo.Enabled = True

        End If
    Case 1 'shift
         rptOutCombo.clear
        If Option1(1).Value = True Then
             IntOption = 2
             
'            Option1(1).Enabled = True
            rptOutCombo.Enabled = True
            cboShift.Enabled = False
            Check2.Enabled = False
            
               
            Adodc1.ConnectionString = strcn.Connection_String
      Adodc1.RecordSource = "select distinct(Shift_name) from Shift_setup"
      Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        While Adodc1.Recordset.EOF = False
        rptOutCombo.AddItem Adodc1.Recordset!shift_name
            Adodc1.Recordset.MoveNext
        Wend
        rptOutCombo.Text = rptOutCombo.List(0)
    End If
      
                     
        Else
            cboShift.Enabled = True
            Check2.Enabled = True
            
'            Option1(1).Enabled = False
            Combo1.Enabled = False
        End If
            
            
    Case 2 'test_head
    
          rptOutCombo.clear
        If Option1(2).Value = True Then
             IntOption = 3
             
'
 rptOutCombo.Enabled = True
            
            Adodc1.ConnectionString = strcn.Connection_String
      Adodc1.RecordSource = "select distinct(m_name) from pat_info_sub1_out_door"
      Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        While Adodc1.Recordset.EOF = False
        rptOutCombo.AddItem Adodc1.Recordset!m_name
            Adodc1.Recordset.MoveNext
        Wend
        rptOutCombo.Text = rptOutCombo.List(0)
    End If
        Else
'
            Combo1.Enabled = False
        End If
      Case 3   ''doc department
        rptOutCombo.clear
        If Option1(3).Value = True Then
             IntOption = 4
             
'
     rptOutCombo.Enabled = True
            
              
       Adodc1.ConnectionString = strcn.Connection_String
      Adodc1.RecordSource = "select distinct(doc_dept) from doctor_info"
      Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        While Adodc1.Recordset.EOF = False
       rptOutCombo.AddItem Adodc1.Recordset!doc_dept
            Adodc1.Recordset.MoveNext
        Wend
  rptOutCombo.Text = rptOutCombo.List(0)
        
    End If
            
            
            
            
        Else
'            Option1(1).Enabled = False
            Combo1.Enabled = False
  

        End If
        
   Case 5  ''user
        rptOutCombo.clear
        If Option1(5).Value = True Then
             IntOption = 5
             
'
     rptOutCombo.Enabled = True
            
              
       Adodc1.ConnectionString = strcn.Connection_String
      Adodc1.RecordSource = "select user_id from security"
      Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        While Adodc1.Recordset.EOF = False
       rptOutCombo.AddItem Adodc1.Recordset!user_id
            Adodc1.Recordset.MoveNext
        Wend
  rptOutCombo.Text = rptOutCombo.List(0)
        
    End If
            
            
            
            
        Else
'            Option1(1).Enabled = False
            Combo1.Enabled = False
  

        End If
        
End Select
End Sub

Private Sub Option1_GotFocus(Index As Integer)
If Index <> 1 Then
cboShift.Enabled = True
Check2.Enabled = True
End If
If Index = 1 Then
cboShift.Enabled = False
Check2.Enabled = False
End If


End Sub

Private Sub Option1_LostFocus(Index As Integer)
If Index = 1 Then
cboShift.Enabled = True
Check2.Enabled = True
End If

End Sub
