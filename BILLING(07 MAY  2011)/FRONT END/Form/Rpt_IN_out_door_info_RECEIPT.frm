VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Rpt_IN_out_door_info_RECEIPT 
   BackColor       =   &H00C9AD8F&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Receipt No wise Report"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3302.989
   ScaleMode       =   0  'User
   ScaleWidth      =   4800
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
      Left            =   30
      Picture         =   "Rpt_IN_out_door_info_RECEIPT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Preview"
      Top             =   2850
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
      Left            =   540
      Picture         =   "Rpt_IN_out_door_info_RECEIPT.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exit"
      Top             =   2850
      Width           =   510
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C9AD8F&
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   4830
      Begin VB.ComboBox cboshift1 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         ItemData        =   "Rpt_IN_out_door_info_RECEIPT.frx":0F88
         Left            =   2610
         List            =   "Rpt_IN_out_door_info_RECEIPT.frx":0F8A
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1560
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C9AD8F&
         Caption         =   "Year"
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   1920
         TabIndex        =   19
         Top             =   930
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C9AD8F&
         Caption         =   "Month"
         Height          =   285
         Index           =   6
         Left            =   360
         TabIndex        =   18
         Top             =   930
         Width           =   1155
      End
      Begin VB.ComboBox cboShift 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         ItemData        =   "Rpt_IN_out_door_info_RECEIPT.frx":0F8C
         Left            =   2610
         List            =   "Rpt_IN_out_door_info_RECEIPT.frx":0F8E
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1920
         Width           =   1965
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9AD8F&
         Caption         =   "Shift Specific"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1950
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C9AD8F&
         Caption         =   "Employee "
         Height          =   315
         Index           =   5
         Left            =   3420
         TabIndex        =   15
         Top             =   540
         Width           =   1155
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9AD8F&
         Caption         =   "Date Specific"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1590
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C9AD8F&
         Caption         =   "Booth"
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1920
         TabIndex        =   12
         Top             =   120
         Width           =   765
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C9AD8F&
         Caption         =   "Department"
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1920
         TabIndex        =   11
         Top             =   540
         Width           =   1200
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C9AD8F&
         Caption         =   "Test Head"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   3420
         TabIndex        =   10
         Top             =   120
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C9AD8F&
         Caption         =   "Shift"
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   540
         Width           =   645
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C9AD8F&
         Caption         =   "All"
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
         Top             =   2490
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   38040
      End
      Begin VB.ComboBox rptOutCombo 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Rpt_IN_out_door_info_RECEIPT.frx":0F90
         Left            =   360
         List            =   "Rpt_IN_out_door_info_RECEIPT.frx":0F92
         TabIndex        =   1
         Top             =   1200
         Width           =   4185
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   2580
         TabIndex        =   5
         Top             =   2490
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   38040
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   360
         Left            =   2610
         TabIndex        =   13
         Top             =   1560
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   16744576
         Format          =   24576001
         CurrentDate     =   38040
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C9AD8F&
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
         Top             =   2310
         Width           =   960
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C9AD8F&
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
         Top             =   2280
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
End
Attribute VB_Name = "Rpt_IN_out_door_info_RECEIPT"
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

Private Sub cmdExit_Click()

Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdPreview_Click()
 ''On Error Resume Next
       getpatient_name
   If Option1(6).Value = True Then
        rptMode = 12
       Viewer.Show vbModal
    Else
       rptMode = 11
       Viewer.Show vbModal
   End If
   
       
  
       
End Sub
Private Sub getpatient_name()
Dim var_name
On Error GoTo Errdesc
Dim conn10 As New Connection
Dim rs10 As New Recordset
Dim cmd As New Command

If conn10.State = 0 Then
conn10.ConnectionString = strcn.Connection_String
conn10.Open
End If
var_name = Rpt_IN_out_door_info_RECEIPT.rptOutCombo.Text

cmd.ActiveConnection = conn10
cmd.CommandType = adCmdText
cmd.CommandText = "select user_name from security where upper(user_id)='" & Trim(var_name) & "'"
rs10.CursorLocation = adUseClient
rs10.Open cmd.CommandText, conn10, adOpenDynamic, adLockOptimistic
    If rs10.RecordCount > 0 Then
        user_name = rs10.Fields(0)
    End If
Exit Sub
If conn10.State = 1 Then
    conn10.Close
    Set conn10 = Nothing
End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "Daffodil Software Ltd"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload Me
End If

End Sub

Private Sub Form_Load()
DTPicker3.Value = Date
 '''rptMode = 3 ''out door test info
 Option1(0).Value = True
rptOutCombo.clear
rptOutCombo.Text = rptOutCombo.List(0)
DTPicker1.Enabled = False
DTPicker2.Enabled = False
DTPicker3.Enabled = False
cboShift.Enabled = False



End Sub

Private Sub Option1_Click(Index As Integer)
 On Error Resume Next
 Select Case Index
    Case 0 '''all
        DTPicker3.Enabled = True
        Check1.Caption = "Date Specific"
        Check2.Caption = "Shift Specific"
        Option1(1).Enabled = True
       ' Option1(2).Enabled = True
        'Option1(3).Enabled = True
        'Option1(4).Enabled = True
        Option1(5).Enabled = True
         cboshift1.Visible = False
         rptOutCombo.clear
         cboShift.Enabled = False
         
        If Option1(0).Value = True Then
              IntOption = 1

'            Option1(1).Enabled = False
        rptOutCombo.Enabled = False
        
        Else
'            Option1(1).Enabled = True
        rptOutCombo.Enabled = True

        End If
    Case 1 'shift
            cboshift1.Visible = False
              rptOutCombo.clear
        If Option1(1).Value = True Then
             IntOption = 2
             
'            Option1(1).Enabled = True
            rptOutCombo.Enabled = True
            cboShift.Enabled = False
            Check2.Enabled = False
            Check2.Value = 0
            
               
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
          cboshift1.Visible = False
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
         cboshift1.Visible = False
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
        cboshift1.Visible = False
        rptOutCombo.clear
      
        If Option1(5).Value = True Then
             IntOption = 5
             
'
     rptOutCombo.Enabled = True
'     Combo1.Enabled = True
            
              
       Adodc1.ConnectionString = strcn.Connection_String
      Adodc1.RecordSource = "select user_id,user_name from security"
      Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        While Adodc1.Recordset.EOF = False
            rptOutCombo.AddItem Adodc1.Recordset!user_id
           Adodc1.Recordset.MoveNext
        Wend
  rptOutCombo.Text = rptOutCombo.List(0)
'   Combo1.Text = Combo1.List(0)
         End If
            
       
            
            
        'Else
           
'            Option1(1).Enabled = False
            'Combo1.Enabled = False
  

        End If
     Case 6       ''''MONTH
          'Dim date_string As String
           
            Dim i
            Dim var_year
            Option1(1).Enabled = False
        Option1(2).Enabled = False
        Option1(3).Enabled = False
        Option1(4).Enabled = False
        Option1(5).Enabled = False
        
            
            
            
            cboShift.Enabled = True
            cboshift1.Visible = True
           Check2.Visible = True
           cboShift.Visible = True
           cboShift.Enabled = True
           Option1(5).Enabled = False
           Option1(1).Enabled = False
           Check1.Value = 1
           Check1.Enabled = True
           Check1.Caption = "Month Specific"
           Check2.Value = 0
           Check2.Enabled = True
           DTPicker1.Enabled = False
           DTPicker2.Enabled = False
           DTPicker3.Enabled = False
           cboShift.Enabled = True
           cboshift1.clear
           
           cboshift1.List(0) = "January"
           cboshift1.List(1) = "February"
           cboshift1.List(2) = "March"
           cboshift1.List(3) = "April"
           cboshift1.List(4) = "May"
           cboshift1.List(5) = "June"
           cboshift1.List(6) = "July"
           cboshift1.List(7) = "August"
           cboshift1.List(8) = "September"
           cboshift1.List(9) = "October"
           cboshift1.List(10) = "November"
           cboshift1.List(11) = "December"
           Check2.Caption = "Year Specific"
           cboshift1 = cboshift1.List(0)
           
           Check2.Value = 1
           Check2.Enabled = True
           
           cboShift.clear
             var_year = 1999
            For i = 0 To 100
               cboShift.List(i) = var_year
               var_year = var_year + 1
            Next i
          cboShift = cboShift.List(0)
        
         
    Case 7
            Dim counter
            Dim var_year_counter
            cboshift1.Visible = True
            Check1.Value = 1
             Check1.Caption = "Year Specific"
           Check1.Enabled = True
           Check2.Value = 0
           Check2.Enabled = False
           DTPicker1.Enabled = False
           DTPicker2.Enabled = False
           DTPicker3.Enabled = False
           cboShift.Enabled = False
            cboshift1.clear
             var_year_counter = 1999
            For counter = 0 To 100
               
               cboshift1.List(counter) = var_year_counter
               var_year_counter = var_year_counter + 1
            Next counter
             cboshift1 = cboshift1.List(0)
         Check2.Caption = "Shift Specific"
         Check2.Value = 1
         cboShift.Enabled = False
            
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
