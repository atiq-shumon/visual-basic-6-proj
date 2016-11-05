VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Rpt_test_info 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Test Information Statements"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   465
      Left            =   780
      Top             =   2700
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   820
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
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   3690
      TabIndex        =   5
      Top             =   2970
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "VIEW REPORT"
      Height          =   375
      Left            =   2340
      TabIndex        =   4
      Top             =   2970
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   -30
      TabIndex        =   7
      Top             =   -120
      Width           =   5145
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Test Information Statements"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   660
         TabIndex        =   8
         Top             =   210
         Width           =   3810
      End
      Begin VB.Image Image1 
         Height          =   705
         Left            =   -150
         Picture         =   "Rpt_Test_info.frx":0000
         Stretch         =   -1  'True
         Top             =   90
         Width           =   9390
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1125
      Top             =   1590
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
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
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   0
      TabIndex        =   0
      Top             =   540
      Width           =   5175
      Begin VB.ComboBox cboDeptCode 
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Rpt_Test_info.frx":5982
         Left            =   210
         List            =   "Rpt_Test_info.frx":5984
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   750
         Width           =   4785
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2910
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   165
         Width           =   780
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Rpt_Test_info.frx":5986
         Left            =   210
         List            =   "Rpt_Test_info.frx":59C6
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   1350
         Width           =   4815
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "Specific"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   1020
         TabIndex        =   2
         Top             =   330
         Width           =   1245
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "All"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   360
         Width           =   645
      End
      Begin VB.Shape Shape2 
         Height          =   345
         Left            =   180
         Top             =   300
         Width           =   4815
      End
   End
   Begin VB.Shape Shape1 
      Height          =   435
      Left            =   2280
      Top             =   2940
      Width           =   2655
   End
End
Attribute VB_Name = "Rpt_test_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()

Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdPreview_Click()
   Screen.MousePointer = vbHourglass
       rptMode = 2
       Viewer.Show vbModal

End Sub

Private Sub Combo1_Click()
     Adodc1.ConnectionString = strcn.Connection_String
     Adodc1.RecordSource = "select m_code from test_info_main where m_name='" & Trim(Combo1.Text) & "'"
     Adodc1.Refresh
     If Adodc1.Recordset.RecordCount > 0 Then
         Text1.Text = Adodc1.Recordset!M_Code
     Else
         Text1.Text = ""
     End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyEscape Then
                   Unload Me
            End If

End Sub

Private Sub Form_Load()
            rptMode = 2
            Call Load_dept_Code
            
'             Text1.Text = "01"
'           Option1(0).Value = True
'           Combo1.Text = "HAEMATOLOGY"
End Sub
Private Sub Load_dept_Code()
     
        Adodc1.ConnectionString = strcn.Connection_String
        Adodc1.RecordSource = "select distinct(dept_code)from test_info_main "
        Adodc1.Refresh
      

    If Adodc1.Recordset.RecordCount > 0 Then
         cboDeptCode.clear
         
        Adodc1.Recordset.MoveFirst
        While Adodc1.Recordset.EOF = False
         cboDeptCode.AddItem Adodc1.Recordset!dept_Code
        Adodc1.Recordset.MoveNext
        Wend
        End If
    
End Sub
Private Sub Option1_Click(Index As Integer)
Select Case Index
    Case 0
        If Option1(0).Value = True Then
          IntOption = 1

'            Option1(1).Enabled = False

            Combo1.Enabled = False
        Else
'            Option1(1).Enabled = True
            Combo1.Enabled = True

        End If
    Case 1
        If Option1(1).Value = True Then
          IntOption = 2

'            Option1(1).Enabled = True
            Combo1.Enabled = True
        Else
'            Option1(1).Enabled = False
            Combo1.Enabled = False

        End If
End Select
End Sub

