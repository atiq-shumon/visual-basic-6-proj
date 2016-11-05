VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Rpt_advance_reg 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      ToolTipText     =   "CLOSE"
      Top             =   2580
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   3930
      TabIndex        =   9
      ToolTipText     =   "VIEW REPORT"
      Top             =   2580
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1350
      Top             =   2550
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
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   1665
      Left            =   -30
      TabIndex        =   0
      Top             =   690
      Width           =   6675
      Begin VB.Frame Frame3 
         BackColor       =   &H80000009&
         Height          =   855
         Left            =   0
         TabIndex        =   4
         Top             =   60
         Width           =   6705
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0FF&
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
            Left            =   2250
            TabIndex        =   8
            Top             =   450
            Width           =   4335
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   450
            Width           =   885
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000009&
            Caption         =   "User Wise"
            Height          =   225
            Index           =   1
            Left            =   180
            TabIndex        =   6
            Top             =   480
            Width           =   1035
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "All"
            Height          =   225
            Index           =   0
            Left            =   180
            TabIndex        =   5
            Top             =   150
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.Shape Shape2 
            Height          =   285
            Index           =   1
            Left            =   150
            Top             =   450
            Width           =   1185
         End
         Begin VB.Shape Shape2 
            Height          =   285
            Index           =   0
            Left            =   150
            Top             =   120
            Width           =   1185
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   1140
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57933825
         CurrentDate     =   38197
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   4530
         TabIndex        =   2
         Top             =   1140
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57933825
         CurrentDate     =   38197
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--------Date Range----------"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2220
         TabIndex        =   3
         Top             =   1200
         Width           =   1935
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Register"
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
      Height          =   405
      Left            =   960
      TabIndex        =   11
      Top             =   150
      Width           =   4755
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   0
      Picture         =   "Rpt_advance_reg.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11610
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   3870
      Top             =   2520
      Width           =   2565
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   0
      Picture         =   "Rpt_advance_reg.frx":5982
      Stretch         =   -1  'True
      Top             =   2340
      Width           =   11610
   End
End
Attribute VB_Name = "Rpt_advance_reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()

       Unload Me
   
End Sub

Private Sub cmdPreview_Click()
  Screen.MousePointer = vbHourglass
       rptMode = 313
       Viewer.Show vbModal
       
End Sub

Private Sub Combo1_Click()
   load_name
End Sub

Private Sub Form_Load()
 DTPicker1.Value = Date
 DTPicker2.Value = Date
 
' rptMode = 1
' Option1(0).Value = True
' Combo1.Text = "Medicine"
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
Private Sub Option1_Click(Index As Integer)
       Select Case Index
              Case 0
                    Combo1.Enabled = False
                    Text1.Enabled = False
              Case 1
                    Combo1.Enabled = True
                    Text1.Enabled = True
                    Combo1.SetFocus
                    load_user
                  
       End Select
End Sub
Private Sub load_user()
   Adodc1.ConnectionString = strcn.Connection_String
   Adodc1.RecordSource = "Select TO_NUMBER(user_id) from security ORDER BY TO_NUMBER(user_id) aSC"
   Adodc1.Refresh
   Combo1.clear
   
   If Adodc1.Recordset.RecordCount > 0 Then
      Adodc1.Recordset.MoveFirst
      Do Until Adodc1.Recordset.EOF
         Combo1.AddItem Adodc1.Recordset(0)
         Adodc1.Recordset.MoveNext
      Loop
        
   End If
End Sub
Private Sub load_name()
   Adodc1.ConnectionString = strcn.Connection_String
   Adodc1.RecordSource = "Select user_name from security where user_id='" & Trim(Combo1.Text) & "'"
   Adodc1.Refresh
   
   If Adodc1.Recordset.RecordCount > 0 Then
         Text1.Text = Adodc1.Recordset!user_name
   Else
       Text1.Text = ""
   End If
End Sub
