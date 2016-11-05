VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Rpt_Indoor_door_info 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4663.043
   ScaleMode       =   0  'User
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   5220
      TabIndex        =   8
      ToolTipText     =   "CLOSE"
      Top             =   4230
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   3990
      TabIndex        =   7
      ToolTipText     =   "VIEW REPORT"
      Top             =   4230
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   -30
      TabIndex        =   0
      Top             =   720
      Width           =   6660
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All Department(Outdoor)"
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
         Index           =   2
         Left            =   3060
         TabIndex        =   12
         Top             =   570
         Width           =   2550
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All Department(Indoor)"
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
         Index           =   1
         Left            =   390
         TabIndex        =   11
         Top             =   1020
         Width           =   2370
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Department Specific"
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
         Index           =   3
         Left            =   3120
         TabIndex        =   6
         Top             =   1020
         Value           =   -1  'True
         Width           =   2340
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All Department"
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
         Index           =   0
         Left            =   375
         TabIndex        =   5
         Top             =   570
         Width           =   1980
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   330
         TabIndex        =   2
         Top             =   2700
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   58327041
         CurrentDate     =   38040
      End
      Begin VB.ComboBox rptOutCombo 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
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
         ItemData        =   "Rpt_Indoor_door_info.frx":0000
         Left            =   420
         List            =   "Rpt_Indoor_door_info.frx":0002
         TabIndex        =   1
         Top             =   1920
         Width           =   5775
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4110
         TabIndex        =   3
         Top             =   2700
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   58327041
         CurrentDate     =   38040
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H008080FF&
         BorderStyle     =   2  'Dash
         BorderWidth     =   2
         Height          =   975
         Left            =   270
         Top             =   540
         Width           =   5955
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-------SELECT DEPARTMENT---------"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1410
         TabIndex        =   10
         Top             =   1620
         Width           =   3720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATE RANGE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   240
         Left            =   2460
         TabIndex        =   4
         Top             =   2760
         Width           =   1485
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1710
      Top             =   1980
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
   Begin VB.Shape Shape2 
      Height          =   465
      Left            =   3960
      Top             =   4170
      Width           =   2505
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   0
      Picture         =   "Rpt_Indoor_door_info.frx":0004
      Stretch         =   -1  'True
      Top             =   3930
      Width           =   11610
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENTAL INCOME STATEMENT"
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
      Left            =   180
      TabIndex        =   9
      Top             =   180
      Width           =   5865
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   -180
      Picture         =   "Rpt_Indoor_door_info.frx":5986
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   11610
   End
End
Attribute VB_Name = "Rpt_Indoor_door_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
'DTPicker3.Enabled = True
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
     cboShift.AddItem "Dental"
Else
cboShift.Enabled = False
End If

End Sub

Private Sub cmdExit_Click()
        Unload Me
End Sub

Private Sub cmdPreview_Click()
  Screen.MousePointer = vbHourglass
  
    Select Case IntOption1
           Case 0, 1, 2
                      IntOption = IntOption1
                      rptMode = 502
           Case 3
                rptMode = 5
'    ElseIf Option1(3).Value = True Then
'       rptMode = 5 ''indoor info
'  End If
  End Select
       Viewer.Show vbModal
  
       
End Sub

Private Sub Form_Load()
 DTPicker1.Value = Date
 DTPicker2.Value = Date
 
 IntOption1 = 3
 'Option1(0).Value = True
rptOutCombo.clear
rptOutCombo.Text = rptOutCombo.List(0)

GET_DEPT
End Sub

Private Sub GET_DEPT()

                  rptOutCombo.clear
                If Option1(3).Value = True Then
                                IntOption1 = 4
                                
             
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
                  
                  
             End If
             rptOutCombo.AddItem "Dental"
             rptOutCombo.AddItem "Emergency"
             rptOutCombo.AddItem "Physiotherapy"
             rptOutCombo.AddItem "Immunization"
  End Sub

Private Sub Option1_Click(Index As Integer)
  IntOption1 = Index
  If Option1(3).Value = True Then
     rptOutCombo.Enabled = True
  Else
    rptOutCombo.Enabled = False
  End If
End Sub
