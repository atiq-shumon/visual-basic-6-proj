VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   39
      Top             =   6360
      Width           =   13365
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Software Programmer, IT Division, DNMIH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2940
         TabIndex        =   41
         Top             =   60
         Width           =   4725
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Developed && Maintenanced by:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   150
         TabIndex        =   40
         ToolTipText     =   "Developed && Maintenanced by:A.N.M. Atiqur Rahman,Software Programmer, IT Division, DNMIH"
         Top             =   90
         Width           =   2715
      End
   End
   Begin VB.Frame Frame13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   0
      TabIndex        =   1
      Top             =   630
      Width           =   10155
      Begin VB.TextBox txtPatDept 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1530
         Width           =   2445
      End
      Begin VB.TextBox txtBedType 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1530
         Width           =   2955
      End
      Begin VB.TextBox txtFiscalYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   630
         Width           =   2445
      End
      Begin VB.TextBox txtAdvanceRelease 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   7290
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   0
         Top             =   1530
         Width           =   2460
      End
      Begin VB.TextBox txtReg 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   630
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker DT_TM 
         Height          =   330
         Left            =   8745
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Delevary Time"
         Top             =   630
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "HH:MM:SS"
         Format          =   58916866
         UpDown          =   -1  'True
         CurrentDate     =   37163
      End
      Begin MSComCtl2.DTPicker Dt 
         Height          =   330
         Left            =   7245
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   630
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   49152
         Format          =   58916865
         CurrentDate     =   37114
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fiscal Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   4560
         TabIndex        =   17
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   4575
         TabIndex        =   9
         Top             =   1275
         Width           =   1215
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cabin / Ward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   390
         TabIndex        =   8
         Top             =   1275
         Width           =   1365
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reg No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   390
         TabIndex        =   7
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Admission Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   7245
         TabIndex        =   6
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Previous  Advance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   7290
         TabIndex        =   5
         Top             =   1275
         Width           =   1965
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   0
      TabIndex        =   20
      Top             =   2580
      Width           =   10125
      Begin VB.ComboBox cboDMY 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Edit_patient_information.frx":0000
         Left            =   990
         List            =   "Edit_patient_information.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1380
         Width           =   765
      End
      Begin VB.CheckBox chkCareOf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Care Of"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   270
         Left            =   8640
         MaskColor       =   &H00FFFF80&
         TabIndex        =   36
         Top             =   285
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   4620
         MaxLength       =   50
         TabIndex        =   28
         Top             =   585
         Width           =   5130
      End
      Begin VB.TextBox txtAddr 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   300
         MaxLength       =   100
         TabIndex        =   27
         Top             =   2115
         Width           =   9540
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   300
         MaxLength       =   50
         TabIndex        =   26
         Top             =   585
         Width           =   4065
      End
      Begin VB.CheckBox chkHusband 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6540
         TabIndex        =   25
         Top             =   315
         Width           =   195
      End
      Begin VB.CheckBox chkFather 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4635
         TabIndex        =   24
         Top             =   315
         Value           =   1  'Checked
         Width           =   180
      End
      Begin VB.ComboBox Combo5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Edit_patient_information.frx":001A
         Left            =   3030
         List            =   "Edit_patient_information.frx":002D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1380
         Width           =   1305
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   270
         MaxLength       =   17
         TabIndex        =   22
         Top             =   1380
         Width           =   705
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Edit_patient_information.frx":005B
         Left            =   2070
         List            =   "Edit_patient_information.frx":0065
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1380
         Width           =   945
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y/M/D"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   20
         Left            =   990
         TabIndex        =   38
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Husband's Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   6765
         TabIndex        =   35
         Top             =   285
         Width           =   1800
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Father's Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   4920
         TabIndex        =   34
         Top             =   285
         Width           =   1530
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   315
         TabIndex        =   33
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   2100
         TabIndex        =   32
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   300
         TabIndex        =   31
         Top             =   330
         Width           =   630
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   330
         TabIndex        =   30
         Top             =   1830
         Width           =   885
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   3030
         TabIndex        =   29
         Top             =   1080
         Width           =   885
      End
   End
   Begin VB.CommandButton CMDEXIT 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   8700
      TabIndex        =   15
      Top             =   5730
      Width           =   1215
   End
   Begin VB.CommandButton CmdForward 
      Caption         =   "FORWARD"
      Height          =   375
      Left            =   7470
      TabIndex        =   14
      Top             =   5730
      Width           =   1215
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "BACKWARD"
      Height          =   375
      Left            =   6240
      TabIndex        =   13
      Top             =   5730
      Width           =   1215
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   5010
      TabIndex        =   12
      Top             =   5730
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   5730
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
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10305
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EDIT PATIENT INFORMATION"
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
         Left            =   3300
         TabIndex        =   11
         Top             =   120
         Width           =   4755
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -60
         Picture         =   "Edit_patient_information.frx":006F
         Top             =   -90
         Width           =   11820
      End
   End
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   4950
      Top             =   5670
      Width           =   4995
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public checkNameIndicator As Integer

Private Sub Command1_Click()

End Sub

Private Sub chkCareOf_Click()
 If chkCareOf.Value = 1 Then
    txtPatFatherName = "C/O:"
    chkFather.Value = 0
    chkHusband.Value = 0
    chkFather.ForeColor = vbWhite
  Else
    chkFather.Enabled = True
    chkFather.Value = 1
    chkHusband.ForeColor = vbWhite
    chkFather.ForeColor = &HFFFF80
   End If
End Sub

Private Sub chkFather_Click()
  If chkFather.Value = 1 Then
    chkHusband.Value = 0
    chkCareOf.Value = 0
    chkFather.ForeColor = &HFFFF80
    txtPatFatherName = "S/D/O:"
  Else
    chkHusband.Enabled = True
    chkHusband.ForeColor = &HFFFF80
    chkHusband.Value = 1
    txtPatFatherName = "W/O:"
  End If
End Sub

Private Sub chkHusband_Click()
 If chkHusband.Value = 1 Then
     chkFather.Value = 0
     chkCareOf.Value = 0
     chkFather.ForeColor = vbWhite
  Else
    chkFather.Enabled = True
    chkFather.Value = 0
    chkCareOf.Value = 1
    chkHusband.ForeColor = vbWhite
    chkFather.ForeColor = &HFFFF80

   End If
End Sub

Private Sub CmdBack_Click()
      txtReg = txtReg - 1
      LOAD_PAT_INFO
      LOAD_PAT_BED
End Sub

Private Sub CMDDELETE_Click()

End Sub

Private Sub cmdExit_Click()
  Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub CmdForward_Click()
  txtReg = txtReg + 1
      LOAD_PAT_INFO
      LOAD_PAT_BED
End Sub

Private Sub CMDREPORT_Click()

End Sub

Private Sub Command14_Click()
    
  
End Sub
Private Sub EDIT_PAT_INFO()
'  On Error GoTo ErrDes
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
    
   If Conn.State = 0 Then
      Conn.Open strcn.Connection_String
  End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 60, Trim(TxtName.Text))
    cmd.Parameters.Append Param1 'patient_name

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 60, Trim(Text4.Text))
    cmd.Parameters.Append Param2 'guardgian_name
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 200, Trim(txtAddr.Text))
    cmd.Parameters.Append Param3 'pat_addr1
        
    Set Param4 = cmd.CreateParameter("param4", adSingle, adParamInput, 5, Trim(txtAge.Text))
    cmd.Parameters.Append Param4 'age
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 3, Trim(cboDMY.Text))
    cmd.Parameters.Append Param5 'Y/M/D
   
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 6, Trim(Combo3.Text))
    
    cmd.Parameters.Append Param6 'Sex
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 10, Trim(Combo5.Text))
    cmd.Parameters.Append Param7 'Religion
    
    
    Set Param8 = cmd.CreateParameter("param8", adInteger, adParamInput, 10, checkNameIndicator)
    cmd.Parameters.Append Param8 'check_value----father's name or husband name OR CARE OF
    
    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 20, txtReg.Text)
    cmd.Parameters.Append Param9 'registrat9ion number
    
    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 10, frmReg_for_EDIT_PAT.CBOYRCODE)
    cmd.Parameters.Append Param10 'YEAR CODE

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Indoor_EDIT_Patient_info(?,?,?,?,?,?,?,?,?,?)}"
    
   Set RS = cmd.Execute
 
 cmd.Properties("PLSQLRSet") = False
If Conn.State = 1 Then
       Conn.Close
       Set Conn = Nothing
 End If

End Sub
Private Sub Command15_Click()
 

End Sub

Private Sub cmdSave_Click()
  If TxtName = "" Then
          MsgBox "Patient Name Required", vbInformation, "Warning..."
          TxtName.SetFocus
           Exit Sub
  End If
    If txtAddr = "" Then
        MsgBox "Address Required", vbInformation, "Warning..."
        txtAddr.SetFocus
        Exit Sub
    End If
  If txtAge = "" Then
        MsgBox "Age Required", vbInformation, "Warning..."
        txtAge.SetFocus
        Exit Sub
    End If
    
  If chkFather.Value = 1 Then
      checkNameIndicator = 1 ''father
  ElseIf chkHusband.Value = 1 Then ''husband
     checkNameIndicator = 0  ''husband
  Else
    checkNameIndicator = 2
  End If
  
    
      Call EDIT_PAT_INFO
    
    MsgBox "Operation Successful", vbInformation, " IT, DNMIH."
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys Chr(9)
   End If
End Sub

Private Sub Form_Load()
      txtReg = frmReg_for_EDIT_PAT.txtReg_noOpr
      txtFiscalYear = frmReg_for_EDIT_PAT.CBOYRCODE
      
      LOAD_PAT_INFO
      LOAD_PAT_BED
      LOAD_PAT_DEPT
      LOAD_ADVANCE
End Sub
  Private Sub LOAD_ADVANCE()
   Dim Conn As New ADODB.Connection
   Dim RS As New ADODB.Recordset
   Dim cmd As New ADODB.Command
  
   Conn.ConnectionString = strcn.Connection_String
   Conn.Open
   cmd.ActiveConnection = Conn
   cmd.CommandType = adCmdText
   cmd.CommandText = "select  nvl(sum(advance),0) as advance  From advance Where in_reg_no ='" & Trim(frmReg_for_EDIT_PAT.txtReg_noOpr.Text) & "' AND YRCODE= '" & Trim(frmReg_for_EDIT_PAT.CBOYRCODE.Text) & "'"
      
   cmd.Properties("iRowsetChange") = True
   cmd.Properties("updatability") = 7
   RS.CursorLocation = adUseClient
   RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
   cmd.Properties("iRowsetChange") = False

   If IsNull(RS!advance) = True Then
       txtAdvanceRelease = 0
   Else
       txtAdvanceRelease = RS!advance
   End If
   Conn.Close
   Set Conn = Nothing
   Set RS = Nothing
   Set cmd = Nothing
   
End Sub

  Private Sub LOAD_PAT_INFO()
     Dim CHECK_VAR
     Adodc1.ConnectionString = strcn.Connection_String
     Adodc1.RecordSource = "select pat_name,pat_guard_name,sex,AGE,Religion,addr1,phone,check_name,Y_M_D  from in_door_Pat_Info_Main where in_reg_no='" & Trim(txtReg) & "' AND YRCODE='" & frmReg_for_EDIT_PAT.CBOYRCODE & "'"
     Adodc1.Refresh
    
    If Adodc1.Recordset.RecordCount > 0 Then
       If Not IsNull(Adodc1.Recordset!pat_name) Then
                     TxtName = Adodc1.Recordset!pat_name
       End If
    End If
    
    If Adodc1.Recordset.RecordCount > 0 Then
        If Not IsNull(Adodc1.Recordset!pat_guard_name) Then
                     Text4 = Adodc1.Recordset!pat_guard_name
        End If

    End If
    
     If Adodc1.Recordset.RecordCount > 0 Then
        If Not IsNull(Adodc1.Recordset!sex) Then
                     Combo3 = Adodc1.Recordset!sex
        End If
        cboDMY = Adodc1.Recordset!Y_M_D
     End If
    
    If Adodc1.Recordset.RecordCount > 0 Then
       If Not IsNull(Adodc1.Recordset!age) Then
                     txtAge = Adodc1.Recordset!age
        End If

    End If
    If Adodc1.Recordset.RecordCount > 0 Then
       If Not IsNull(Adodc1.Recordset!religion) Then
                     Combo5 = Adodc1.Recordset!religion
       End If
    End If
    If Adodc1.Recordset.RecordCount > 0 Then
          If Not IsNull(Adodc1.Recordset!addr1) Then
                     txtAddr = Adodc1.Recordset!addr1
          End If

    End If
    
    If Adodc1.Recordset.RecordCount > 0 Then
           If Not IsNull(Adodc1.Recordset!phone) Then
                     txtPhone = Adodc1.Recordset!phone
           Else
             txtPhone = ""
           End If
      End If
      
      
    If Adodc1.Recordset.RecordCount > 0 Then
           If Not IsNull(Adodc1.Recordset!CHECK_NAME) Then
                     CHECK_VAR = Adodc1.Recordset!CHECK_NAME
                     
           End If
      End If
      
      
'    If CHECK_VAR = 1 Then
'       Check1.Value = 1
'       Check2.Value = 0
'    Else
'      Check1.Value = 0
'    End If
'    If CHECK_VAR = 0 Then
'       Check2.Value = 1
'    Else
'      Check2.Value = 0
'    End If
'
    
End Sub
Private Sub LOAD_PAT_DEPT()
  Dim Conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim RS As New ADODB.Recordset
  
  If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
    End If
        cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        cmd.CommandText = "select doc_dept,serial_no  From INDOOR_PAT_DEPT_INFO Where in_reg_no ='" & Trim(frmReg_for_EDIT_PAT.txtReg_noOpr.Text) & "' AND YRCODE ='" & Trim(frmReg_for_EDIT_PAT.CBOYRCODE.Text) & "'  AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(frmReg_for_EDIT_PAT.txtReg_noOpr.Text) & "' AND YRCODE='" & Trim(frmReg_for_EDIT_PAT.CBOYRCODE) & "')"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        RS.CursorLocation = adUseClient

        RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False
          
       If RS.RecordCount > 0 Then
         txtPatDept = "" & RS!doc_dept
'         var_dept_serial_no = rs!SERIAL_NO
       End If
       
       If Conn.State = 1 Then
        Conn.Close
        Set Conn = Nothing
        Set RS = Nothing
        Set cmd = Nothing
     End If

End Sub
Private Sub LOAD_PAT_BED()
  Dim Conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim RS As New ADODB.Recordset
  
  If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
    End If
        cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        cmd.CommandText = "select BED_TYPE,Bed_type_no,BED_NO,extra_bed_flag ,SERIAL_NO From Indoor_pat_bed_info Where in_reg_no ='" & Trim(frmReg_for_EDIT_PAT.txtReg_noOpr.Text) & "' AND YRCODE ='" & Trim(frmReg_for_EDIT_PAT.CBOYRCODE.Text) & "'  AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM Indoor_pat_bed_info WHERE in_reg_no ='" & Trim(frmReg_for_EDIT_PAT.txtReg_noOpr.Text) & "' AND YRCODE='" & Trim(frmReg_for_EDIT_PAT.CBOYRCODE) & "')"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        RS.CursorLocation = adUseClient

        RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False
          
       If RS.RecordCount > 0 Then
         txtbedType = "" & RS!Bed_type & " -  " & RS!bed_TYPE_no & " -  " & RS!bed_no
'         txtExtraBedFlag = rs!Extra_bed_flag
'         VAR_CUR_BED_SERIAL_NO = rs!SERIAL_NO
       End If
       
       If Conn.State = 1 Then
        Conn.Close
        Set Conn = Nothing
        Set RS = Nothing
        Set cmd = Nothing
     End If

End Sub
    
Private Sub txtAge_Change()
  If Not IsNumeric(txtAge) Then
     txtAge = ""
  End If
     
End Sub
