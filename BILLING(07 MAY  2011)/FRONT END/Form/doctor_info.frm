VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Doctors_info 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10635
   Icon            =   "doctor_info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   9390
      TabIndex        =   34
      ToolTipText     =   "CLICK TO CLOSE"
      Top             =   7410
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   8160
      TabIndex        =   33
      ToolTipText     =   "CLICK TO VIEW REPORT"
      Top             =   7410
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   6930
      TabIndex        =   32
      ToolTipText     =   "CLICK TO DELETE"
      Top             =   7410
      Width           =   1215
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "NEW"
      Height          =   375
      Left            =   5700
      TabIndex        =   15
      ToolTipText     =   "CLICK TO CLEAR"
      Top             =   7410
      Width           =   1215
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   4470
      TabIndex        =   14
      ToolTipText     =   "CLICK TO SAVE "
      Top             =   7410
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   -30
      TabIndex        =   30
      Top             =   -120
      Width           =   10725
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DOCTOR'S INFORMATION SETUP"
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
         Left            =   3270
         TabIndex        =   31
         Top             =   300
         Width           =   5235
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -30
         Picture         =   "doctor_info.frx":08CA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   10860
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6570
      Left            =   -90
      TabIndex        =   0
      Top             =   660
      Width           =   10740
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "doctor_info.frx":624C
         Height          =   2175
         Left            =   120
         TabIndex        =   25
         Top             =   4320
         Width           =   10380
         _ExtentX        =   18309
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         ForeColor       =   8388608
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
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "doctor_info.frx":6261
         Left            =   7920
         List            =   "doctor_info.frx":626B
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3180
         Width           =   2550
      End
      Begin VB.ComboBox Combo2 
         DataSource      =   "Adodc3"
         Height          =   315
         ItemData        =   "doctor_info.frx":6283
         Left            =   4560
         List            =   "doctor_info.frx":628D
         TabIndex        =   11
         Text            =   "Combo2"
         Top             =   3180
         Width           =   2100
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   10020
         TabIndex        =   5
         Top             =   870
         Width           =   465
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   7890
         TabIndex        =   9
         Top             =   2610
         Width           =   2565
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   1320
         TabIndex        =   8
         Top             =   2610
         Width           =   5325
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1320
         TabIndex        =   10
         Top             =   3180
         Width           =   2295
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   6480
         Top             =   3735
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
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
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   405
         Index           =   5
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1980
         Width           =   9135
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1320
         TabIndex        =   6
         Top             =   1440
         Width           =   9165
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   7830
         TabIndex        =   4
         Top             =   870
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   4560
         TabIndex        =   3
         Top             =   870
         Width           =   2010
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   870
         Width           =   2235
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   330
         Width           =   795
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   8475
         Top             =   3720
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
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
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Top             =   3750
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   60882947
         CurrentDate     =   38007
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   9300
         TabIndex        =   29
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Degree"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   28
         Top             =   1470
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mid Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   3660
         TabIndex        =   27
         Top             =   900
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last  Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   6690
         TabIndex        =   26
         Top             =   900
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dept."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3660
         TabIndex        =   24
         Top             =   3210
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Martial Stat."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   6720
         TabIndex        =   23
         Top             =   3240
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   270
         TabIndex        =   22
         Top             =   3720
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   270
         TabIndex        =   21
         Top             =   2700
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   20
         Top             =   3210
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   6750
         TabIndex        =   19
         Top             =   2670
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   18
         Top             =   2100
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First  Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   17
         Top             =   900
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reff. Code"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   16
         Top             =   390
         Width           =   825
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2760
      Top             =   6570
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
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
      Left            =   4410
      Top             =   7350
      Width           =   6225
   End
End
Attribute VB_Name = "Doctors_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecordSource As New ADODB.Recordset
Public Con As New MyConnection
Dim conn As New Connection
Dim cmd As New Command
Dim rs As New Recordset

'Public rptMode As Integer
Public strUid As String
Public strcn        As New MyConnection


Private Sub flush_grid()
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select refer_code,f_name,m_name,l_name,position,degree,addr,email,phone,fax,doc_dept,birth_date,marriage_status from doctor_info order by refer_code"
    Adodc1.Refresh
    
End Sub

Private Sub cmdADD_Click()
Call clear
End Sub

Private Sub CMDDELETE_Click()
Dim reply As String
    reply = MsgBox("Do you want to Delete?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
    End If
MsgBox "DELETE OPERATION IS PROHIBITED"
'Flush_Grid
'
'Dim i As Integer
'
'Conn.ConnectionString = strcn.Connection_String
'Conn.Open
'cmd.ActiveConnection = Conn
'cmd.CommandType = adCmdText
'cmd.CommandText = "delete from  doctor_info where refer_code='" & Trim(txtfield(0).Text) & "'"
'cmd.Execute
'Conn.Close
'
'On Error Resume Next
' For i = 0 To txtfield.Count
'     txtfield(i) = ""
' Next
'
' Call Flush_Grid
'
'End Sub
'
'
'Private Sub cmdExit_Click()
' Dim reply As String
'    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
'    If reply = vbYes Then
'        Unload Me
'    End If
End Sub

Private Sub CMDEXIT_Click()
  
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdPreview_Click()
  MsgBox "No such report added", vbInformation
End Sub

Private Sub cmdSave_Click()
'Dim i
''On Error Resume Next
'
'
'For i = 0 To 10
'
'  Select Case i
'
'  Case 1, , 3, 4, 5, 9
'
'      If txtfield(i) = Empty Then
'        MsgBox Label1(i) + " Requied"
'        txtfield(i).SetFocus
'        Exit Sub
'   End If
'
' End Select
'
'Next

If txtfield(1) = Empty Then
       MsgBox "First Name Requied"
        txtfield(1).SetFocus
        Exit Sub
End If
If txtfield(9) = Empty Then
       MsgBox "Position Requied"
        txtfield(9).SetFocus
        Exit Sub
End If

If txtfield(4) = Empty Then
       MsgBox "Degree Requied"
        txtfield(4).SetFocus
        Exit Sub
End If

If txtfield(5) = Empty Then
       MsgBox "Address Requied"
        txtfield(5).SetFocus
        Exit Sub
End If




    
    Call SaveDoctorInfo
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
flush_grid
'    To_Get_Valueinto_Grid

Call clear

End Sub

Private Sub SaveDoctorInfo()

    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset

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
    
    
    
If conn.State = 0 Then
        conn.Open strcn.Connection_String
End If
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
'

    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, txtfield(0).Text)
    cmd.Parameters.Append Param1 'refer_code

    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 200, txtfield(5).Text)
    cmd.Parameters.Append Param2 'addr
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 25, txtfield(8).Text)
    
    cmd.Parameters.Append Param3 'phone
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 25, txtfield(6).Text)
    
    cmd.Parameters.Append Param4 'Fax
    
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 25, txtfield(7).Text)
    cmd.Parameters.Append Param5 'E-mail
    
    Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, 10, DTPicker1.Value)
    cmd.Parameters.Append Param6 'birth_date
    
  
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 20, Combo1.Text)
    
    cmd.Parameters.Append Param7 'marriage_status
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 10, "NMIH")
    cmd.Parameters.Append Param8 'u_id
 
     Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 40, Combo2.Text)
     cmd.Parameters.Append Param9 'doc_depart
    
    Set Param10 = cmd.CreateParameter("param10", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param10 'temporary date
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 200, txtfield(4).Text)
    cmd.Parameters.Append Param11 'degree
    
    Set Param12 = cmd.CreateParameter("param12", adVarChar, adParamInput, 30, txtfield(1).Text)
    cmd.Parameters.Append Param12 'f_name
    
    Set Param13 = cmd.CreateParameter("param13", adVarChar, adParamInput, 30, txtfield(2).Text)
    cmd.Parameters.Append Param13 'l_name
    
    Set Param14 = cmd.CreateParameter("param14", adVarChar, adParamInput, 30, txtfield(3).Text)
    cmd.Parameters.Append Param14 'm_name
    
    Set Param15 = cmd.CreateParameter("param15", adInteger, adParamInput, 2, txtfield(9).Text)
    cmd.Parameters.Append Param15 'position
    
    
    
    

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL SaveDoctor_info(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set rs = cmd.Execute
    
   cmd.Properties("PLSQLRSet") = False
   If conn.State = 1 Then
      conn.Close
   End If
End Sub



Private Sub DataGrid1_Click()

'If DataGrid1.Row > 0 Then
txtfield(0).Text = DataGrid1.Columns(0)
txtfield(1).Text = DataGrid1.Columns(1)
txtfield(2).Text = DataGrid1.Columns(3)
txtfield(3).Text = DataGrid1.Columns(2)
txtfield(4).Text = DataGrid1.Columns(5)
txtfield(5).Text = DataGrid1.Columns(6)
txtfield(7).Text = DataGrid1.Columns(7)
txtfield(8).Text = DataGrid1.Columns(8)
txtfield(6).Text = DataGrid1.Columns(9)
txtfield(9).Text = DataGrid1.Columns(4)
DTPicker1.Value = DataGrid1.Columns(11)
Combo2.Text = DataGrid1.Columns(10)
Combo1.Text = DataGrid1.Columns(12)
'End If

End Sub




Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     cmdSAVE.SetFocus
  End If
End Sub

Private Sub Form_Load()
    Combo1 = Combo1.List(0)
    Combo2 = Combo2.List(0)
'
'     Adodc3.ConnectionString = strcn.Connection_String
'     Adodc3.RecordSource = "select distinct(doc_dept) from doctor_info"
'     Adodc3.Refresh
''
'    If Adodc3.Recordset.RecordCount > 0 Then
'        Adodc3.Recordset.MoveFirst
'        While Adodc3.Recordset.EOF = False
'        Combo2.AddItem Adodc3.Recordset!doc_dept
'        Adodc3.Recordset.MoveNext
'        Wend
        
        
    
'    End If
    Call flush_grid

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
End If


End Sub


Private Sub Flush_Data()

On Error Resume Next
Dim i As Integer
    Adodc2.ConnectionString = strcn.Connection_String
    Adodc2.RecordSource = "select refer_code,doc_name,addr,phone,fax,email,birth_date,marriage_status,marriage_date,u_id,dt,doc_dept from doctor_info where refer_code='" & Trim(txtfield(0).Text) & "'"
    Adodc2.Refresh
    
    If Adodc2.Recordset.RecordCount > 0 Then
        

'         txtField(6).Text = Adodc2.Recordset!birth_date
'         txtField(2).Text = Adodc2.Recordset!marriage_status
'         txtField(2).Text = Adodc2.Recordset!marriage_date

''
         For i = 0 To Adodc2.Recordset.Fields.Count
            txtfield(i).Text = Adodc2.Recordset.Fields(i).Value
            
            If i = 7 Then   '' Marital Status
                Combo1.ListIndex = (Adodc2.Recordset.Fields(i).Value)
            End If
            
            
            If i = 8 Then   '' DOB
                'MsgBox Adodc2.Recordset.Fields(i).Value
                DTPicker1.Value = (Adodc2.Recordset.Fields(i).Value)
                
            End If
           
         Next
    Else
         For i = 0 To txtfield.Count
            txtfield(i) = ""
         Next
        
    End If
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)

    If txtfield(0) <> "" And KeyAscii = 13 Then
            
        Call Flush_Data
    
    End If

End Sub

'Private Sub txtField_LostFocus(Index As Integer)
''    If txtField(Index) = 0 Then
''    Call Flush_Data
''End If
'End Sub

Private Sub clear()
Dim i
'On Error Resume Next
  

For i = 0 To 5


          txtfield(i).Text = ""
          txtfield(i).SetFocus

   
Next
txtfield(6).Text = ""
txtfield(7).Text = ""
txtfield(8).Text = ""

txtfield(9).Text = ""
 txtfield(1).SetFocus

End Sub
