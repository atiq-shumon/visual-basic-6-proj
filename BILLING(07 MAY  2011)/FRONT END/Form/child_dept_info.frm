VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmChild_dept 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11805
   Icon            =   "child_dept_info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   5280
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "Adodc3"
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
   Begin VB.TextBox txtNCharge 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7530
      TabIndex        =   6
      Top             =   1020
      Width           =   825
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5250
      Top             =   6015
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5385
      Top             =   6000
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "child_dept_info.frx":08CA
      Height          =   4560
      Left            =   0
      TabIndex        =   19
      Top             =   1350
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   8043
      _Version        =   393216
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   0
      TabIndex        =   16
      Top             =   -60
      Width           =   11850
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   -30
         TabIndex        =   31
         Top             =   -30
         Width           =   12165
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PAEDIATIC DEPARTMENT INFORMATION SETUP"
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
            Left            =   1920
            TabIndex        =   32
            Top             =   240
            Width           =   7905
         End
         Begin VB.Image Image1 
            Height          =   855
            Left            =   -150
            Picture         =   "child_dept_info.frx":08DF
            Stretch         =   -1  'True
            Top             =   -30
            Width           =   12570
         End
      End
      Begin VB.TextBox txtAdmissionFee 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4920
         TabIndex        =   3
         Top             =   1080
         Width           =   930
      End
      Begin VB.TextBox txtIncubatorCharge 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5820
         TabIndex        =   4
         Top             =   1080
         Width           =   840
      End
      Begin VB.TextBox txtBloodSugar 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   9960
         TabIndex        =   9
         Top             =   1080
         Width           =   750
      End
      Begin VB.TextBox txtPhototherapy 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   9180
         TabIndex        =   8
         Top             =   1080
         Width           =   780
      End
      Begin VB.TextBox txtETCharge 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   8340
         TabIndex        =   7
         Top             =   1080
         Width           =   840
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   10665
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   38054
      End
      Begin VB.ComboBox cboChildBed 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "child_dept_info.frx":6261
         Left            =   3780
         List            =   "child_dept_info.frx":6271
         TabIndex        =   2
         Top             =   1080
         Width           =   1140
      End
      Begin VB.ComboBox CboChildName 
         Height          =   315
         ItemData        =   "child_dept_info.frx":629A
         Left            =   945
         List            =   "child_dept_info.frx":629C
         TabIndex        =   1
         Text            =   "Combo2"
         Top             =   1080
         Width           =   2850
      End
      Begin VB.TextBox TxtBCharge 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6630
         TabIndex        =   5
         Top             =   1080
         Width           =   930
      End
      Begin VB.TextBox txtChildCode 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   0
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Fee"
         Height          =   330
         Left            =   5160
         TabIndex        =   30
         Top             =   765
         Width           =   375
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Incu.Charge"
         Height          =   210
         Left            =   5640
         TabIndex        =   29
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "B. Sugar"
         Height          =   210
         Left            =   10080
         TabIndex        =   28
         Top             =   765
         Width           =   705
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "P. Therapy "
         Height          =   210
         Left            =   9240
         TabIndex        =   27
         Top             =   765
         Width           =   780
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "E.T. Charge"
         Height          =   210
         Left            =   8280
         TabIndex        =   26
         Top             =   765
         Width           =   900
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective date"
         Height          =   330
         Left            =   10770
         TabIndex        =   25
         Top             =   765
         Width           =   1005
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "N. Charge"
         Height          =   210
         Left            =   7560
         TabIndex        =   24
         Top             =   765
         Width           =   765
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Type"
         Height          =   210
         Left            =   3960
         TabIndex        =   21
         Top             =   765
         Width           =   705
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "B. C. Charge"
         Height          =   210
         Left            =   6600
         TabIndex        =   20
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   165
         Index           =   0
         Left            =   450
         TabIndex        =   18
         Top             =   765
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   1650
         TabIndex        =   17
         Top             =   765
         Width           =   420
      End
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
      Height          =   465
      Left            =   1005
      Picture         =   "child_dept_info.frx":629E
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Delete"
      Top             =   5955
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
      Height          =   465
      Left            =   510
      Picture         =   "child_dept_info.frx":6DD8
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "New"
      Top             =   5955
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
      Height          =   465
      Left            =   1515
      Picture         =   "child_dept_info.frx":7442
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Preview"
      Top             =   5955
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
      Height          =   465
      Left            =   30
      Picture         =   "child_dept_info.frx":7AAC
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Save"
      Top             =   5955
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
      Height          =   465
      Left            =   2025
      Picture         =   "child_dept_info.frx":8116
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Exit"
      Top             =   5955
      Width           =   510
   End
   Begin VB.Shape Shape1 
      Height          =   585
      Left            =   0
      Top             =   5880
      Width           =   2595
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C9AD8F&
      Caption         =   "Charge"
      Height          =   330
      Left            =   8370
      TabIndex        =   23
      Top             =   360
      Width           =   1185
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C9AD8F&
      Caption         =   "Charge"
      Height          =   330
      Left            =   8955
      TabIndex        =   22
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmChild_dept"
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

Public strUid As String
Public strcn        As New MyConnection

Private Sub cmdADD_Click()
Call Clear_form
End Sub
Private Sub Clear_form()
TxtBCharge = ""
txtAdmissionfee = ""
txtBloodSugar = ""
txtChildCode(0) = ""
txtETCharge = ""
txtNCharge = ""
txtPhototherapy = ""

End Sub

Private Sub CMDDELETE_Click()
Dim reply As String
    reply = MsgBox("Do you want to Delete?", vbQuestion + vbYesNo, "Delete...")
    If reply = vbYes Then
    If txtChildCode(0).Text = "" Then
    MsgBox "Operation code Required", vbInformation, "Warning: IT, DNMIH."
    Else
    
        'Call DeleteOperationInfo
    MsgBox "You can Only Update Rather than Delete", vbInformation + vbOKOnly, "Warning..."
        Call flush_grid
 
        Clear_form
  End If
  
End If
End Sub
Private Sub DeleteOperationInfo()
    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
        
    conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, txtChildCode(0).Text)
    cmd.Parameters.Append Param1 'Opr_code
     
     cmd.CommandText = "{CALL delete_Opr_info(?)}"
    Set rs = cmd.Execute
    
    Debug.Print cmd.CommandText
    
End Sub
Private Sub CMDEXIT_Click()
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub flush_grid()
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select * from Child_dept order by code"
    
       Adodc1.Refresh
   DataGrid1.Columns(0).Width = 530
    DataGrid1.Columns(1).Width = 2880
  DataGrid1.Columns(2).Width = 1150
   DataGrid1.Columns(3).Width = 940
   DataGrid1.Columns(4).Width = 830

    DataGrid1.Columns(5).Width = 900
    DataGrid1.Columns(6).Width = 800
     DataGrid1.Columns(7).Width = 770
     DataGrid1.Columns(8).Width = 850
    DataGrid1.Columns(9).Width = 750
    DataGrid1.Columns(10).Width = 1100

    
End Sub

Private Sub cmdSave_Click()
    If Me.txtChildCode(0) = Empty Then
        MsgBox "Main code Requied", vbInformation, "Warning: IT, DNMIH"
        txtChildCode(0).SetFocus
        Exit Sub
   End If
   If Me.CboChildName = Empty Then
        MsgBox "Charge Name Requied", vbInformation, "Warning: IT, DNMIH"
       CboChildName.SetFocus
        Exit Sub
   End If
   
   If TxtBCharge = Empty Then
        MsgBox "Baby Care  charge Requied", vbInformation, "Warning: IT, DNMIH"
        TxtBCharge.SetFocus
        Exit Sub
   End If
   
   If txtNCharge = Empty Then
        MsgBox "Neunetal  Charge Requied", vbInformation, "Warning: IT, DNMIH"
        txtNCharge.SetFocus
        Exit Sub
   End If
   If txtETCharge = Empty Then
        MsgBox "Exchange Tranfusion Charge Requied", vbInformation, "Warning: IT, DNMIH"
        txtETCharge.SetFocus
        Exit Sub
   End If
   
 If txtPhototherapy = Empty Then
        MsgBox "Photo Therapy Charge Requied", vbInformation, "Warning: IT, DNMIH"
        txtPhototherapy.SetFocus
        Exit Sub
   End If
   
If txtBloodSugar = Empty Then
        MsgBox "Blood Sugar Charge Requied", vbInformation, "Warning: IT, DNMIH"
        txtBloodSugar.SetFocus
        Exit Sub
   End If
   

    Call SaveChildInfo
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
    Call flush_grid

'Clear_form

End Sub

Private Sub SaveChildInfo()

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
    
 
  If conn.State = 0 Then
        conn.Open strcn.Connection_String
  End If
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    

    '----------------------------------------------------------------------------------
    
    
    
   
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, txtChildCode(0).Text)
    cmd.Parameters.Append Param1 'code

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 60, CboChildName.Text)
    cmd.Parameters.Append Param2 'name
    
   
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 25, cboChildBed.Text)
    cmd.Parameters.Append Param3 'bed Type
    
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, txtAdmissionfee.Text)
    cmd.Parameters.Append Param4 'admission fee
    
    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 10, txtIncubatorCharge.Text)
    cmd.Parameters.Append Param5 'txtIncubatorCharge
    
    
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 30, TxtBCharge.Text)
    cmd.Parameters.Append Param6 'baby care charge
    
    
    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 30, Trim(txtNCharge.Text))
    cmd.Parameters.Append Param7 'Neunetal charge
    
    Set Param8 = cmd.CreateParameter("param8", adDouble, adParamInput, 10, txtETCharge.Text)
    cmd.Parameters.Append Param8 'ET charge
    
    
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 10, Trim(txtPhototherapy.Text))
    cmd.Parameters.Append Param9 ''txtPhototherapy
 
'    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 10, Trim(txtBloodSugar.Text))
'    cmd.Parameters.Append Param9 ''Service_charge

    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, Trim(txtBloodSugar.Text))
    cmd.Parameters.Append Param10 'txtBloodSugar
 
    
'    Set Param9 = cmd.CreateParameter("param9", adDate, adParamInput, 10, DTPicker1.Value)
'    cmd.Parameters.Append Param9 'Effective date

   
'----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Save_Child_info(?,?,?,?,?,?,?,?,?,?)}"
     Set rs = cmd.Execute
    
    Debug.Print cmd.CommandText

 cmd.Properties("PLSQLRSet") = False

    If conn.State = 1 Then
       conn.Close
    End If
End Sub

Private Sub DataGrid1_Click()

If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "Nothing to Show", vbInformation, "Warning: IT, DNMIH"
Else

txtChildCode(0).Text = DataGrid1.Columns(0)
CboChildName.Text = DataGrid1.Columns(1)
cboChildBed.Text = DataGrid1.Columns(2)
txtAdmissionfee.Text = DataGrid1.Columns(3)
txtIncubatorCharge = DataGrid1.Columns(4)
TxtBCharge.Text = DataGrid1.Columns(5)
txtNCharge.Text = DataGrid1.Columns(6)
'
txtETCharge.Text = DataGrid1.Columns(7)
txtPhototherapy.Text = DataGrid1.Columns(8)
txtBloodSugar.Text = DataGrid1.Columns(9)
'
'DTPicker1.Value = DataGrid1.Columns(9)

End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
End If
End Sub

Private Sub Form_Load()
'
  Call flush_grid
'
'      Adodc3.ConnectionString = strcn.Connection_String
'      Adodc3.RecordSource = "select distinct(doc_dept) from doctor_info"
'      Adodc3.Refresh
'
'    If Adodc3.Recordset.RecordCount > 0 Then
'        Adodc3.Recordset.MoveFirst
'        While Adodc3.Recordset.EOF = False
'            cboOprDept.AddItem Adodc3.Recordset!doc_dept
'            Adodc3.Recordset.MoveNext
'        Wend
'    End If
'    If cboOprDept <> "" Then
'  cboOprDept = cboOprDept.List(0)
'    End If
End Sub
'
'Private Sub TxtBCharge_Change()
'If Not IsNumeric(TxtBCharge_Change.Text) Then
'            TxtBCharge_Change = ""
'End If
'
'End Sub

'Private Sub txtBloodSugar_Change()
'If Not IsNumeric(txtBloodSugar_Change.Text) Then
'            txtBloodSugar_Change = ""
'End If
'
'End Sub
'
''Private Sub txtETCharge_Change()
'If Not IsNumeric(txtETCharge_Change.Text) Then
'            txtETCharge_Change = ""
'End If
'
'End Sub

'Private Sub txtNCharge_Change()
'If Not IsNumeric(txtNCharge_Change) Then
'            txtNCharge_Change = ""
'End If
'
'End Sub
'
'Private Sub txtPhototherapy_Change()
'If Not IsNumeric(txtPhototherapy_Change.Text) Then
'            txtPhototherapy_Change = ""
'End If
'End Sub


Private Sub txtAdmissionFee_Change()
If Not IsNumeric(txtAdmissionfee.Text) Then
            txtAdmissionfee = ""
End If
End Sub

Private Sub TxtBCharge_Change()
If Not IsNumeric(TxtBCharge.Text) Then
            TxtBCharge = ""
End If

End Sub

Private Sub txtBloodSugar_Change()
If Not IsNumeric(txtBloodSugar.Text) Then
           txtBloodSugar = ""
End If

End Sub

Private Sub txtETCharge_Change()
If Not IsNumeric(txtETCharge) Then
txtETCharge = ""
End If


End Sub

Private Sub txtIncubatorCharge_Change()
If Not IsNumeric(txtIncubatorCharge.Text) Then
           txtIncubatorCharge = ""
End If

End Sub

Private Sub txtNCharge_Change()
If Not IsNumeric(txtNCharge.Text) Then
txtNCharge = ""
End If

End Sub

Private Sub txtPhototherapy_Change()
    If Not IsNumeric(txtPhototherapy.Text) Then
            txtPhototherapy = ""
     End If
End Sub
