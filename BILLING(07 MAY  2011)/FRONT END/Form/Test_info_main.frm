VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Test_info_main 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13290
   Icon            =   "Test_info_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   13290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   28
      Top             =   8340
      Width           =   13365
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
         TabIndex        =   30
         ToolTipText     =   "Developed && Maintenanced by:A.N.M. Atiqur Rahman,Software Programmer, IT Division, DNMIH"
         Top             =   90
         Width           =   2715
      End
      Begin VB.Label Label12 
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
         TabIndex        =   29
         Top             =   60
         Width           =   4725
      End
   End
   Begin VB.CommandButton cmdUpdateAll 
      Caption         =   "UPDATE ALL"
      Height          =   375
      Index           =   2
      Left            =   9390
      TabIndex        =   26
      ToolTipText     =   "DELETE"
      Top             =   7890
      Width           =   1245
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   630
      Top             =   7980
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Test_info_main.frx":08CA
      Height          =   5610
      Left            =   0
      TabIndex        =   17
      Top             =   2190
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   9895
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   -2147483624
      BorderStyle     =   0
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   6
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
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Index           =   1
      Left            =   5700
      TabIndex        =   10
      ToolTipText     =   "SAVE DATA"
      Top             =   7890
      Width           =   1215
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "NEW"
      Height          =   375
      Left            =   6930
      TabIndex        =   9
      ToolTipText     =   "NEW ENTRY"
      Top             =   7890
      Width           =   1215
   End
   Begin VB.CommandButton CMDDELETE 
      Caption         =   "DELETE"
      Height          =   375
      Index           =   4
      Left            =   8160
      TabIndex        =   11
      ToolTipText     =   "DELETE"
      Top             =   7890
      Width           =   1215
   End
   Begin VB.CommandButton CMDREPORT 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   10680
      TabIndex        =   12
      ToolTipText     =   "VIEW REPORT"
      Top             =   7890
      Width           =   1215
   End
   Begin VB.CommandButton CMDEXIT 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   11880
      TabIndex        =   13
      ToolTipText     =   "CLOSE"
      Top             =   7890
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3480
      Top             =   7440
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
      Left            =   2400
      Top             =   7440
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
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   0
      TabIndex        =   14
      Top             =   -30
      Width           =   13395
      Begin VB.TextBox txtfield 
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   5
         Left            =   10080
         TabIndex        =   7
         Text            =   "0"
         Top             =   1830
         Width           =   1125
      End
      Begin VB.TextBox txtfield 
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   6
         Left            =   11220
         TabIndex        =   8
         Text            =   "0"
         Top             =   1830
         Width           =   1125
      End
      Begin VB.TextBox txtfield 
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   3
         Left            =   7800
         TabIndex        =   5
         Text            =   "0"
         Top             =   1830
         Width           =   1125
      End
      Begin VB.TextBox txtfield 
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   4
         Left            =   8940
         TabIndex        =   6
         Text            =   "0"
         Top             =   1830
         Width           =   1125
      End
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
         ItemData        =   "Test_info_main.frx":08DF
         Left            =   270
         List            =   "Test_info_main.frx":08E1
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1050
         Width           =   1155
      End
      Begin VB.TextBox txtfield 
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   2
         Left            =   6660
         TabIndex        =   4
         Text            =   "0"
         Top             =   1830
         Width           =   1125
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Edit"
         Height          =   345
         Left            =   12330
         TabIndex        =   24
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdShowAll 
         Caption         =   "SHOW ALL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8190
         TabIndex        =   23
         Top             =   1020
         Width           =   2475
      End
      Begin VB.TextBox txtMainTitle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1050
         Width           =   5775
      End
      Begin VB.ComboBox cboMainCode 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   795
         Left            =   0
         TabIndex        =   20
         Top             =   -60
         Width           =   22545
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEST INFORMATION SETUP"
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
            Left            =   3870
            TabIndex        =   21
            Top             =   240
            Width           =   4755
         End
         Begin VB.Image Image1 
            Height          =   690
            Left            =   0
            Picture         =   "Test_info_main.frx":08E3
            Stretch         =   -1  'True
            Top             =   60
            Width           =   13410
         End
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   3
         Top             =   1830
         Width           =   5265
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   285
         TabIndex        =   2
         Top             =   1830
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Out-Case"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11220
         TabIndex        =   34
         Top             =   1605
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OPD"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10470
         TabIndex        =   33
         Top             =   1605
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Free"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9360
         TabIndex        =   32
         Top             =   1605
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paying"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7920
         TabIndex        =   31
         Top             =   1605
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   27
         Top             =   780
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Cabin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6690
         TabIndex        =   25
         Top             =   1605
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Name"
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
         Left            =   1800
         TabIndex        =   19
         Top             =   1605
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Code"
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
         Index           =   2
         Left            =   270
         TabIndex        =   18
         Top             =   1605
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M.Code"
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
         Index           =   0
         Left            =   1500
         TabIndex        =   16
         Top             =   825
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M.Test Name"
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
         Index           =   1
         Left            =   2580
         TabIndex        =   15
         Top             =   825
         Width           =   1140
      End
   End
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   5610
      Top             =   7830
      Width           =   7545
   End
End
Attribute VB_Name = "Test_info_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecordSource As New ADODB.Recordset
Public Con As New MyConnection
Dim Conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset

'Public rptMode As Integer
Public strUid As String
Public strcn        As New MyConnection

Private Sub cboDeptCode_Click()
 Load_Main_Code (cboDeptCode)
 cboMainCode = cboMainCode.List(0)
End Sub

Private Sub cboMainCode_Click()
     Adodc3.ConnectionString = strcn.Connection_String
     Adodc3.RecordSource = "select a.m_name   from test_info_main a where to_char(a.m_code)='" & Trim(cboMainCode.Text) & "'"
     Adodc3.Refresh
        
     If Adodc3.Recordset.RecordCount > 0 Then
             Adodc3.Recordset.MoveFirst
        While Adodc3.Recordset.EOF = False
           txtMainTitle = Adodc3.Recordset!M_NAME
           Adodc3.Recordset.MoveNext
        Wend
        Else
          txtMainTitle.Text = ""
       
       End If
          
    DataGrid1.ClearFields
     
'     txtMainTitle.Text = cboMainName.List(0)
End Sub



Private Sub cboMainCode_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtfield(0).SetFocus
   End If
End Sub

Private Sub chkDoneInNational_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     cmdSave_Click (1)
   End If
End Sub

Private Sub cmdADD_Click()
Call Clear_form
txtfield(0).SetFocus

End Sub
Private Sub CMDDELETE_Click(Index As Integer)
Dim reply As String
    reply = MsgBox("Do you want to Delete?", vbQuestion + vbYesNo, "Delete...")
    If reply = vbYes Then
    If txtfield(0).Text = "" Then
     MsgBox "Sub Code Required", vbInformation, "Warning: IT, DNMIH "
     Exit Sub
     End If
     
Call SaveTestInfo(2)
End If
Call flush_grid(1)
Call Clear_form

End Sub
Private Sub deleteTestInfo()

    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    

    
  If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
  End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    

    '----------------------------------------------------------------------------------
    
    
    
    ''''''''''''para meter for test info main
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 2, txtfield(0).Text)
    cmd.Parameters.Append Param1 'm_code

    
    
    
    
    
    '''''''''''''''parameter for test_info_sub''''''''''''''''''''
    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, txtfield(2).Text)
    cmd.Parameters.Append Param2 'Sub code
    
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL delete_Test_info_main(?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

End Sub
Private Sub cmdExit_Click()
  Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub flush_grid(MODE As Integer)
    If MODE = 1 Or MODE = 2 Then ''ALL
        Adodc1.ConnectionString = strcn.Connection_String
        Adodc1.RecordSource = "select b.s_code,b.s_name TITLE,b.charge_Cabin CABIN,b.Charge_Paying PAYING,b.charge_free FREE,b.charge_OPD OPD,b.charge_OutCase OUTCASE from test_info_sub b where b.dept_Code='" & cboDeptCode & "' and b.m_code='" & cboMainCode & "' order by B.s_code asc"
        Adodc1.Refresh
'
'   ElseIf MODE = 3 Then
'        Adodc1.ConnectionString = strcn.Connection_String
'        Adodc1.RecordSource = "select b.s_code,b.s_name,b.type,b.s_code_sub_code as Case,b.charge,b.service_charge,PATIENT_BENEFIT_CHARGE P_BENEFIT,b.charge+b.service_charge+PATIENT_BENEFIT_CHARGE TOTAL,DECODE(b.test_done_in_national,1,'Yes',' ') NATIONAL_TEST from test_info_sub b where  B.M_CODE='" & cboMainCode & "'  AND B.S_CODE LIKE '" & txtfield(0) & "%'  order by  B.s_code,UNIQUE_ID asc"
'        Adodc1.Refresh
'
   End If
'
'
'
'
    DataGrid1.Columns(0).Width = 1000
    DataGrid1.Columns(1).Width = 5300
    DataGrid1.Columns(2).Width = 1200
    DataGrid1.Columns(3).Width = 1200
    DataGrid1.Columns(4).Width = 1200
    DataGrid1.Columns(5).Width = 1100
    DataGrid1.Columns(6).Width = 1100
    DataGrid1.Columns(6).Width = 1100
'
'    With DataGrid1
'        .Columns(4).Caption = "Popular"
'        .Columns(5).Caption = "National"
'    End With
'
    
    
End Sub

Private Sub cmdPreview_Click()

End Sub

Private Sub CMDREPORT_Click()
   Rpt_test_info.Show 1
End Sub
Private Sub cmdSave_Click(Index As Integer)
   Dim i
  
    If Len(cboMainCode) = 0 Then
        MsgBox "Main code Requied"
        cboMainCode.SetFocus
        Exit Sub
   End If
   If txtfield(0) = Empty Then
        MsgBox "Sub code Requied"
        txtfield(0).SetFocus
        Exit Sub
   End If
   
   If txtfield(1) = Empty Then
        MsgBox "Sub Name Requied"
        txtfield(1).SetFocus
        Exit Sub
   End If
   
   If txtfield(2) = Empty Then
        MsgBox "Cabin Charge Requied"
        txtfield(2).SetFocus
        txtfield(2).Text = 0
        Exit Sub
   End If
    If txtfield(3) = Empty Then
        MsgBox "Paying Charge Requied"
        txtfield(3).SetFocus
        txtfield(3).Text = 0
        Exit Sub
   End If
    If txtfield(4) = Empty Then
        MsgBox "FreeBed Charge Requied"
        txtfield(4).SetFocus
        txtfield(4).Text = 0
        Exit Sub
   End If
    If txtfield(5) = Empty Then
        MsgBox "OPDCharge Requied"
        txtfield(5).SetFocus
        txtfield(5).Text = 0
        Exit Sub
   End If
    If txtfield(6) = Empty Then
        MsgBox "Out Case Charge Requied"
        txtfield(6).SetFocus
        txtfield(6).Text = 0
        Exit Sub
   End If
  Adodc2.ConnectionString = strcn.Connection_String
  Adodc2.RecordSource = "select S_CODE from test_info_sub where dept_code='" & cboDeptCode.Text & "' and m_code='" & cboMainCode.Text & "' and s_code='" & txtfield(0).Text & "' "
  Adodc2.Refresh
   
   If Index = 1 And Adodc2.Recordset.RecordCount = 0 Then
       Call SaveTestInfo(Index)
       MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
   ElseIf Index = 2 And Adodc2.Recordset.RecordCount > 0 Then
       Call SaveTestInfo(Index)
       MsgBox "Update successfull", vbInformation + vbOKOnly, "Update..."
   End If
Call flush_grid(1)
 txtfield(0).SetFocus
End Sub
Private Sub SaveTestInfo(MODE As Integer)

    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param0 As New Parameter
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
    

    
If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    

    '----------------------------------------------------------------------------------
          
    ''''''''''''para meter for test info main
    Set Param0 = cmd.CreateParameter("param0", adInteger, adParamInput, 2, MODE)
    cmd.Parameters.Append Param0 'mode

    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 3, cboDeptCode.Text)
    cmd.Parameters.Append Param1 'dept_code

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 3, cboMainCode.Text)
    cmd.Parameters.Append Param2 'm_code
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 5, txtfield(0).Text)
    cmd.Parameters.Append Param3 's_code
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 200, txtfield(1).Text)
    cmd.Parameters.Append Param4 's_title

    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 8, txtfield(2).Text)
    cmd.Parameters.Append Param5 'charge cabin
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 8, txtfield(3).Text)
    cmd.Parameters.Append Param6 'charge paying
  
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 8, txtfield(4).Text)
    cmd.Parameters.Append Param7 'charge free

    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 8, txtfield(5).Text)
    cmd.Parameters.Append Param8 'charge OPD
      
    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 8, txtfield(6).Text)
    cmd.Parameters.Append Param9 'charge outcase
      
    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 10, "SProg")
    cmd.Parameters.Append Param10 'user id
 
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL s_u_d_Test_info_sub(?,?,?,?,?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
     Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    
    
    Dim i As Integer
     For i = 1 To 6
        txtfield(i).Text = ""
     Next i
    
    

End Sub

Private Sub Command1_Click()
  
End Sub

Private Sub cmdShow_Click()
 If Len(cboMainCode) = 0 Then
    cboMainCode.SetFocus
    Exit Sub
 ElseIf Len(txtfield(0)) = 0 Then
    MsgBox "PUT TEST CODE HERE", vbInformation, "IT, DNMIH"
    txtfield(2).SetFocus
 End If
  flush_grid (3)
End Sub

Private Sub cmdShowAll_Click()
 If Len(cboMainCode) = 0 Then
    Exit Sub
 End If
  flush_grid (2)
End Sub

Private Sub cmdUpdate_Click()
  cmdSave_Click (2)
  txtfield(0).SetFocus
End Sub

Private Sub cmdUpdateAll_Click(Index As Integer)
  cmdSave_Click (Index)
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     txtfield(4).SetFocus
  End If
End Sub

Private Sub DataGrid1_Click()

If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "Nothing to Show", vbInformation, "Warning: IT, DNMIH"
Else


txtfield(0).Text = DataGrid1.Columns(0)
txtfield(1).Text = DataGrid1.Columns(1)
txtfield(2).Text = DataGrid1.Columns(2)
txtfield(3).Text = DataGrid1.Columns(3)
txtfield(4).Text = DataGrid1.Columns(4)
txtfield(5).Text = DataGrid1.Columns(5)
txtfield(6).Text = DataGrid1.Columns(6)
End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
End If
End Sub

Private Sub Form_Load()

'
    
    Call Load_dept_Code
    Call flush_grid(1)
     cboDeptCode = "PAT"
    Call Load_Main_Code(cboDeptCode)
    cboMainCode = cboMainCode.List(0)
    

End Sub
Private Sub Load_dept_Code()
     
        Adodc2.ConnectionString = strcn.Connection_String
        Adodc2.RecordSource = "select distinct(dept_code)from test_info_main "
        Adodc2.Refresh
      

    If Adodc2.Recordset.RecordCount > 0 Then
         cboDeptCode.clear
         
        Adodc2.Recordset.MoveFirst
        While Adodc2.Recordset.EOF = False
         cboDeptCode.AddItem Adodc2.Recordset!dept_Code
        Adodc2.Recordset.MoveNext
        Wend
        End If
    
End Sub
Private Sub Load_Main_Code(deptCode As String)
     
        Adodc2.ConnectionString = strcn.Connection_String
        Adodc2.RecordSource = "select distinct(m_code)from test_info_main where dept_Code='" & deptCode & "'"
        Adodc2.Refresh
      

    If Adodc2.Recordset.RecordCount > 0 Then
         cboMainCode.clear
         
        Adodc2.Recordset.MoveFirst
        While Adodc2.Recordset.EOF = False
         cboMainCode.AddItem Adodc2.Recordset!M_Code
        Adodc2.Recordset.MoveNext
        Wend
        End If
    
End Sub
Private Sub Clear_form()

    txtfield(0).Text = ""
    txtfield(1).Text = ""
    txtfield(2).Text = ""
    txtfield(3).Text = ""
    txtfield(4).Text = "0"
    txtfield(5).Text = "0"
    
End Sub



Private Sub txtfield_GOTFOCUS(Index As Integer)
   Select Case Index
        Case 0, 1, 2, 3, 4, 5, 6
        txtfield(Index).SetFocus
        txtfield(Index).SelStart = 0
        txtfield(Index).SelLength = Len(txtfield(Index))
   End Select
End Sub

Private Sub txtfield_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Not IsNumeric(txtfield(5).Text) Then
       txtfield(5) = ""
       txtfield(2).Text = ""
       txtfield(3).Text = ""
       txtfield(4).Text = ""

Else
'   Select Case Index
'          Case 5:
'            ''PATHOLOGY
'             If cboMainCode.Text = "01" Then
'                txtfield(2).Text = (txtfield(5) * 50) / 100
'                txtfield(3).Text = (txtfield(5) * 25) / 100
'                txtfield(4).Text = (txtfield(5) * 25) / 100
'                '' IMAGING
'              ElseIf cboMainCode.Text = "02" Then
'                 txtfield(2).Text = (txtfield(5) * 70) / 100
'                 txtfield(3).Text = (txtfield(5) * 30) / 100
'                 txtfield(4).Text = (txtfield(5) * 0) / 100
'                 '' CT SCAN
'              ElseIf cboMainCode.Text = "03" Then
'                   If txtfield(5).Text = 3500 Then
'                      txtfield(2).Text = 2000
'                      txtfield(3).Text = 1500
'                      txtfield(4).Text = 0
'                   ElseIf txtfield(5).Text = 5000 Then
'                      txtfield(2).Text = 3000
'                      txtfield(3).Text = 2000
'                      txtfield(4).Text = 0
'
'                   ElseIf txtfield(5).Text = 10000 Then
'                      txtfield(2).Text = 6000
'                      txtfield(3).Text = 4000
'                      txtfield(4).Text = 0
'                  End If
'                '' MRI
'              ElseIf cboMainCode.Text = "04" Then
'                 txtfield(2).Text = (txtfield(5) * 70) / 100
'                 txtfield(3).Text = (txtfield(5) * 30) / 100
'                 txtfield(4).Text = (txtfield(5) * 0) / 100
'              End If
'
'          Case 3:  '''TXTFIELD(3) KEYUP
'               Dim helper As Integer
'               helper = IIf(Len(txtfield(3)) = 0, 0, txtfield(3)) + CInt((txtfield(2)))
'              If txtfield(5).Text - helper >= 0 Then
'                  txtfield(4).Text = txtfield(5).Text - helper
'              Else
'                   ''PATHOLOGY
'                     If cboMainCode.Text = "01" Then
'                           txtfield(2).Text = (txtfield(5) * 50) / 100
'                           txtfield(3).Text = (txtfield(5) * 25) / 100
'                           txtfield(4).Text = (txtfield(5) * 25) / 100
'                      ElseIf cboMainCode.Text = "02" Then
'                           txtfield(2).Text = (txtfield(5) * 70) / 100
'                           txtfield(3).Text = (txtfield(5) * 30) / 100
'                           txtfield(4).Text = (txtfield(5) * 0) / 100
'                      End If
'
'             End If
'
'   End Select
'
End If
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
 
  If KeyAscii = 13 Then
     Select Case Index
            Case 0:
                 txtfield(1).SetFocus
                 txtfield(1).SelStart = 0
                 txtfield(1).SelLength = Len(txtfield(1))
                 Call getTestInformation(cboDeptCode.Text, cboMainCode.Text, txtfield(0).Text)
            Case 1:
                 txtfield(2).SetFocus
                 txtfield(2).SelStart = 0
                 txtfield(2).SelLength = Len(txtfield(2))
     

            Case 2:
                 txtfield(3).SetFocus
                 txtfield(3).SelStart = 0
                 txtfield(3).SelLength = Len(txtfield(3))
           Case 3:
                 txtfield(4).SetFocus
                 txtfield(4).SelStart = 0
                 txtfield(4).SelLength = Len(txtfield(4))
   
           Case 4:
                 txtfield(5).SetFocus
                 txtfield(5).SelStart = 0
                 txtfield(5).SelLength = Len(txtfield(5))
           Case 5:
                 txtfield(6).SetFocus
                 txtfield(6).SelStart = 0
                 txtfield(6).SelLength = Len(txtfield(6))
          Case 6:
                 cmdUpdate.SetFocus
                 
   
    End Select
  End If
End Sub
Private Sub getTestInformation(deptCode As String, mCode As String, testCode As String)
        Adodc2.ConnectionString = strcn.Connection_String
        Adodc2.RecordSource = "select S_NAME,charge_cabin,charge_paying,charge_free,charge_opd,charge_outcase from test_info_sub where dept_code='" & deptCode & "' and m_code='" & mCode & "' and s_code='" & testCode & "' "
        Adodc2.Refresh
      

    If Adodc2.Recordset.RecordCount > 0 Then
        Adodc2.Recordset.MoveFirst
        txtfield(1).Text = Adodc2.Recordset!s_name
        txtfield(2).Text = Adodc2.Recordset!charge_Cabin
        txtfield(3).Text = Adodc2.Recordset!charge_Paying
        txtfield(4).Text = Adodc2.Recordset!charge_Free
        txtfield(5).Text = Adodc2.Recordset!charge_OPD
        txtfield(6).Text = Adodc2.Recordset!charge_outcase
    Else
     txtfield(1).Text = ""
     txtfield(2).Text = 0
     txtfield(3).Text = 0
     txtfield(4).Text = 0
     txtfield(5).Text = 0
    End If
    
  
End Sub


Private Sub txtS_charge_KeyPress(KeyAscii As Integer)
   
End Sub
