VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opening Balance"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   FillColor       =   &H009DD1EE&
   Icon            =   "OPEN_BAL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   825
      Left            =   -30
      TabIndex        =   22
      Top             =   -120
      Width           =   10365
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Balance Entry"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   6930
         TabIndex        =   23
         Top             =   270
         Width           =   2850
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1410
      Left            =   9810
      TabIndex        =   12
      Top             =   2295
      Visible         =   0   'False
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   2487
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14737632
      ForeColor       =   -2147483646
      BackColorFixed  =   14737632
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483635
      BackColorBkg    =   16777215
      FocusRect       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cboUserAcc 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1890
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Top             =   1050
      Width           =   1650
   End
   Begin VB.TextBox txtUnitCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4050
      TabIndex        =   21
      Top             =   4440
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4860
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox txtQuaryPart 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5760
      TabIndex        =   19
      Top             =   4440
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox nbtTotCr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8370
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   3960
      Width           =   1245
   End
   Begin VB.TextBox nbtTotDr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7110
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   3960
      Width           =   1245
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2970
      Top             =   4440
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
      Bindings        =   "OPEN_BAL.frx":030A
      Height          =   2580
      Left            =   225
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1380
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4551
      _Version        =   393216
      AllowUpdate     =   0   'False
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      ColumnCount     =   5
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3569.953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
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
      Height          =   420
      Left            =   2175
      Picture         =   "OPEN_BAL.frx":031F
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Exit"
      Top             =   4320
      Width           =   510
   End
   Begin MSComCtl2.DTPicker dtvou_dt 
      Height          =   330
      Left            =   540
      TabIndex        =   0
      Top             =   1050
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Format          =   58392577
      CurrentDate     =   37200
   End
   Begin VB.TextBox txtACC_NAME 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3555
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1050
      Width           =   3585
   End
   Begin VB.TextBox nbrOPEN_DR 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7155
      MaxLength       =   12
      TabIndex        =   3
      Top             =   1050
      Width           =   1245
   End
   Begin VB.TextBox nbrOPEN_CR 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8415
      MaxLength       =   12
      TabIndex        =   4
      Top             =   1050
      Width           =   1245
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
      Height          =   420
      Left            =   135
      Picture         =   "OPEN_BAL.frx":0C3D
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Save"
      Top             =   4320
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
      Height          =   420
      Left            =   1665
      Picture         =   "OPEN_BAL.frx":12A7
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Preview"
      Top             =   4320
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
      Height          =   420
      Left            =   645
      Picture         =   "OPEN_BAL.frx":1911
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "New"
      Top             =   4320
      Width           =   510
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
      Height          =   420
      Left            =   1155
      Picture         =   "OPEN_BAL.frx":1F7B
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Delete"
      Top             =   4320
      Width           =   510
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3090
      Top             =   4380
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3000
      Top             =   4350
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   2970
      Top             =   4290
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
   Begin VB.Shape Shape2 
      Height          =   525
      Left            =   90
      Top             =   4260
      Width           =   2625
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Index           =   1
      Left            =   1890
      TabIndex        =   18
      Top             =   780
      Width           =   375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   630
      TabIndex        =   17
      Top             =   780
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Debit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   7965
      TabIndex        =   16
      Top             =   780
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Credit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   9135
      TabIndex        =   15
      Top             =   780
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Index           =   0
      Left            =   3555
      TabIndex        =   14
      Top             =   780
      Width           =   705
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H80000001&
      FillStyle       =   0  'Solid
      Height          =   765
      Index           =   2
      Left            =   -255
      Top             =   4170
      Width           =   10215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboUnit_Click()
    Call GetGrdData
End Sub

Private Sub cboUserAcc_Click()
    Call GetAccName(Me, Trim(Me.cboUserAcc.Text))
End Sub
Private Sub FlushOpenBal()
    Adodc4.ConnectionString = strcn.Connection_String
    Adodc4.RecordSource = "select dr_amt,cr_amt from ledger where acc_code in (select acc_code from acct where user_acc ='" & Trim(Me.cboUserAcc.Text) & "')"
    Adodc4.Refresh
    If Adodc4.Recordset.RecordCount > 0 Then
      Me.nbrOPEN_DR = Adodc4.Recordset!dr_amt
      Me.nbrOPEN_CR = Adodc4.Recordset!cr_amt

    End If
End Sub
Private Sub cmdADD_Click()
    ClearScreen
    cboUserAcc.Text = ""
    cboUserAcc.SetFocus
End Sub

Private Sub cmdDELETE_Click()
    If Len(Trim(cboUserAcc.Text)) = 0 Then Exit Sub
    
    On Error GoTo err_loop
    
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    
    
    Dim userid, mode As String
    userid = "Emdad"
    mode = "2"
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mode)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 40, cboUserAcc.Text)
    cmd.Parameters.Append Param2

    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, dtvou_dt.Value)
    cmd.Parameters.Append Param3
       
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 9, Val(nbrOPEN_DR.Text))
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 9, Val(nbrOPEN_CR.Text))
    cmd.Parameters.Append Param5
        
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 50, userid)
    cmd.Parameters.Append Param6
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL PostOpnBal(?, ?, ?, ?, ?, ?)}"
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

    '''''-----------------------------------------------------------------------------
        Adodc3.ConnectionString = strcn.Connection_String
        Adodc3.RecordSource = "select msg from message"
        Adodc3.Refresh
        If Adodc3.Recordset.RecordCount > 0 Then
            MsgBox Adodc3.Recordset!msg, vbOKOnly + vbInformation, "Save..."
        End If
    ''''------------------------------------------------------------------
    
    
    Call cmdADD_Click
    Call GetGrdData
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

Private Sub cmdEXIT_Click()
    Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdPREVIEW_Click()
'    rptMode = 4
'    Me.txtQuaryPart.Text = "and a.vou_type=''OP''"
'
'    Me.txtTitle.Text = "Opening Trial Balance"
'    Form17.Show vbModal
End Sub

Private Sub cmdSAVE_Click()
    On Error GoTo err_loop
    If Len(Trim(cboUserAcc.Text)) = 0 Then
       MsgBox "Accounts code required", vbInformation, "IT Division,DNMIH"
       cboUserAcc.SetFocus
       Exit Sub
    End If


    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    
    
    Dim userid, mode As String
    userid = "Emdad"
    mode = "1"
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mode)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 40, cboUserAcc.Text)
    cmd.Parameters.Append Param2

    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, dtvou_dt.Value)
    cmd.Parameters.Append Param3
       
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 9, Val(nbrOPEN_DR.Text))
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 9, Val(nbrOPEN_CR.Text))
    cmd.Parameters.Append Param5
        
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 50, userid)
    cmd.Parameters.Append Param6
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL PostOpnBal(?, ?, ?, ?, ?, ?)}"
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

    '''''-----------------------------------------------------------------------------
        Adodc3.ConnectionString = strcn.Connection_String
        Adodc3.RecordSource = "select msg from message"
        Adodc3.Refresh
        If Adodc3.Recordset.RecordCount > 0 Then
            MsgBox Adodc3.Recordset!msg, vbOKOnly + vbInformation, "Save..."
        End If
    ''''------------------------------------------------------------------

    Call cmdADD_Click
    Call GetGrdData
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

Private Sub DataGrid1_Click()
    On Error GoTo err_desc
    Me.dtvou_dt.Value = Me.DataGrid1.Columns(0).Text
    Me.cboUserAcc.Text = Me.DataGrid1.Columns(1).Text
    Me.txtacc_name.Text = Me.DataGrid1.Columns(2).Text
    Me.nbrOPEN_DR.Text = Me.DataGrid1.Columns(3).Text
    Me.nbrOPEN_CR.Text = Me.DataGrid1.Columns(4).Text
    Exit Sub
err_desc:
        MsgBox Err.Description, vbCritical, "IT Division,DNMIH"
End Sub

Private Sub dtvou_dt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub Form_Load()
    dtvou_dt.Value = Date
'    Call GetUserAcc(Me)
    Call GetGrdData
End Sub

Private Sub MSFlexGrid1_DblClick()
    
    If Len(Trim(MSFlexGrid1.Text)) <> 0 Then
       cboUserAcc.Text = MSFlexGrid1.Text
       Call cboUserAcc_LostFocus
       nbrOPEN_DR.SetFocus
       MSFlexGrid1.Visible = False
    Else
       cboUserAcc.SetFocus
    End If
    MSFlexGrid1.Visible = False
End Sub

Private Sub MSFlexGrid1_LostFocus()
    Call MSFlexGrid1_DblClick
End Sub

Private Sub nbrOPEN_CR_Change()
    If Len(Trim(nbrOPEN_CR.Text)) = 0 Then Exit Sub

    If IsNumeric(nbrOPEN_CR.Text) Then
       If Val(nbrOPEN_DR.Text) > 0 Then
          nbrOPEN_CR.Text = ".00"
       End If
       If Val(nbrOPEN_CR.Text) > 0 Then
          nbrOPEN_DR.Text = ".00"
       End If
       Exit Sub
    Else
       MsgBox "Accept numeric value only", vbInformation
       nbrOPEN_CR.Text = ".00"
       nbrOPEN_CR.SetFocus
       Exit Sub
    End If
End Sub

Private Sub nbrOPEN_CR_GotFocus()
    nbrOPEN_CR.SelLength = Len(nbrOPEN_CR.Text)
End Sub

Private Sub nbrOPEN_DR_Change()
    If Len(nbrOPEN_DR.Text) = 0 Then Exit Sub

    If IsNumeric(nbrOPEN_DR.Text) Then
       If Val(nbrOPEN_DR.Text) > 0 Then
          nbrOPEN_CR.Text = ".00"
       End If
       Exit Sub
    Else
       MsgBox "Accept numeric value only", vbInformation
       nbrOPEN_DR.Text = ".00"
       nbrOPEN_DR.SetFocus
       Exit Sub
    End If
End Sub

Private Sub nbrOPEN_DR_GotFocus()
    nbrOPEN_DR.SelLength = Len(nbrOPEN_DR.Text)
End Sub

Private Sub cboUserAcc_LostFocus()
    If Len(Trim(cboUserAcc.Text)) = 0 Then Exit Sub
     Dim Conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset

    Conn.Open strcn.Connection_String
    Set cmd.ActiveConnection = Conn
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "select acc_name from acct where user_acc='" & Trim(cboUserAcc.Text) & "'"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
    
    If Not (RS.EOF Or RS.BOF) Then
      txtacc_name.Text = RS!acc_name
    Else
        Call getAcc_Code
    End If
    Call FlushOpenBal
'    Exit Sub
'err_loop:
'    MsgBox Err.Description, vbCritical
'    Resume Next
End Sub

Private Sub ClearScreen()
    txtacc_name.Text = ""
    nbrOPEN_DR.Text = "0.00"
    nbrOPEN_CR.Text = "0.00"
End Sub

Private Sub getAcc_Code()
    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 0
    
    MSFlexGrid1.Left = cboUserAcc.Left
    MSFlexGrid1.Top = cboUserAcc.Top
    
    MSFlexGrid1.ColWidth(0) = "1200"
    MSFlexGrid1.ColAlignment(0) = 1
    
    MSFlexGrid1.ColWidth(1) = "5800"
    
    On Error GoTo err_loop
    
    Adodc2.ConnectionString = strcn.Connection_String
    Adodc2.RecordSource = "select user_acc,acc_name from acct where upper(acc_name) like '" & _
            UCase(Trim(cboUserAcc.Text)) & "%' and acc_code not in(select acc_head from acct)"
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then
       Do Until Adodc2.Recordset.EOF
          MSFlexGrid1.AddItem Adodc2.Recordset!user_acc & vbTab & Adodc2.Recordset!acc_name
          Adodc2.Recordset.MoveNext
       Loop

    End If
    

    MSFlexGrid1.Visible = True
    MSFlexGrid1.SetFocus
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub
Private Sub GetGrdData()
'    On Error GoTo err_loop
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select vou_date,(select user_acc from acct where acct.acc_code=ledger.acc_code) as Code,(select acc_name from acct where acct.acc_code=ledger.acc_code) as Accounts,dr_amt as Debit,cr_amt as Credit from ledger where ledger.vou_type='op' and ledger.acc_code=(select max(l.acc_code) from ledger l where l.vou_no=ledger.vou_no)"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        Dim TotOpenDr, TotOpenCr As Double
        Adodc1.Recordset.MoveFirst
        While Adodc1.Recordset.EOF = False
            TotOpenDr = TotOpenDr + Adodc1.Recordset!debit
            TotOpenCr = TotOpenCr + Adodc1.Recordset!credit
            Adodc1.Recordset.MoveNext
        Wend
        
        nbtTotDr.Text = TotOpenDr
        nbtTotCr.Text = TotOpenCr
    End If
    DataGrid1.Columns(0).Width = 1305.071
    DataGrid1.Columns(1).Width = 1679.811

    DataGrid1.Columns(2).Width = 3614.74

    DataGrid1.Columns(3).Width = 1289.764
    DataGrid1.Columns(3).Alignment = dbgRight

    DataGrid1.Columns(4).Width = 1200
    DataGrid1.Columns(4).Alignment = dbgRight


''=====================================================
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

