VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Setup"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   Icon            =   "CompSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2910
      Top             =   6120
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
      Caption         =   ""
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
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   885
      Left            =   -30
      TabIndex        =   8
      Top             =   -120
      Width           =   8325
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Information &&Financial Year  Setup"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1380
         TabIndex        =   9
         Top             =   270
         Width           =   5670
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00B18A2E&
      Height          =   5415
      Index           =   0
      Left            =   -30
      TabIndex        =   18
      Top             =   630
      Width           =   8325
      Begin VB.Frame Frame3 
         BackColor       =   &H00B18A2E&
         Caption         =   "Fiscal Year"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   1635
         Left            =   0
         TabIndex        =   10
         Top             =   1560
         Width           =   8385
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   4380
            TabIndex        =   1
            Top             =   330
            Width           =   3495
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   690
            TabIndex        =   0
            Top             =   360
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker dtpdate 
            Height          =   285
            Index           =   1
            Left            =   4455
            TabIndex        =   3
            Top             =   1020
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   503
            _Version        =   393216
            Format          =   19791873
            CurrentDate     =   37637
         End
         Begin MSComCtl2.DTPicker dtpdate 
            Height          =   285
            Index           =   0
            Left            =   735
            TabIndex        =   2
            Top             =   1020
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   503
            _Version        =   393216
            Format          =   19791873
            CurrentDate     =   37637
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Left            =   4020
            TabIndex        =   23
            Top             =   360
            Width           =   120
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Title"
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
            Left            =   3480
            TabIndex        =   22
            Top             =   360
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Left            =   540
            TabIndex        =   21
            Top             =   360
            Width           =   60
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Code "
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
            Left            =   90
            TabIndex        =   20
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fiscal Year"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   1
            Left            =   2940
            TabIndex        =   11
            Top             =   990
            Width           =   1335
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H00808080&
            Height          =   375
            Index           =   0
            Left            =   4410
            Top             =   975
            Width           =   2355
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00808080&
            Height          =   375
            Index           =   0
            Left            =   690
            Top             =   975
            Width           =   2085
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ending Date"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   0
            Left            =   4425
            TabIndex        =   12
            Top             =   780
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Starting Date"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   0
            Left            =   675
            TabIndex        =   13
            Top             =   750
            Width           =   930
         End
      End
      Begin VB.TextBox txtCompAddress 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   8.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Index           =   0
         Left            =   900
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   645
         Width           =   6990
      End
      Begin VB.TextBox txtCompName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   0
         Left            =   900
         MaxLength       =   100
         TabIndex        =   15
         Top             =   210
         Width           =   6990
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "CompSetup.frx":030A
         Height          =   2415
         Left            =   -300
         TabIndex        =   19
         Top             =   2970
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         BackColor       =   15456182
         BorderStyle     =   0
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
         RowDividerStyle =   4
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   645
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   375
         TabIndex        =   17
         Top             =   285
         Width           =   420
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3240
      Top             =   3900
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
      Left            =   1140
      Picture         =   "CompSetup.frx":031F
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Delete"
      Top             =   6090
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
      Left            =   630
      Picture         =   "CompSetup.frx":0E59
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "New"
      Top             =   6090
      Width           =   510
   End
   Begin VB.CommandButton cmdEXIT 
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
      Left            =   1650
      Picture         =   "CompSetup.frx":14C3
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exit"
      Top             =   6090
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
      Height          =   420
      Left            =   120
      Picture         =   "CompSetup.frx":1DE1
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Save"
      Top             =   6090
      Width           =   510
   End
   Begin VB.Shape Shape2 
      Height          =   555
      Left            =   90
      Top             =   6000
      Width           =   2115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BorderColor     =   &H00000000&
      Height          =   600
      Left            =   0
      Top             =   6000
      Width           =   8310
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdADD_Click()
    txtCompName(0).Text = ""
    txtCompAddress(0).Text = ""
    dtpdate(0).Value = Date
    dtpdate(1).Value = Date
    txtField(0) = ""
    txtField(1) = ""
    txtField(0).SetFocus
End Sub

Private Sub cmdDELETE_Click()
   If Len(Trim(txtCompName(0).Text)) = 0 Then
       MsgBox "Company name required", vbCritical
       txtCompName(0).SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txtCompAddress(0).Text)) = 0 Then
       MsgBox "Company address required", vbCritical
       txtCompAddress(0).SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txtField(0).Text)) = 0 Then
       MsgBox "Code required", vbCritical
       txtField(0).SetFocus
       Exit Sub
    End If
    
     If Len(Trim(txtField(1).Text)) = 0 Then
       MsgBox "Title required", vbCritical
       txtField(1).SetFocus
       Exit Sub
    End If
    
    
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select * from fiscal_year where code=" & Trim(txtField(0)) & ""
    Adodc1.Refresh
    
    If Adodc1.Recordset.RecordCount = 0 Then
        MsgBox "No such Code exists", vbCritical, "IT Division,DNMIH"
        txtField(0).SetFocus
        Exit Sub
     End If
     If MsgBox("Are you sure to Delete?", vbCritical + vbYesNo, "Deleting..") = vbYes Then
         Call deletefiscalyr
         MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
     End If
    Call FlushCompSetup
     Call load_fiscal
     cmdADD_Click
End Sub
Private Sub deletefiscalyr()
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    Dim Param0 As New Parameter
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    
    Dim userid As String
    userid = "Emdad"
    
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param0 = cmd.CreateParameter("param0", adInteger, adParamInput, 5, 3)
    cmd.Parameters.Append Param0

    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 50, txtField(0).Text)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 300, txtField(1).Text)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, dtpdate(0).Value)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, dtpdate(1).Value)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 10, userid)
    cmd.Parameters.Append Param5
    
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL save_fiscal_year(?,?,?,?,?,?)}"
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

End Sub
Private Sub cmdEXIT_Click()
    Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdSAVE_Click()
    If Len(Trim(txtCompName(0).Text)) = 0 Then
       MsgBox "Company name required", vbCritical
       txtCompName(0).SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txtCompAddress(0).Text)) = 0 Then
       MsgBox "Company address required", vbCritical
       txtCompAddress(0).SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txtField(0).Text)) = 0 Then
       MsgBox "Code required", vbCritical
       txtField(0).SetFocus
       Exit Sub
    End If
    
     If Len(Trim(txtField(1).Text)) = 0 Then
       MsgBox "Title required", vbCritical
       txtField(1).SetFocus
       Exit Sub
    End If
    
    
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select * from fiscal_year where code=" & Trim(txtField(0)) & ""
    Adodc1.Refresh
    
    If Adodc1.Recordset.RecordCount > 0 Then
        MsgBox "Same Code already exists", vbCritical, "IT Division,DNMIH"
        txtField(0).SetFocus
        Exit Sub
     End If
    
    Call SaveCompSetup
    Call savefiscalyr
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
    Call FlushCompSetup
    Call load_fiscal
    cmdADD_Click
End Sub


Private Sub DataGrid1_Click()
  If Adodc2.Recordset.RecordCount > 0 Then
        txtField(0).Text = "" & DataGrid1.Columns(0).Text
        txtField(1).Text = "" & DataGrid1.Columns(1).Text
        dtpdate(0).Value = "" & DataGrid1.Columns(2).Text
        dtpdate(1).Value = "" & DataGrid1.Columns(3).Text
          
  End If
End Sub

Private Sub dtpdate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
        Case 0, 1
            If KeyCode = 13 Then
                SendKeys Chr(9)
            End If
        End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     SendKeys (Chr(9))
  End If
End Sub

Private Sub Form_Load()
    Call FlushCompSetup
    Call load_fiscal
End Sub
Private Sub load_fiscal()
   Adodc2.ConnectionString = strcn.Connection_String
   Adodc2.RecordSource = "select code,comp_setup as title,st_year as start_yr,ed_year as end_yr from fiscal_year"
   Adodc2.Refresh
   
'   If Adodc2.Recordset.RecordCount > 0 Then
   
   DataGrid1.Columns(0).Width = 500
   DataGrid1.Columns(1).Width = 4000
   DataGrid1.Columns(2).Width = 1280
   DataGrid1.Columns(0).Width = 1265
  
   
    
End Sub
Private Sub SaveCompSetup()
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    
    Dim userid As String
    userid = "Emdad"
    
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 50, txtCompName(0).Text)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 300, txtCompAddress(0).Text)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, dtpdate(0).Value)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, dtpdate(1).Value)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 10, userid)
    cmd.Parameters.Append Param5
    
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL SaveCompSetup(?, ?, ?,?,?)}"
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

End Sub
Private Sub savefiscalyr()
  Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    Dim Param0 As New Parameter
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    
    Dim userid As String
    userid = "Emdad"
    
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param0 = cmd.CreateParameter("param0", adInteger, adParamInput, 5, 1)
    cmd.Parameters.Append Param0

    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 50, txtField(0).Text)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 300, txtField(1).Text)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, dtpdate(0).Value)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, dtpdate(1).Value)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 10, userid)
    cmd.Parameters.Append Param5
    
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL save_fiscal_year(?,?,?,?,?,?)}"
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

End Sub
Private Sub FlushCompSetup()
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select * from comp_setup"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        txtCompName(0).Text = Adodc1.Recordset!comp_name
        txtCompAddress(0).Text = Adodc1.Recordset!comp_addr
        dtpdate(0).Value = Adodc1.Recordset!st_dt
        dtpdate(1).Value = Adodc1.Recordset!ed_dt
    End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
        Case 0
           txtField(0).BackColor = &H80000018
         Case 1
            txtField(1).BackColor = &H80000018
   End Select
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   Select Case Index
        Case 0
           txtField(0).BackColor = vbWhite
         Case 1
            txtField(1).BackColor = vbWhite
   End Select
End Sub
