VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCCU_BED_Main 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   7590
      Width           =   11385
      Begin VB.Label Label24 
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
         TabIndex        =   26
         ToolTipText     =   "Developed && Maintenanced by:A.N.M. Atiqur Rahman,Software Programmer, IT Division, DNMIH"
         Top             =   90
         Width           =   2715
      End
      Begin VB.Label Label19 
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
         TabIndex        =   25
         Top             =   60
         Width           =   4725
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6285
      Left            =   0
      TabIndex        =   5
      Top             =   750
      Width           =   9855
      Begin VB.TextBox txtDepartment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   7020
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1230
         Width           =   2655
      End
      Begin VB.TextBox txtCurrentBed 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1950
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1230
         Width           =   2655
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2115
         Left            =   90
         TabIndex        =   27
         Top             =   3450
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   3731
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         ForeColor       =   8421631
         HeadLines       =   1
         RowHeight       =   19
         AllowDelete     =   -1  'True
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
               LCID            =   2057
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
               LCID            =   2057
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
      Begin VB.Frame Frame6 
         Height          =   165
         Left            =   0
         TabIndex        =   28
         Top             =   3240
         Width           =   9855
      End
      Begin VB.Frame Frame4 
         Height          =   165
         Left            =   30
         TabIndex        =   23
         Top             =   2310
         Width           =   9855
      End
      Begin VB.Frame Frame1 
         Height          =   165
         Left            =   30
         TabIndex        =   22
         Top             =   1500
         Width           =   9825
      End
      Begin VB.TextBox txtAdvance 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2010
         TabIndex        =   17
         Top             =   2730
         Width           =   1245
      End
      Begin VB.ComboBox CBOYRCODE 
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
         ItemData        =   "frmCCU_BED_Main.frx":0000
         Left            =   7380
         List            =   "frmCCU_BED_Main.frx":000A
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "YR-0708"
         Top             =   330
         Width           =   2055
      End
      Begin VB.TextBox TxtAgeRelease 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7380
         Locked          =   -1  'True
         MaxLength       =   17
         TabIndex        =   9
         Top             =   780
         Width           =   735
      End
      Begin VB.TextBox txtNameRelease 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1950
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   810
         Width           =   4275
      End
      Begin VB.TextBox txtCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4830
         MaxLength       =   17
         TabIndex        =   7
         Text            =   "600"
         Top             =   2820
         Width           =   555
      End
      Begin VB.TextBox txtReg_no_extra 
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
         Height          =   315
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   300
         Width           =   2235
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   8130
         Top             =   5670
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
         Left            =   4140
         Top             =   5640
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   345
         Left            =   1980
         TabIndex        =   20
         Top             =   1830
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   16711680
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mm-YYYY"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   345
         Left            =   7380
         TabIndex        =   21
         Top             =   1830
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   192
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mm-YYYY"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTotal 
         Caption         =   "Label9"
         Height          =   255
         Left            =   1530
         TabIndex        =   33
         Top             =   5910
         Width           =   1425
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Dept :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5280
         TabIndex        =   32
         Top             =   1200
         Width           =   1530
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Bed :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   210
         TabIndex        =   30
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "CCU.Rel. Date :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5790
         TabIndex        =   19
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CCU Adm. Date :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   18
         Top             =   1830
         Width           =   1725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Advance :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   16
         Top             =   2820
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FISCAL YEAR :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   0
         Left            =   5895
         TabIndex        =   15
         Top             =   390
         Width           =   1320
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reg. No :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   705
         TabIndex        =   13
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6660
         TabIndex        =   12
         Top             =   780
         Width           =   555
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   900
         TabIndex        =   11
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Charge"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4800
         TabIndex        =   10
         Top             =   2580
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   7140
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   7350
      TabIndex        =   3
      Top             =   7140
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   8580
      TabIndex        =   2
      Top             =   7140
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   9825
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CCU BED REGISTRATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   315
         Left            =   2670
         TabIndex        =   1
         Top             =   210
         Width           =   4755
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -510
         Picture         =   "frmCCU_BED_Main.frx":0020
         Top             =   -30
         Width           =   11820
      End
   End
   Begin VB.Shape Shape3 
      Height          =   465
      Left            =   6030
      Top             =   7080
      Width           =   3825
   End
End
Attribute VB_Name = "frmCCU_BED_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Con As New MyConnection
Dim UTILITY As New clsUtility
Dim Conn2 As New Connection
Dim rs2 As New Recordset
Dim cmd As New Command
Private Sub cmdExit_Click()
    Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If

End Sub
Private Sub setParamAndDate()
   If MaskEdBox1.Enabled = True Then
        paramMode = 1
        paramDate = MaskEdBox1.Text
     Else
       paramMode = 2
       paramDate = MaskEdBox2.Text
  End If
End Sub
Private Sub cmdSave_Click()
If paramMode = 1 Then
    If MaskEdBox1.Text = "__/__/__" Then
          MsgBox "Please put a valid date here", vbInformation, "DNMIH"
          MaskEdBox1.SetFocus
     End If
ElseIf paramMode = 2 Then
    If MaskEdBox2.Text = "__/__/__" Then
          MsgBox "Please put a valid date here", vbInformation, "DNMIH"
          MaskEdBox2.SetFocus
     End If
End If
If txtAdvance = "" Then
        MsgBox "Advance Required"
        txtAdvance.SetFocus
        Exit Sub
End If

 If UTILITY.User_Shift_validation(frmMAIN.lbluser_id, frmMAIN.lblUserType.Caption) = False Then
    MsgBox "Mr. " & frmMAIN.lblName & "  Your Shift has been Expired.. " & vbCrLf & " " & vbCrLf & "Please Contact With Administrator", vbInformation, " IT, DNMIH."
    Exit Sub
End If
            


setParamAndDate
Call saveCCUBedInfo
MsgBox "Operation successful", vbInformation + vbOKOnly, "Save..."
Call GetCCUBedInfo(txtReg_no_extra, CBOYRCODE)
dateEnabledDisabled
End Sub
Private Sub FormatGrid()
  
'   DataGrid1.Columns.Item(0) = "Start Date"
'   DataGrid1.Columns.Item(1) = "End Date"
'   DataGrid1.Columns.Item(2) = "Amount"
'
End Sub
Private Sub saveCCUBedInfo()
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
If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
End If
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, paramMode)
    cmd.Parameters.Append Param1 'mode
    
    Set Param2 = cmd.CreateParameter("param2", adDouble, adParamInput, 5, txtReg_no_extra.Text)
    cmd.Parameters.Append Param2 'in_reg_no
    
   
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, Trim(txtCharge.Text))
    cmd.Parameters.Append Param3 'Bed_charge
    
    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 12, paramDate)
    cmd.Parameters.Append Param4 'START OR END DATE
 
   Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 7, CBOYRCODE.Text)
   cmd.Parameters.Append Param5 'fiscal year
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 2, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param6 'U_id default Sumon

      
   Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 5, frmMAIN.lblBooth)
    cmd.Parameters.Append Param7 'booth
    
    
       cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Save_ccu_Bed_info_indoor(?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set rs2 = cmd.Execute
    

  cmd.Properties("PLSQLRSet") = False
   If Conn.State = 1 Then
       Conn.Close
       Set Conn = Nothing
    End If
End Sub
Private Sub dateEnabledDisabled()
 If paramMode = 1 Then
       MaskEdBox1.Enabled = True
       MaskEdBox2.Enabled = False
    Else
      MaskEdBox1.Enabled = False
      MaskEdBox2.Enabled = True
    End If
End Sub
Private Sub Form_Load()
    FormatGrid
    txtReg_no_extra = frmReadvancepayment.txtReg_noInTest
    CBOYRCODE.Text = frmReadvancepayment.CBOYRCODE.Text
    MaskEdBox1.Text = Format(Date, "DD/MM/YY")
    MaskEdBox2.Text = Format(Date, "DD/MM/YY")
    txtCurrentBed = UTILITY.GetCurrentBed(txtReg_no_extra, CBOYRCODE)
    txtDepartment = UTILITY.GetCurrentDepartment(txtReg_no_extra, CBOYRCODE)
    
    Call GetCCUBedInfo(txtReg_no_extra, CBOYRCODE)
    
    dateEnabledDisabled
    
    If Conn2.State = 0 Then
       Conn2.ConnectionString = strcn.Connection_String
       Conn2.Open
    End If
        cmd.ActiveConnection = Conn2
        cmd.CommandType = adCmdText
        cmd.CommandText = "select pat_name,pat_guard_name,sex,age,doc_dept  From in_door_pat_info_main Where in_reg_no ='" & Trim(txtReg_no_extra.Text) & "' AND YRCODE='" & Trim(CBOYRCODE.Text) & "'"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        rs2.CursorLocation = adUseClient

        rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
        If rs2.RecordCount > 0 Then
           txtNameRelease = rs2!pat_name
           TxtAgeRelease = rs2!age
           cmd.Properties("iRowsetChange") = False
    If Conn2.State = 1 Then
         Conn2.Close
    End If
    lblTotal = totalCCUCharge
Else
 MsgBox "Invalid Registration No", vbInformation, "Warning: IT, DNMIH"
 'rs2.Close
 
 
  
If Conn2.State = 1 Then
    Conn2.Close
End If
 Exit Sub
 Unload Me

End If

End Sub
Public Function GetCCUBedInfo(registrationNo As String, fiscalYear As String) As ADODB.Recordset
   Set DataGrid1.DataSource = UTILITY.GetCCUBedInfo(registrationNo, fiscalYear)
   
End Function

