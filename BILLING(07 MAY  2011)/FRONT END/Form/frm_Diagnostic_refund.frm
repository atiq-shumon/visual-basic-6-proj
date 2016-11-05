VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_Diagnostic_refund 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4230
   ClientLeft      =   -105
   ClientTop       =   390
   ClientWidth     =   9315
   FillColor       =   &H007DABD0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTPRINT_diag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc3"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5130
      TabIndex        =   15
      Top             =   3450
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      TabIndex        =   17
      Top             =   3510
      Width           =   9315
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "CLOSE"
         Height          =   375
         Left            =   8040
         TabIndex        =   18
         ToolTipText     =   "PRESS TO CLOSE"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "PREVIEW"
         Height          =   375
         Left            =   6810
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "PRINT"
         Height          =   375
         Left            =   5580
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CMDSAVE 
         Caption         =   "SAVE"
         Height          =   375
         Left            =   4350
         TabIndex        =   3
         ToolTipText     =   "PRESS TO SAVE"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         Height          =   465
         Left            =   4290
         Top             =   180
         Width           =   4995
      End
      Begin VB.Image Image2 
         Height          =   855
         Left            =   0
         Picture         =   "frm_Diagnostic_refund.frx":0000
         Top             =   -60
         Width           =   11820
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1365
      Left            =   -30
      TabIndex        =   10
      Top             =   2160
      Width           =   9405
      Begin VB.TextBox Text1 
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
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   390
         TabIndex        =   1
         Top             =   540
         Width           =   6645
      End
      Begin VB.TextBox TxtPreviousPayment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   7080
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   540
         Width           =   1875
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   7470
         TabIndex        =   14
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Numbers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   390
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      TabIndex        =   9
      Top             =   -120
      Width           =   9315
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DIAGNOSTIC REFUND ENTRY"
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
         Left            =   2160
         TabIndex        =   16
         Top             =   270
         Width           =   4755
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -1260
         Picture         =   "frm_Diagnostic_refund.frx":5982
         Top             =   30
         Width           =   11820
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   9405
      Begin VB.TextBox txtAddrInTest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   330
         TabIndex        =   0
         Top             =   1080
         Width           =   8625
      End
      Begin MSComCtl2.DTPicker Dt_date 
         Height          =   330
         Left            =   6990
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   450
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   49152
         Format          =   62062593
         CurrentDate     =   37114
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   6360
         TabIndex        =   11
         Top             =   465
         Width           =   510
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   300
         TabIndex        =   8
         Top             =   795
         Width           =   480
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   240
      Top             =   0
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
      Left            =   360
      Top             =   0
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
      Left            =   120
      Top             =   0
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   3570
      Top             =   3600
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   5040
      Top             =   3660
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   360
      Top             =   0
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   120
      Top             =   0
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
      Height          =   525
      Left            =   30
      Top             =   3570
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   -6660
      TabIndex        =   6
      Top             =   4770
      Width           =   270
   End
End
Attribute VB_Name = "frm_Diagnostic_refund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Con As New MyConnection
Dim Conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset
Dim RS1 As New Recordset
Dim Conn2 As New Connection
Dim rs2 As New Recordset
Public strUid As String
Dim VoucherNumber
Dim UTILITY As New clsUtility
Dim OptionBtnvalue As Integer
Public strcn        As New MyConnection
Private Sub Get_Voucher_Number()
On Error GoTo Errdesc
Dim conn10 As New Connection
Dim rs10 As New Recordset
Dim cmd As New Command
If conn10.State = 0 Then
conn10.ConnectionString = strcn.Connection_String
conn10.Open
End If
VoucherNumber = 0
cmd.ActiveConnection = conn10
cmd.CommandType = adCmdText
cmd.CommandText = "select max(acct.vou.vou_no)+1 from acct.vou where upper(acct.vou.vou_type)=upper('cr')"
rs10.CursorLocation = adUseClient
rs10.Open cmd.CommandText, conn10, adOpenDynamic, adLockOptimistic
    If rs10.RecordCount > 0 Then
        If IsNull(rs10.Fields(0)) Then
            VoucherNumber = 1
        Else
            VoucherNumber = rs10.Fields(0)
       End If
    Else
        VoucherNumber = 1
    End If
Exit Sub
If conn10.State = 1 Then
    conn10.Close
    Set conn10 = Nothing
End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, " IT, DNMIH"
End Sub

'Private Sub Check1_Click()
'  '''Check1.Value = 1
' TxtPreviousPayment.Top = 780
'   If Check1.Value = 0 Then
'      Label4.Enabled = False
'
'    End If
'    If Check1.Value = 1 Then
'      TxtPreviousPayment.Top = 780
'      OptionBtnvalue = 1
'      Label4.Enabled = True
'       Check2.Value = 0
'      Check3.Value = 0
'    End If
'End Sub

'Private Sub Check2_Click()
'  TxtPreviousPayment.Top = 1110
'  If Check2.Value = 1 Then
'    Label6.Enabled = True
'     TxtPreviousPayment.Top = 1110
'    OptionBtnvalue = 2
'  Else
'    Label6.Enabled = False
'  End If
'  Check1.Value = 0
'  Check3.Value = 0
'End Sub

'Private Sub Check3_Click()
'TxtPreviousPayment.Top = 1560
'   If Check3.Value = 1 Then
'      OptionBtnvalue = 3
'      Label1.Visible = True
'      Text1.Visible = True
'      Label12.Enabled = True
'      TxtPreviousPayment.Top = 1560
'      Text1.SetFocus
'
'   Else
'      Label1.Visible = False
'      Text1.Visible = False
'      Label12.Enabled = False
'  End If
'  Check1.Value = 0
'  Check2.Value = 0
'End Sub

Private Sub CMDEXIT_Click()
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdPrint_Click()
If TXTPRINT_diag.Visible = False Then
       TXTPRINT_diag.Visible = True
   End If
  TXTPRINT_diag.ForeColor = vbBlue
'  PREVIEW_VAR = Val(TXTPRINT_diag)
  
   If TXTPRINT_diag = "" Then
     TXTPRINT_diag.SetFocus
     Exit Sub
   Else
'    PREVIEW_VAR = Val(TXTPRINT_diag)
'        rptMode = 6
'      Viewer.Show vbModal
    
      TXTPRINT_diag.Visible = False
End If


If TXTPRINT_diag.Text <> "" Then
        print_diag_refund
End If
 TXTPRINT_diag = ""

End Sub
Private Sub cmdSave_Click()
If UTILITY.User_Shift_validation(frmMAIN.lbluser_id, frmMAIN.lblUserType.Caption) = False Then
    MsgBox "Dear. " & frmMAIN.lblName & "  Your Shift has been Expired.. " & vbCrLf & " " & vbCrLf & "Please Contact With Administrator", vbInformation, " IT, DNMIH."
    Exit Sub
End If


Call save_diag_refund

MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."

CMDEXIT.SetFocus
     
print_diag_refund

txtAddrInTest = ""
TxtPreviousPayment = ""

CMDEXIT.SetFocus
If Conn.State = 1 Then
    Conn.Close
    Set Conn = Nothing
    Set RS = Nothing
    Set cmd = Nothing
End If
End Sub
Private Sub print_diag_refund()
    rptMode = 43
    Viewer.Show vbModal
End Sub
 Private Sub save_diag_refund()
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
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, txtAddrInTest)
    cmd.Parameters.Append Param1 'IN_REG_NO
     
     Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 30, Text1)
    cmd.Parameters.Append Param2 'IN_REG_NO
     TxtPreviousPayment = -Val(TxtPreviousPayment) ''''negation of refund
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, TxtPreviousPayment)
    cmd.Parameters.Append Param3 'readvance
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 30, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param4 'U_id default SHumon
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 30, frmMAIN.lblBooth)
    cmd.Parameters.Append Param5 'booth
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 30, 3)
    cmd.Parameters.Append Param6 'booth
    
   
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL save_diag_refund(?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
       Conn.Close
    End If
    
    Adodc4.ConnectionString = strcn.Connection_String
    Adodc4.RecordSource = "SELECT MAX(REC_NO) AS REC_NO FROM RECEIPT_NO_COUNTER"
    Adodc4.Refresh
    If Adodc4.Recordset.RecordCount > 0 Then
        TXTPRINT_diag = Adodc4.Recordset!REC_NO
    End If
    
    
    
    
End Sub
'Private Sub DataGrid1_Click()
'
'If DataGrid1.Row > 0 Then
'End If
'
'
'End Sub



Private Sub Dt_date_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     txtAddrInTest.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
        Unload Me
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    SendKeys Chr(9)
'End If
End Sub
Private Sub Form_Load()
        Dim temp
        Dt_date = Date
   End Sub
Private Sub total_adv()

                 Adodc5.ConnectionString = strcn.Connection_String
                    Adodc5.RecordSource = "select  nvl(sum(advance),0)as advance from advance where in_reg_no ='" & Trim(frmReadvancepayment.txtReg_noInTest.Text) & "'"
                    Adodc5.Refresh
                If Adodc5.Recordset.RecordCount > 0 Then
                     TxtPreviousPayment = Adodc5.Recordset!advance

                End If
       Adodc5.Refresh

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     TxtPreviousPayment.SetFocus
  End If
End Sub

Private Sub txtAddrInTest_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Text1.SetFocus
  End If
End Sub

'Private Sub txtCurpayment_Change()
'If Not IsNumeric(txtCurpayment) Then
'     txtCurpayment = ""
'Else
'   If Val(txtCurpayment) < 0 Then
'      MsgBox "Advance Can't be minus Fiegure", vbInformation, " IT, DNMIH."
'Else
'     txtTotalPayment = Val(txtCurpayment) + Val(TxtPreviousPayment)
'  End If
'End If
'
'End Sub

Private Sub TxtPreviousPayment_Change()
  If Not IsNumeric(TxtPreviousPayment) Then
     TxtPreviousPayment = ""
  End If
End Sub

Private Sub TxtPreviousPayment_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     cmdSAVE.SetFocus
  End If
End Sub
