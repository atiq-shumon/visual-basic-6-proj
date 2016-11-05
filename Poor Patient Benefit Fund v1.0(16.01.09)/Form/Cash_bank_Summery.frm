VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form19 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report :Cash/Bank Book Summary"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   Icon            =   "Cash_bank_Summery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   885
      Left            =   -30
      TabIndex        =   2
      Top             =   -150
      Width           =   6195
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash/Bank Book at a glance"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   420
         Left            =   330
         TabIndex        =   3
         Top             =   210
         Width           =   5460
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1395
      Left            =   -60
      TabIndex        =   7
      Top             =   600
      Width           =   6195
      Begin VB.ComboBox cboStUserCode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   765
         Sorted          =   -1  'True
         TabIndex        =   9
         Text            =   "cboStUserCode"
         Top             =   660
         Width           =   1560
      End
      Begin VB.ComboBox cboStAccName 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2340
         Sorted          =   -1  'True
         TabIndex        =   8
         Text            =   "cboStAccName"
         Top             =   660
         Width           =   3345
      End
      Begin MSComCtl2.DTPicker dtst_dt 
         Height          =   315
         Left            =   765
         TabIndex        =   10
         Top             =   330
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22675457
         CurrentDate     =   36949
      End
      Begin MSComCtl2.DTPicker dted_dt 
         Height          =   315
         Left            =   4125
         TabIndex        =   11
         Top             =   330
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22675457
         CurrentDate     =   36949
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3825
         TabIndex        =   14
         Top             =   345
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   300
         TabIndex        =   13
         Top             =   375
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   315
         TabIndex        =   12
         Top             =   675
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   915
      Left            =   -60
      TabIndex        =   4
      Top             =   1860
      Width           =   6255
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5490
         Picture         =   "Cash_bank_Summery.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exit"
         Top             =   270
         Width           =   510
      End
      Begin VB.CommandButton cmdPREVIEW 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4950
         Picture         =   "Cash_bank_Summery.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Preview"
         Top             =   285
         Width           =   510
      End
      Begin VB.Shape Shape1 
         Height          =   540
         Index           =   4
         Left            =   4890
         Top             =   210
         Width           =   1170
      End
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   420
      Top             =   1890
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
      Caption         =   "Adodc4"
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
   Begin VB.TextBox txtOpenBal 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   1920
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
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
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3285
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   990
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   210
      Top             =   1920
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
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
      Left            =   300
      Top             =   1890
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   150
      Top             =   1950
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
      Caption         =   "Adodc4"
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
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboEdAccName_LostFocus()
    
End Sub

Private Sub cboStAccName_Click()
'    cboStUserCode.Text = GetUserCode(cboStAccName.Text)
Adodc5.ConnectionString = strcn.Connection_String
Adodc5.RecordSource = "select acc_code,acc_name from acct where acc_name='" & Trim(cboStAccName.Text) & "'"
Adodc5.Refresh
If Adodc5.Recordset.RecordCount > 0 Then
    cboStUserCode = Adodc5.Recordset!acc_code
End If
End Sub

Private Sub cboStAccName_LostFocus()
    If Len(Trim(cboStAccName.Text)) = 0 Then Exit Sub
    cboStUserCode.Text = GetUserCode(cboStAccName.Text)
End Sub

Private Sub cboStUserCode_Click()
    cboStAccName.Text = GetAccName(cboStUserCode.Text)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPREVIEW_Click()
 On Error GoTo err_loop
 txtOpenBal.Text = 0
 Dim dr_amt, cr_amt As Double
'    If Len(Trim(Me.cboStUserCode.Text)) = 0 Then
'       MsgBox "Account code required", vbCritical
'       Me.cboStUserCode.SetFocus
'       Exit Sub
'    End If
     If dtst_dt.Value > dted_dt.Value Then
       MsgBox "Improper date range ", vbCritical
       Exit Sub
    End If
    ''===================for opening balance====================
    Screen.MousePointer = vbHourglass
    Adodc4.ConnectionString = strcn.Connection_String
    ''Adodc4.RecordSource = "select sum(dr_amt) as debit,sum(cr_amt) as credit from ledger where  acc_code in(select acc_code from acct where user_acc='" & Trim(cboStUserCode.Text) & "' and  to_date(vou_date,'dd-mon-yyyy') < '" & dtst_dt.Value& )"'"
    'if  to_char(vou_date,'yyyy') < '" & Format(dtst_dt.Value, "yyyy") & "')" then
    
   
    'Adodc4.RecordSource = "select sum(dr_amt) as debit,sum(cr_amt) as credit from ledger where  acc_code in(select acc_code from acct where user_acc='" & Trim(cboStUserCode.Text) & "') and(to_char((vou_date),'dd-mon-yyyy')<('" & Format(dtst_dt.Value, "dd-mmm-yyyy") & "'))"
    
      
    Adodc4.RecordSource = "select sum(dr_amt) as debit ,sum(cr_amt) as credit  from ledger where to_char(VOU_DATE,'dd-mon-yyyy')<('" & Format(dtst_dt.Value, "dd-mmm-yyyy") & "') and acc_code in(select acc_code from acct where user_acc='" & Trim(cboStUserCode.Text) & "')"
      
      
      
    Adodc4.Refresh
 
    If Adodc4.Recordset.RecordCount > 0 Then



        If IsNull(Adodc4.Recordset.Fields(0)) Then
            dr_amt = 0
        Else
            dr_amt = Val(Adodc4.Recordset!debit)
        End If
         If IsNull(Adodc4.Recordset!credit) = True Then
            cr_amt = 0
        Else
            cr_amt = Val(Adodc4.Recordset!credit)
        End If

        txtOpenBal.Text = Val(dr_amt) - Val(cr_amt)

    End If
    '=======================End======================
    CRViewer1.Show vbModal
    
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    
    Resume Next
End Sub

Private Sub dted_dt_CloseUp()
'    dted_dt.MaxDate = objectCompSetup.ed_dt
'    dted_dt.MinDate = objectCompSetup.st_dt
End Sub

Private Sub dted_dt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub dted_dt_LostFocus()
    dted_dt_CloseUp
End Sub


Private Sub dtst_dt_CloseUp()
'    dtst_dt.MaxDate = objectCompSetup.ed_dt
'    dtst_dt.MinDate = objectCompSetup.st_dt
End Sub

Private Sub dtst_dt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub dtst_dt_LostFocus()
    dtst_dt_CloseUp
End Sub

Private Sub Form_DblClick()




'    Adodc3.ConnectionString = strcn.Connection_String
''    Adodc3.RecordSource = "select sum(dr_amt) as debit,sum(cr_amt) as credit from ledger where  acc_code in(select acc_code from acct where user_acc='" & Trim(cboStUserCode.Text) & "' and  vou_date < '" & CDate(dtst_dt.Value) & "')"
'    Adodc3.RecordSource = "select sum(dr_amt) as debit,sum(cr_amt) as credit from ledger where  acc_code in(select acc_code from acct where user_acc='" & Trim(cboStUserCode.Text) & "') "
''    Debug.Print Adodc3.RecordSource
'    Adodc3.Refresh
'    If Adodc3.Recordset.RecordCount > 0 Then
'        Dim dr_amt, cr_amt As Double
'        If IsNull(Adodc3.Recordset!debit) = True Then
'           dr_amt = 0
'        Else
'            dr_amt = Val(Adodc3.Recordset!debit)
'        End If
'         If IsNull(Adodc3.Recordset!credit) = True Then
'            cr_amt = 0
'        Else
'            cr_amt = Val(Adodc3.Recordset!credit)
'        End If
'
'        txtOpenBal.Text = Val(dr_amt) - Val(cr_amt)
'    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub Form_Load()
    
    Call GetUserAcc
    rptMode = 18
    
'    objectCompSetup.Flush_Comp (strcn)
    
    dtst_dt.Value = Date
    dted_dt.Value = Date
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form9 = Nothing
End Sub
Private Sub GetUserAcc()
    On Error GoTo err_loop
        Me.cboStUserCode.Clear
        Me.cboStAccName.Clear
        
        Adodc1.ConnectionString = strcn.Connection_String
        Adodc1.RecordSource = "select user_acc,acc_name from acct where acc_lbl<>0  and acc_code like '21%' order by user_acc"
        Adodc1.Refresh
        If Adodc1.Recordset.RecordCount > 0 Then
            Do Until Adodc1.Recordset.EOF
            Me.cboStUserCode.AddItem Adodc1.Recordset!user_acc
            Me.cboStAccName.AddItem Adodc1.Recordset!acc_name
            Adodc1.Recordset.MoveNext
            Loop
        End If
       
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

Private Function GetAccName(strUserAcc As String) As String
    On Error GoTo err_loop
        Adodc1.ConnectionString = strcn.Connection_String
        Adodc1.RecordSource = "select acc_name from acct where user_acc='" & Trim(strUserAcc) & "'"
        Adodc1.Refresh
        If Adodc1.Recordset.RecordCount > 0 Then
            GetAccName = Adodc1.Recordset!acc_name
        End If
    Exit Function
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Function
Private Function GetUserCode(strAccName As String) As String
    On Error GoTo err_loop
        Adodc1.ConnectionString = strcn.Connection_String
        Adodc1.RecordSource = "select user_acc from acct where acc_name='" & Trim(strAccName) & "'"
        Adodc1.Refresh
        If Adodc1.Recordset.RecordCount > 0 Then
          GetUserCode = Adodc1.Recordset!user_acc
        End If
    Exit Function
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Function

