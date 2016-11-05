VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOperation_info 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Operation Info Setup"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Operation_info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   22
      Top             =   5820
      Width           =   11895
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
         Left            =   2235
         Picture         =   "Operation_info.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Exit"
         Top             =   195
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
         Left            =   60
         Picture         =   "Operation_info.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Save"
         Top             =   210
         Width           =   495
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
         Left            =   1695
         Picture         =   "Operation_info.frx":1852
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Preview"
         Top             =   195
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
         Left            =   600
         Picture         =   "Operation_info.frx":1EBC
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "New"
         Top             =   195
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
         Height          =   465
         Left            =   1155
         Picture         =   "Operation_info.frx":2526
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Delete"
         Top             =   195
         Width           =   510
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   3720
         Top             =   270
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   6195
         Top             =   285
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
         Left            =   5010
         Top             =   270
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
      Begin VB.Shape Shape1 
         BorderStyle     =   6  'Inside Solid
         Height          =   555
         Left            =   0
         Top             =   150
         Width           =   2805
      End
   End
   Begin VB.TextBox txtAnnayCharge 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   8835
      TabIndex        =   6
      Top             =   1230
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Operation_info.frx":3060
      Height          =   4080
      Left            =   0
      TabIndex        =   12
      Top             =   1740
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   7197
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Height          =   1740
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11850
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   1005
         Left            =   -30
         TabIndex        =   28
         Top             =   -90
         Width           =   11895
         Begin VB.Image Image1 
            Height          =   480
            Left            =   6030
            Picture         =   "Operation_info.frx":3075
            Top             =   270
            Width           =   480
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Operation Information Setup"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   675
            Left            =   6750
            TabIndex        =   29
            Top             =   240
            Width           =   5205
         End
      End
      Begin VB.TextBox txtServiceCharge 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   9697
         TabIndex        =   7
         Top             =   1230
         Width           =   900
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   10575
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   529
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   38054
      End
      Begin VB.ComboBox cboOprDept 
         DataSource      =   "Adodc3"
         Height          =   315
         ItemData        =   "Operation_info.frx":393F
         Left            =   5655
         List            =   "Operation_info.frx":3941
         TabIndex        =   3
         Top             =   1230
         Width           =   1125
      End
      Begin VB.ComboBox cboOprBed 
         Height          =   315
         ItemData        =   "Operation_info.frx":3943
         Left            =   6750
         List            =   "Operation_info.frx":3950
         TabIndex        =   4
         Top             =   1230
         Width           =   960
      End
      Begin VB.ComboBox cboOprName 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Operation_info.frx":3963
         Left            =   915
         List            =   "Operation_info.frx":3965
         TabIndex        =   1
         Text            =   "Combo2"
         Top             =   1230
         Width           =   3750
      End
      Begin VB.TextBox TxtOPrCharge 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7703
         TabIndex        =   5
         Top             =   1230
         Width           =   1125
      End
      Begin VB.ComboBox cboOprType 
         Height          =   315
         ItemData        =   "Operation_info.frx":3967
         Left            =   4635
         List            =   "Operation_info.frx":3969
         TabIndex        =   2
         Top             =   1230
         Width           =   1050
      End
      Begin VB.TextBox txtOprcode 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   330
         TabIndex        =   0
         Top             =   1230
         Width           =   585
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "S.Charge"
         Height          =   180
         Left            =   9765
         TabIndex        =   21
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective date"
         Height          =   180
         Left            =   10620
         TabIndex        =   20
         Top             =   1005
         Width           =   1005
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "A.Charge"
         Height          =   210
         Left            =   8850
         TabIndex        =   19
         Top             =   1005
         Width           =   1005
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Case"
         Height          =   210
         Left            =   6960
         TabIndex        =   16
         Top             =   1005
         Width           =   585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Charge"
         Height          =   210
         Left            =   7980
         TabIndex        =   15
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   210
         Left            =   4830
         TabIndex        =   14
         Top             =   975
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         Height          =   195
         Left            =   5640
         TabIndex        =   13
         Top             =   975
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Code"
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   11
         Top             =   990
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operation  Name"
         Height          =   195
         Index           =   1
         Left            =   1980
         TabIndex        =   10
         Top             =   1005
         Width           =   1200
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C9AD8F&
      Caption         =   "Charge"
      Height          =   330
      Left            =   8370
      TabIndex        =   18
      Top             =   360
      Width           =   1185
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C9AD8F&
      Caption         =   "Charge"
      Height          =   330
      Left            =   8955
      TabIndex        =   17
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmOperation_info"
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


Private Sub CMDDELETE_Click()
Dim reply As String
    reply = MsgBox("Do you want to Delete?", vbQuestion + vbYesNo, "Delete...")
    If reply = vbYes Then
    If txtOprcode(0).Text = "" Then
    MsgBox "Operation code Required", vbInformation, " IT, DNMIH"
    Else
    
        Call DeleteOperationInfo
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
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
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, txtOprcode(0).Text)
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
    Adodc1.RecordSource = "select OPR_code,opr_name,opr_type,opr_department,opr_bed,opr_charge,annay_charge,service_charge,effective_date from Operation_info order by OPR_code"
    
       Adodc1.Refresh
    DataGrid1.Columns(0).Width = 530
    DataGrid1.Columns(1).Width = 3700
    DataGrid1.Columns(2).Width = 1050
    DataGrid1.Columns(3).Width = 1140
    DataGrid1.Columns(4).Width = 900
    DataGrid1.Columns(5).Width = 1100
    DataGrid1.Columns(6).Width = 900
     DataGrid1.Columns(7).Width = 900
     DataGrid1.Columns(8).Width = 1200
    
    
End Sub

Private Sub cmdSave_Click()
    If txtOprcode(0) = Empty Then
        MsgBox "Main code Requied"
        txtOprcode(0).SetFocus
        Exit Sub
   End If
   If cboOprName = Empty Then
        MsgBox "Opreation Name Requied"
        cboOprName.SetFocus
        Exit Sub
   End If
   
   If TxtOPrCharge = Empty Then
        MsgBox "Operation charge Requied"
        TxtOPrCharge.SetFocus
        Exit Sub
   End If
   
   If txtannayCharge = Empty Then
        MsgBox "Annay Charge Requied"
        txtannayCharge.SetFocus
        Exit Sub
   End If
   If txtServiceCharge = Empty Then
        MsgBox "Service Charge Requied"
        txtServiceCharge.SetFocus
        Exit Sub
   End If
   



    Call SaveOprInfo
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
    Call flush_grid

Clear_form

End Sub

Private Sub SaveOprInfo()

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
    
    
    
   
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, txtOprcode(0).Text)
    cmd.Parameters.Append Param1 'Opr_code

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 60, cboOprName.Text)
    cmd.Parameters.Append Param2 'Opr_name
    
     
   
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 25, cboOprType.Text)
    cmd.Parameters.Append Param4 'opr Type
    
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 30, cboOprDept.Text)
    cmd.Parameters.Append Param5 'Department
    
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 30, Trim(cboOprBed.Text))
    cmd.Parameters.Append Param6 'Opr_bed
    
    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 10, TxtOPrCharge.Text)
    cmd.Parameters.Append Param7 'charge
    
    Set Param8 = cmd.CreateParameter("param8", adDouble, adParamInput, 10, Trim(txtannayCharge.Text))
    cmd.Parameters.Append Param8 ''annay_charge
    
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, Trim(txtServiceCharge.Text))
    cmd.Parameters.Append Param10 ''Service_charge
    
    
    Set Param9 = cmd.CreateParameter("param9", adDate, adParamInput, 10, DTPicker1.Value)
    cmd.Parameters.Append Param9 'Effective date
    
  
     Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, "NHMI")
    cmd.Parameters.Append Param3 'u_id
     
    
   
'----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Save_Operation_info(?,?,?,?,?,?,?,?,?,?)}"
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

txtOprcode(0).Text = DataGrid1.Columns(0)
cboOprName.Text = DataGrid1.Columns(1)
cboOprType.Text = DataGrid1.Columns(2)
cboOprDept.Text = DataGrid1.Columns(3)
cboOprBed.Text = DataGrid1.Columns(4)

TxtOPrCharge.Text = DataGrid1.Columns(5)
txtannayCharge.Text = DataGrid1.Columns(6)
txtServiceCharge.Text = DataGrid1.Columns(7)

DTPicker1.Value = DataGrid1.Columns(8)

End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
End If
End Sub

Private Sub Form_Load()

    Call flush_grid
  
      Adodc3.ConnectionString = strcn.Connection_String
      Adodc3.RecordSource = "select distinct(doc_dept) from doctor_info"
      Adodc3.Refresh

    If Adodc3.Recordset.RecordCount > 0 Then
        Adodc3.Recordset.MoveFirst
        While Adodc3.Recordset.EOF = False
            cboOprDept.AddItem Adodc3.Recordset!doc_dept
            Adodc3.Recordset.MoveNext
        Wend
    End If
  cboOprDept = cboOprDept.List(0)

End Sub





Private Sub Clear_form()

    txtOprcode(0).Text = ""
  
    cboOprName.Text = ""
   cboOprType.Text = ""
   TxtOPrCharge = ""
   txtannayCharge = ""
   txtServiceCharge = ""
   
    
End Sub

Private Sub txtAnnayCharge_Change()
If Not IsNumeric(txtannayCharge.Text) Then
           txtannayCharge = ""
End If
End Sub

Private Sub txtOprCharge_Change()
If Not IsNumeric(TxtOPrCharge.Text) Then
            TxtOPrCharge = ""
End If

End Sub

Private Sub txtOprcode_Change(Index As Integer)
If Not IsNumeric(txtOprcode(0).Text) Then
          txtOprcode(0) = ""
End If
End Sub

Private Sub txtServiceCharge_Change()
If Not IsNumeric(txtServiceCharge.Text) Then
           txtServiceCharge = ""
End If
End Sub
