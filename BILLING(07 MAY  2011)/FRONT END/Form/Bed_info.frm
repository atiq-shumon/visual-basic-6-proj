VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Bed_info 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9225
   Icon            =   "Bed_info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox SERIAL_NO 
      Height          =   315
      Left            =   0
      TabIndex        =   29
      Text            =   "0"
      Top             =   5460
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   7920
      TabIndex        =   28
      Top             =   6030
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   6690
      TabIndex        =   27
      Top             =   6030
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   5460
      TabIndex        =   9
      Top             =   6030
      Width           =   1215
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "NEW"
      Height          =   375
      Left            =   4230
      TabIndex        =   8
      Top             =   6030
      Width           =   1215
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   6030
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   7290
      TabIndex        =   12
      Text            =   "0"
      Top             =   5430
      Visible         =   0   'False
      Width           =   1515
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   345
      Left            =   15
      Top             =   6345
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
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
      Height          =   345
      Left            =   4545
      Top             =   5925
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   609
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
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5925
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9225
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Bed_info.frx":08CA
         Height          =   4440
         Left            =   210
         TabIndex        =   11
         Top             =   1440
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   7832
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   -2147483624
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   0
         TabIndex        =   25
         Top             =   -120
         Width           =   9225
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BED INFORMATION SETUP"
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
            Left            =   2580
            TabIndex        =   26
            Top             =   240
            Width           =   4755
         End
         Begin VB.Image Image1 
            Height          =   855
            Left            =   -270
            Picture         =   "Bed_info.frx":08DF
            Top             =   -120
            Width           =   11820
         End
      End
      Begin VB.ComboBox doc_dept 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc3"
         Height          =   315
         ItemData        =   "Bed_info.frx":6261
         Left            =   6840
         List            =   "Bed_info.frx":6263
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1095
         Width           =   1230
      End
      Begin VB.TextBox txtServiceCharge 
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
         Left            =   5970
         TabIndex        =   6
         Top             =   1095
         Width           =   840
      End
      Begin VB.TextBox txtFee 
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
         Index           =   3
         Left            =   5085
         TabIndex        =   5
         Text            =   "0"
         Top             =   1095
         Width           =   840
      End
      Begin VB.TextBox Text1 
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
         Index           =   0
         Left            =   3555
         TabIndex        =   3
         Top             =   1095
         Width           =   660
      End
      Begin VB.TextBox Text1 
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
         Index           =   1
         Left            =   4245
         TabIndex        =   4
         Text            =   "0"
         Top             =   1095
         Width           =   810
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Bed_info.frx":6265
         Left            =   2595
         List            =   "Bed_info.frx":62EA
         TabIndex        =   2
         Text            =   "Combo3"
         Top             =   1095
         Width           =   960
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Bed_info.frx":638F
         Left            =   570
         List            =   "Bed_info.frx":639F
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   1095
         Width           =   1275
      End
      Begin VB.ComboBox Combo4 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Bed_info.frx":63C8
         Left            =   1830
         List            =   "Bed_info.frx":63FC
         TabIndex        =   1
         Text            =   "Combo4"
         Top             =   1095
         Width           =   780
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   8055
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1095
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58261505
         CurrentDate     =   38027
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         Height          =   195
         Index           =   9
         Left            =   6990
         TabIndex        =   24
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "S. Charge"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   6000
         TabIndex        =   21
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type#"
         Height          =   195
         Index           =   7
         Left            =   1875
         TabIndex        =   20
         Top             =   855
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         Height          =   195
         Index           =   6
         Left            =   8040
         TabIndex        =   19
         Top             =   855
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bed No"
         Height          =   195
         Index           =   0
         Left            =   3600
         TabIndex        =   18
         Top             =   855
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Type"
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   17
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Charge"
         Height          =   195
         Index           =   2
         Left            =   4230
         TabIndex        =   16
         Top             =   855
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A.Fee"
         Height          =   195
         Index           =   3
         Left            =   5340
         TabIndex        =   15
         Top             =   855
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seat Capacity"
         Height          =   195
         Index           =   5
         Left            =   2535
         TabIndex        =   14
         Top             =   855
         Width           =   1080
      End
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   30
      Top             =   6330
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
   Begin VB.Shape Shape2 
      Height          =   465
      Left            =   2910
      Top             =   5970
      Width           =   6285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bed Flag"
      Height          =   195
      Index           =   4
      Left            =   5400
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   630
   End
End
Attribute VB_Name = "Bed_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Con As New MyConnection
Dim Conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset
'Public rptMode As Integer
Public strUid As String
Public strcn        As New MyConnection
Private Sub flush_grid()
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select bed_type,bed_ext_col as type_no,seat_capacity as capacity,bed_no as No ,bed_charge as charge,bed_group as Admission_fee,service_charge, DOC_DEPARTMENT,temp_date as Effective_Date ,serial_no from Bed_info  order by serial_no desc "  '''bed_type ,
    Adodc1.Refresh
        
    DataGrid1.Columns(0).Width = 1300
    DataGrid1.Columns(1).Width = 720
    DataGrid1.Columns(2).Width = 1000
    DataGrid1.Columns(3).Width = 670
    DataGrid1.Columns(4).Width = 850
    DataGrid1.Columns(5).Width = 850
    DataGrid1.Columns(6).Width = 900
    DataGrid1.Columns(7).Width = 1200
    DataGrid1.Columns(8).Width = 1200
    
End Sub

Private Sub clear()
'On Error Resume Next
Text1(1).Text = 0
Text1(0).SetFocus

Combo1 = Combo1.List(0)
Combo4 = Combo4.List(0)
txtFee(3).Text = 0
Combo3 = Combo3.List(0)

Text1(0).SetFocus
End Sub
Private Sub cmdADD_Click()
Call clear
End Sub

Private Sub CMDDELETE_Click()

flush_grid

Dim reply As String
    reply = MsgBox("Do you want to Delete?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
    
Dim i As Integer
If Conn.State = 0 Then
    Conn.ConnectionString = strcn.Connection_String
    Conn.Open
End If
cmd.ActiveConnection = Conn
cmd.CommandType = adCmdText
cmd.CommandText = "delete from  Bed_info where SERIAL_NO='" & Trim(SERIAL_NO.Text) & "'"
cmd.Execute
If Conn.State = 1 Then
         Conn.Close
    Set Conn = Nothing
End If
     Text1(1) = 0
     Text1(2) = ""

 End If
 Call flush_grid
 Call clear

End Sub

Private Sub cmdExit_Click()
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdSave_Click()
If Text1(0) = "" Then
     MsgBox "Bed no Requied"
        Text1(0).SetFocus
        Exit Sub
   End If
   
If Text1(1) < 0 Then
     MsgBox "Bed Charge Requied"
        Text1(1).SetFocus
        Exit Sub
   End If
       
    Call SaveBedInfo
    MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
'    Call FlushCompSetup
Call flush_grid
Call clear
End Sub

Private Sub SaveBedInfo()

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
    Dim Param11 As New Parameter
    Dim Param12 As New Parameter
       
If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
End If
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 15, Text1(0).Text)
    cmd.Parameters.Append Param1 'Bed_no


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, Combo1.Text)
    cmd.Parameters.Append Param2 'Bed_Type
    
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, Text1(1).Text)
    cmd.Parameters.Append Param3 'Bed_charge
    
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 8, txtFee(3).Text)
    cmd.Parameters.Append Param4 'Fee
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 2, Text1(2).Text)
    cmd.Parameters.Append Param5 'Occupy flag default 0


    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 2, "na")
    cmd.Parameters.Append Param6 'U_id default Sumon

    Set Param7 = cmd.CreateParameter("param7", adDate, adParamInput, 12, DTPicker1.Value)
    cmd.Parameters.Append Param7 'Tmp Date default sysdate


    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 2, Null)
    cmd.Parameters.Append Param8 'IN_REG_NO
    
    
    Set Param9 = cmd.CreateParameter("param9", adInteger, adParamInput, 8, Combo3.Text)
    cmd.Parameters.Append Param9 'Seat_capacity

    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 25, Combo4.Text)
    cmd.Parameters.Append Param10 'TYPE_NO

    Set Param11 = cmd.CreateParameter("param11", adDouble, adParamInput, 10, txtServiceCharge.Text)
    cmd.Parameters.Append Param11 'Servce charge
    
     Set Param12 = cmd.CreateParameter("param12", adVarChar, adParamInput, 15, Me.doc_dept)
    cmd.Parameters.Append Param12 'dept
    
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL SaveBed_info(?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
        Conn.Close
    End If
    
End Sub

Private Sub DataGrid1_Click()
    If Adodc1.Recordset.RecordCount = 0 Then
                MsgBox "Nothing to show", vbInformation, "Warning: IT, DNMIH"
 
 Else
 
     Combo1.Text = DataGrid1.Columns(0).Text
    Combo4.Text = DataGrid1.Columns(1).Text
     txtFee(3) = DataGrid1.Columns(5).Text
    Combo3.Text = DataGrid1.Columns(2).Text
     Text1(0).Text = DataGrid1.Columns(3).Text
     Text1(1).Text = DataGrid1.Columns(4).Text
     txtServiceCharge.Text = DataGrid1.Columns(6)
     Me.doc_dept = DataGrid1.Columns(7)
     DTPicker1.Value = DataGrid1.Columns(8)
     SERIAL_NO.Text = DataGrid1.Columns(9)
     
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
End If
End Sub



Private Sub Form_Load()

      
Combo1 = Combo1.List(0)
Combo4 = Combo4.List(0)
'Combo2 = TxtFee.List(0)
Combo3 = Combo3.List(0)
      Adodc3.ConnectionString = strcn.Connection_String
      Adodc3.RecordSource = "select distinct(doc_dept) from doctor_info"
      Adodc3.Refresh

    If Adodc3.Recordset.RecordCount > 0 Then
        Adodc3.Recordset.MoveFirst
        While Adodc3.Recordset.EOF = False
          doc_dept.AddItem Adodc3.Recordset!doc_dept
          Adodc3.Recordset.MoveNext
        Wend
    End If
    
    If doc_dept.ListCount <> Empty Then
         doc_dept = doc_dept.List(0)
 End If
 
Call flush_grid

End Sub

Private Sub Flush_Data()

'   On Error Resume Next
   Dim i As Integer
    Adodc2.ConnectionString = strcn.Connection_String
    Adodc2.RecordSource = "select bed_no ,bed_type,bed_charge from Bed_info where bed_no='" & Trim(Text1(0).Text) & "'"
    Adodc2.Refresh
    
    If Adodc2.Recordset.RecordCount > 0 Then
        

         For i = 0 To Adodc2.Recordset.Fields.Count
           Text1(i).Text = Adodc2.Recordset.Fields(i).Value
'
           If i = 1 Then   '' Marital Status
                Combo1.ListIndex = (Adodc2.Recordset.Fields(i).Value)
            End If
           
         Next
    Else
    End If
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)

            
        Call Flush_Data

End Sub

Private Sub txtFee_Change(Index As Integer)
If Not IsNumeric(txtFee(3).Text) Then
            txtFee(3) = ""
End If
End Sub
