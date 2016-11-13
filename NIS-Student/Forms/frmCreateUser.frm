VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmCreateUser 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H80000001&
      Height          =   945
      Left            =   -60
      TabIndex        =   18
      Top             =   -150
      Width           =   8145
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Creation"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEF0F7&
         Height          =   270
         Index           =   2
         Left            =   2880
         TabIndex        =   25
         Top             =   420
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   1260
         Left            =   60
         Picture         =   "frmCreateUser.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   7245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Security :User Creation "
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   345
         Index           =   0
         Left            =   2340
         TabIndex        =   19
         Top             =   330
         Width           =   3120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Security :User Creation "
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   1
         Left            =   2280
         TabIndex        =   21
         Top             =   330
         Width           =   3120
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   5880
         Picture         =   "frmCreateUser.frx":CEA5
         Top             =   300
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   1230
         Picture         =   "frmCreateUser.frx":D2E7
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Height          =   825
      Left            =   -30
      TabIndex        =   17
      Top             =   4890
      Width           =   8115
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   1110
         Top             =   210
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
            Name            =   "Arial"
            Size            =   9.75
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
         Left            =   330
         Top             =   480
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
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton cmdSAVE 
         BackColor       =   &H80000016&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2610
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdADD 
         BackColor       =   &H80000016&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEXIT 
         BackColor       =   &H80000016&
         Cancel          =   -1  'True
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   6030
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H80000016&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         Height          =   435
         Left            =   2580
         Top             =   330
         Width           =   4575
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000000&
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         Height          =   75
         Index           =   0
         Left            =   0
         Top             =   60
         Width           =   9045
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   -30
      TabIndex        =   12
      Top             =   780
      Width           =   8175
      Begin VB.CommandButton cboSearch 
         Caption         =   ":::"
         Height          =   345
         Left            =   3150
         TabIndex        =   24
         Top             =   180
         Width           =   405
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1470
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1470
         Width           =   2730
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1470
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1020
         Width           =   2730
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   0
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   0
         Top             =   180
         Width           =   1665
      End
      Begin VB.ComboBox cboType 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   5070
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   1965
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   5610
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retype Password :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   4
         Left            =   45
         TabIndex        =   20
         Top             =   1530
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00D5A47B&
         BorderWidth     =   2
         Height          =   75
         Index           =   1
         Left            =   0
         Top             =   30
         Width           =   9465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   3
         Left            =   630
         TabIndex        =   16
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Type :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   4080
         TabIndex        =   15
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   495
         TabIndex        =   14
         Top             =   660
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   765
         TabIndex        =   13
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2265
      Left            =   -30
      TabIndex        =   22
      Top             =   2640
      Width           =   8205
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmCreateUser.frx":D729
         Height          =   1980
         Left            =   60
         TabIndex        =   23
         Top             =   240
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   3493
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
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
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sec. Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2760
      TabIndex        =   11
      Top             =   2970
      Width           =   825
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00FFC0C0&
      Height          =   345
      Left            =   3690
      Top             =   2880
      Width           =   2835
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sec. Code"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   150
      TabIndex        =   10
      Top             =   2970
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Access Information "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   330
      Left            =   180
      TabIndex        =   9
      Top             =   3360
      Width           =   3270
   End
End
Attribute VB_Name = "FrmCreateUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Temp_Tab As ADODB.Recordset
Dim Temp_Tab_Helper As New ADODB.Recordset


Private Sub cmdADD_Click(Index As Integer)
    txtfield(1).Text = ""
    txtfield(2).Text = ""
    txtfield(0).SetFocus
End Sub
Private Sub cmdDelete_Click(Index As Integer)
   Dim Pass As String
    If Len(Trim(txtfield(0).Text)) = 0 Then Exit Sub
    If Trim(txtfield(0).Text) = UCase(strUid) Then
       MsgBox "You can't delete yourself", vbCritical
       Exit Sub
    End If
    
    If Len(Trim(txtfield(2).Text)) = "" Then
        MsgBox "Previous password required", vbCritical, "Daffodil Software Ltd."
        txtfield(2).SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtfield(2).Text)) <> "" Then
            Adodc1.connectionstring = strcn.connection
            Adodc1.RecordSource = "select pass from security where user_id='" & Trim(txtfield(0).Text) & "'"
            Adodc1.Refresh
        
            If Adodc1.Recordset.RecordCount > 0 Then
                Pass = Adodc1.Recordset!Pass
             End If
            
        If Pass <> txtfield(2) Then
            MsgBox "Invalid password,You can't delete", vbCritical, "Daffodil Software Ltd."
             txtfield(2).SetFocus
             Exit Sub
        End If
    
    End If
    con.connectionstring = strcn.connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "Delete from security where user_id='" & Trim(txtfield(0).Text) & "'"
    cmd.Execute
    con.Close
    
    MsgBox "Delete successfully", vbInformation
    txtfield(1).Text = ""
    txtfield(2).Text = ""
    txtfield(0).SetFocus
     getdata
End Sub
Private Sub cmdExit_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdSAVE_Click(Index As Integer)
    If Len(Trim(txtfield(0).Text)) = 0 Then
       MsgBox "User id required", vbCritical
       txtfield(0).SetFocus
       Exit Sub
    End If

    If Len(Trim(txtfield(1).Text)) = 0 Then
       MsgBox "User name required", vbCritical
      txtfield(1).SetFocus
       Exit Sub
    End If

    If Len(Trim(txtfield(3).Text)) = 0 Then
       MsgBox "Password required", vbCritical
       txtfield(3).SetFocus
       Exit Sub
    End If

    Adodc1.connectionstring = strcn.connection
    Adodc1.RecordSource = "select * from security where user_id='" & Trim(txtfield(0).Text) & "'"
    Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
       Dim strEdit As String
       strEdit = MsgBox("Are you sure you want to edit the current user?", vbQuestion + vbYesNo)
       If strEdit = vbYes Then
          '''updSecurity
          MsgBox "Update successfully", vbInformation
       End If
    Else
       strEdit = MsgBox("Are you sure you want to add the current user?", vbQuestion + vbYesNo)
       If strEdit = vbYes Then
          insSecurity
          MsgBox "Save successfully", vbInformation
           txtfield(2).Text = ""
           txtfield(3).Text = ""
           txtfield(0).SetFocus
           getdata
           
       End If
    End If
End Sub

Private Sub DataGrid1_Click()
    flush_grid
End Sub
Private Sub flush_grid()
    If Adodc2.Recordset.RecordCount > 0 Then
                txtfield(0).Text = "" & DataGrid1.Columns(0).Text
                txtfield(1).Text = "" & DataGrid1.Columns(1).Text
                cboType = "" & DataGrid1.Columns(2).Text
     End If
   
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   
     If KeyCode = 13 Then
        SendKeys Chr(9)
     End If
End Sub

Private Sub Form_Load()
    
    cboType.AddItem "Admin"  ''''Admin
    cboType.AddItem "User" ''''User
    cboType.AddItem "Super User"  ''''Super User
    
    cboType.Text = "Admin"
    getdata
    
'    If UCase(validation_var) <> UCase("admin") Then
'        cmdSave(0).Enabled = False
'        cmdDelete(2).Enabled = False
'     Else
'        cmdSave(0).Enabled = True
'        cmdDelete(2).Enabled = True
'     End If

End Sub
Private Sub getdata()
   Adodc2.connectionstring = strcn.connection
   Adodc2.RecordSource = "Select user_id,user_name,type from security "
   Adodc2.Refresh
   
    DataGrid1.Columns(0).Width = 2000
    DataGrid1.Columns(1).Width = 3500
    DataGrid1.Columns(2).Width = 2000
End Sub

'Private Sub updSecurity()
'    con.connectionstring = strcn.adodb.connection
'    con.Open
'
'    Set cmd.ActiveConnection = con
'    cmd.CommandText = " exec pro_security 'U','" & Trim(txtfield(0).Text) & "'" & _
'                      " ,'" & Trim(txtfield(2).Text) & "'" & _
'                      " ,'" & Trim(txtfield(2).Text) & "'" & _
'                      " ,'" & Trim(cboType.Text) & "'" & _
'                      " ," & Val(cboSec_code.Text) & "" & _
'                      " ,'" & Trim(strUid) & "'"
'    cmd.Execute
'    con.Close
'End Sub
Private Sub insSecurity()
    con.connectionstring = strcn.connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = " insert into security values('" & Trim(txtfield(0).Text) & "'" & _
                      " ,'" & Trim(txtfield(1).Text) & "'" & _
                      " ,'" & Trim(txtfield(2).Text) & "'" & _
                      " ,'" & Trim(cboType.Text) & "'" & _
                      " ," & 1 & " " & _
                      " ,'" & frmMain.lblUid & "'" & _
                      " ,'" & Format(Date, "dd-mmm-yyyy") & "')"
   cmd.Execute
    con.Close
End Sub

Private Sub txtField_LostFocus(Index As Integer)
  Select Case Index
         Case 0
              If Len(Trim(txtfield(0).Text)) = 0 Then Exit Sub
                Adodc1.connectionstring = strcn.connection
                Adodc1.RecordSource = "Select user_name,type from security where user_id='" & Trim(txtfield(0).Text) & "'"
                Adodc1.Refresh
    
            If Adodc1.Recordset.RecordCount > 0 Then
                txtfield(1).Text = Adodc1.Recordset!user_name
                cboType.Text = Adodc1.Recordset!Type
           Else
                txtfield(2).Text = ""
                txtfield(3).Text = ""
            End If
                End Select
    
End Sub
