VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Author Information"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   4110
      Top             =   750
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
      Bindings        =   "AuthorInfo.frx":0000
      Height          =   1965
      Left            =   90
      TabIndex        =   10
      Top             =   1680
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3466
      _Version        =   393216
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4110
      TabIndex        =   9
      Top             =   3720
      Width           =   1125
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2970
      TabIndex        =   8
      Top             =   3720
      Width           =   1125
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1830
      TabIndex        =   7
      Top             =   3720
      Width           =   1125
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   3720
      Width           =   1125
   End
   Begin VB.TextBox txtAuthorName 
      Height          =   315
      Left            =   2100
      TabIndex        =   4
      Top             =   1230
      Width           =   2205
   End
   Begin VB.TextBox txtAuthorCode 
      Height          =   315
      Left            =   2100
      TabIndex        =   3
      Top             =   840
      Width           =   1545
   End
   Begin VB.Label Label4 
      Caption         =   "Author Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1740
      TabIndex        =   5
      Top             =   150
      Width           =   2925
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Author Name"
      Height          =   195
      Left            =   870
      TabIndex        =   1
      Top             =   1260
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Author Code"
      Height          =   195
      Left            =   870
      TabIndex        =   0
      Top             =   900
      Width           =   885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label6_Click()
End Sub

Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""

End Sub

Private Sub Command4_Click()
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Save()
'    On Error GoTo err_loop
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = " exec AuthorInfo_Save '" & Trim(Me.txtAuthorCode) & "'" & _
                      " ,'" & Trim(txtAuthorName.Text) & "'" & _
                      " , '" & Trim(strUid1) & " '"
    
    Set RS = cmd.Execute
    
    MsgBox RS!msg, vbInformation
    con.Close
    

    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

Private Sub cmdNew_Click()
txtAuthorCode = ""
txtAuthorName = ""

End Sub

Private Sub cmdSave_Click()
    Call Save
    Call FlushGrid
End Sub

Private Sub Form_Load()
    Call FlushGrid
End Sub
Private Sub FlushGrid()
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "select * from AuthorInfo"
    Adodc1.Refresh
End Sub
    
