VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFiscalYear 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fiscal Year SetUp"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6045
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   3510
      Picture         =   "frmFiscalYear.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3465
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   1080
      Picture         =   "frmFiscalYear.frx":1A0A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3465
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   2295
      Picture         =   "frmFiscalYear.frx":339C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3465
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   4725
      Picture         =   "frmFiscalYear.frx":4D2E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3465
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Height          =   1680
      Left            =   135
      TabIndex        =   3
      Top             =   1710
      Width           =   5820
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1455
         Left            =   45
         TabIndex        =   4
         Top             =   135
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   2566
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fiscal Year SetUp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   1545
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   5820
      Begin MSComCtl2.DTPicker Begin_Fiscal_date 
         Height          =   315
         Left            =   450
         TabIndex        =   9
         Top             =   855
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   8388608
         CalendarTitleForeColor=   8388608
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58523651
         CurrentDate     =   36998
      End
      Begin MSComCtl2.DTPicker End_Fiscal_Year 
         Height          =   315
         Left            =   2970
         TabIndex        =   10
         Top             =   855
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   8388608
         CalendarTitleForeColor=   8388608
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58523651
         CurrentDate     =   36998
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00004080&
         Height          =   1005
         Left            =   180
         Shape           =   4  'Rounded Rectangle
         Top             =   405
         Width           =   5460
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Begin Fiscal Year"
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
         Left            =   675
         TabIndex        =   2
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "End Fiscal  Year"
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
         Left            =   3465
         TabIndex        =   1
         Top             =   540
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Index           =   2
         Left            =   405
         Top             =   825
         Width           =   2235
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Index           =   3
         Left            =   2925
         Top             =   825
         Width           =   2475
      End
   End
End
Attribute VB_Name = "frmFiscalYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FiscalYr As New ClsSt_payscale
Private Sub cmdClear_Click()
Me.Begin_Fiscal_date = Date$
Me.End_Fiscal_Year = Date$
End Sub
Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdDelete_Click()
On Error GoTo Errdes
With FiscalYr
    .Connstring = strCN.Connection_String
    .BEGIN_DATE = Begin_Fiscal_date
    .END_DATE = End_Fiscal_Year
     MsgBox "Data Deted  Successfully", vbInformation, "IT Division, DNMIH"
    .Pay_Scale_SetUp_Delete
    Get_Value_Into_Grid
End With
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub cmdSave_Click()
On Error GoTo Errdes
With FiscalYr
    .Connstring = strCN.Connection_String
    .BEGIN_DATE = Begin_Fiscal_date
    .END_DATE = End_Fiscal_Year
    .ENTRY_DATE = Date$
    .ENTRY_BY = U_Id
     MsgBox "Data has Saved Successfully !", vbInformation, "IT Division, DNMIH"
    .Pay_Scale_SetUp_Save
    Get_Value_Into_Grid
End With
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
    
End Sub
Private Sub Get_Value_Into_Grid()
    With FiscalYr
        .Connstring = strCN.Connection_String
     Set DataGrid1.DataSource = .GetAll
    End With

End Sub
Private Sub DataGrid1_Click()
    Me.Begin_Fiscal_date = DataGrid1.Columns(0)
    Me.End_Fiscal_Year = DataGrid1.Columns(1)
End Sub
Private Sub Form_Load()
    Get_Value_Into_Grid
    get_First_Record_to_Show
End Sub
Private Sub get_First_Record_to_Show()
On Error GoTo Errdes
Dim conn As New Connection
Dim cmd As New Command
Dim Rs As New ADODB.Recordset
conn.ConnectionString = strCN.Connection_String
conn.Open
cmd.ActiveConnection = conn
cmd.CommandType = adCmdText
cmd.CommandText = "Select BEGIN_DATE,END_DATE from fISCAL_Year_SetUp " + _
                " where TRACE_ID=(select max(TRACE_ID) from fISCAL_Year_SetUp) "
cmd.Properties("iRowsetChange") = True
cmd.Properties("updatability") = 7
Rs.CursorLocation = adUseClient
Rs.Open cmd.CommandText, conn, adOpenDynamic, adLockOptimistic

If Not Rs.EOF Then
    Begin_Fiscal_date = Rs.Fields(0)
    End_Fiscal_Year = Rs.Fields(1)
End If

Rs.Close
conn.Close
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
    
End Sub

