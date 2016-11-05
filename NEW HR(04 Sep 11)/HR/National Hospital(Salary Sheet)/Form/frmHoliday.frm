VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   6315
   ClientLeft      =   1050
   ClientTop       =   1590
   ClientWidth     =   9210
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHoliday.frx":0000
   ScaleHeight     =   6315
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdHoliday 
      BackColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   8190
      Picture         =   "frmHoliday.frx":152CE
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5085
      Width           =   555
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Weekend"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   315
      TabIndex        =   21
      Top             =   4995
      Width           =   7845
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Saturday"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   28
         Top             =   270
         Width           =   1005
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sunday"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   1185
         TabIndex        =   27
         Top             =   270
         Width           =   1005
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Monday"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   2235
         TabIndex        =   26
         Top             =   270
         Width           =   1005
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tuesday"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   3285
         TabIndex        =   25
         Top             =   270
         Width           =   1050
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Wednesday"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   4380
         TabIndex        =   24
         Top             =   270
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Thursday"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   5
         Left            =   5700
         TabIndex        =   23
         Top             =   270
         Width           =   1050
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Friday"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   6
         Left            =   6795
         TabIndex        =   22
         Top             =   270
         Value           =   -1  'True
         Width           =   780
      End
   End
   Begin VB.ComboBox cmbYear 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      ItemData        =   "frmHoliday.frx":15B98
      Left            =   360
      List            =   "frmHoliday.frx":15BC6
      TabIndex        =   18
      Top             =   1170
      Width           =   960
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   360
      TabIndex        =   9
      Top             =   2340
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BorderStyle     =   0
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   17
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
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
         Name            =   "Arial"
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
         Caption         =   "Name"
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
         Caption         =   "Start Dt"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd MMM yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   ""
         Caption         =   "End Dt"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd MMM yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   ""
         Caption         =   "Dur"
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
         Caption         =   "Catagory"
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
            DividerStyle    =   6
            ColumnWidth     =   2805.166
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   6
            ColumnWidth     =   1484.787
         EndProperty
         BeginProperty Column02 
            DividerStyle    =   6
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column03 
            DividerStyle    =   6
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column04 
            DividerStyle    =   6
            ColumnWidth     =   1830.047
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5535
      Picture         =   "frmHoliday.frx":15C1E
      ScaleHeight     =   285
      ScaleMode       =   0  'User
      ScaleWidth      =   3264.834
      TabIndex        =   16
      Top             =   5805
      Width           =   3270
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   2445
         Picture         =   "frmHoliday.frx":15F45
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   870
      End
      Begin VB.CommandButton cmdClear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1650
         Picture         =   "frmHoliday.frx":16AC7
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   870
      End
      Begin VB.CommandButton cmdEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   810
         Picture         =   "frmHoliday.frx":17649
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton cmdDel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   -30
         Picture         =   "frmHoliday.frx":181CB
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   810
         Picture         =   "frmHoliday.frx":18D4D
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   870
      End
   End
   Begin VB.TextBox txtHol_name 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   1485
      TabIndex        =   0
      Top             =   1170
      Width           =   2445
   End
   Begin VB.TextBox txtHol_Desc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   1935
      Width           =   8430
   End
   Begin VB.ComboBox cmbCategory 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      ItemData        =   "frmHoliday.frx":198CF
      Left            =   4095
      List            =   "frmHoliday.frx":198E2
      TabIndex        =   1
      Top             =   1170
      Width           =   1905
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7155
      Top             =   360
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
   Begin MSComCtl2.DTPicker dtpEnd_date 
      Height          =   330
      Left            =   7560
      TabIndex        =   3
      Top             =   1170
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   582
      _Version        =   393216
      CalendarForeColor=   8388608
      CalendarTitleForeColor=   8388608
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   58720259
      CurrentDate     =   37004
   End
   Begin MSComCtl2.DTPicker dtpStr_date 
      Height          =   330
      Left            =   6165
      TabIndex        =   2
      Top             =   1170
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   582
      _Version        =   393216
      CalendarForeColor=   8388608
      CalendarTitleForeColor=   8388608
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   58720259
      CurrentDate     =   36995
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FECCC7&
      BorderColor     =   &H00FECCC7&
      Height          =   420
      Index           =   6
      Left            =   315
      Top             =   1125
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   405
      TabIndex        =   19
      Top             =   855
      Width           =   330
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FECCC7&
      BorderColor     =   &H00FECCC7&
      Height          =   420
      Index           =   5
      Left            =   4050
      Top             =   1125
      Width           =   1995
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FECCC7&
      BorderColor     =   &H00FECCC7&
      Height          =   420
      Index           =   4
      Left            =   7515
      Top             =   1125
      Width           =   1320
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FECCC7&
      BorderColor     =   &H00FECCC7&
      Height          =   420
      Index           =   3
      Left            =   6120
      Top             =   1125
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FECCC7&
      BorderColor     =   &H00FECCC7&
      Height          =   330
      Index           =   2
      Left            =   315
      Top             =   1890
      Width           =   8520
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FECCC7&
      BorderColor     =   &H00FECCC7&
      Height          =   2625
      Index           =   1
      Left            =   315
      Top             =   2295
      Width           =   8520
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FECCC7&
      BorderColor     =   &H00FECCC7&
      Height          =   420
      Index           =   0
      Left            =   1440
      Top             =   1125
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Holiday"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1080
      TabIndex        =   15
      Top             =   90
      Width           =   930
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Holiday  Name"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1485
      TabIndex        =   14
      Top             =   855
      Width           =   1035
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From (Date)"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6255
      TabIndex        =   13
      Top             =   855
      Width           =   825
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   315
      TabIndex        =   12
      Top             =   1620
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4140
      TabIndex        =   11
      Top             =   855
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To (Date)"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7650
      TabIndex        =   10
      Top             =   810
      Width           =   675
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Weekend As String


Private Sub cmbYear_Click()
    populate_grd
End Sub

Private Sub cmdClear_Click()
txtHol_name = ""
dtpStr_date = Now
dtpEnd_date = Now
cmbCategory = ""
txtHol_Desc = ""
txtHol_name.SetFocus
End Sub

Private Sub cmdClose_Click()
yes_no = MsgBox("Do you really want to close it?", vbYesNo + vbQuestion)
    If yes_no = vbYes Then
        Unload Me
    Else
    txtHol_name.SetFocus
        Exit Sub
    End If
End Sub

Private Sub cmdDel_Click()
opr = "D"
cmdSave_Click
End Sub

Private Sub cmdEdit_Click()
Grid_Click (True), Form4
cmdSave_Click
End Sub

Private Sub cmdSave_Click()

If txtHol_name = "" Then Exit Sub
 con.ConnectionString = strCN.Connection
    con.Open
    Set Cmd.ActiveConnection = con

    Cmd.CommandText = "exec Hol_List_I_U_D '" + opr + "','" _
    + ChkForQuote(txtHol_name) + "','" _
    + ChkForQuote(txtHol_Desc) + "','" _
    + Format(dtpStr_date, "yyyy-mm-dd") + "','" _
    + Format(dtpEnd_date, "yyyy-mm-dd") + "','" _
    + ChkForQuote(cmbCategory) + "','" _
    + U_Id + "'"

    Cmd.Execute
    con.Close

    cmdClear_Click
    populate_grd
    
    Grid_Click (False), Form4
End Sub

Private Sub CmdHoliday_Click()



If cmbYear = "" Then Exit Sub

    con.ConnectionString = strCN.Connection
    con.Open
    Set Cmd.ActiveConnection = con

    Cmd.CommandText = "Yearly_Holidays '" + cmbYear + "','" _
    + Weekend + "'"
    
    Cmd.Execute
    con.Close

    'cmdClear_Click
    populate_grd

End Sub
Private Sub DataGrid1_DblClick()
If Adodc1.Recordset.EOF Then Exit Sub
    txtHol_name = Adodc1.Recordset!Hol_name
    dtpStr_date.Value = Adodc1.Recordset!str_date
    dtpEnd_date.Value = Adodc1.Recordset!End_date
    cmbCategory = Adodc1.Recordset!Category
    txtHol_Des = Adodc1.Recordset!Hol_Desc
    '-------------------------
    txtHol_name.Refresh
    dtpStr_date.Refresh
    dtpEnd_date.Refresh
    cmbCategory.Refresh
    txtHol_Desc.Refresh
    
    Grid_Click (True), Me
    
End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Grid_Click (False), Me

End Sub

Private Sub dtpStr_date_CloseUp()
dtpEnd_date = dtpStr_date
End Sub

Private Sub Form_Load()
    Weekend = "Friday"
    
    dtpStr_date = Now
    dtpEnd_date = Now
    Grid_Click (False), Form4
    populate_grd
End Sub
Public Sub populate_grd()
    Adodc1.ConnectionString = strCN.Connection
    Adodc1.RecordSource = "exec Pop_Holidays '" + cmbYear + "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount <> 0 Then
        If Adodc1.Recordset.RecordCount > 0 Then
            Adodc1.Recordset.MoveLast
        End If
        Set DataGrid1.DataSource = Adodc1
         
        DataGrid1.Columns(0).DataField = "Hol_name"
        DataGrid1.Columns(1).DataField = "str_date"
        DataGrid1.Columns(2).DataField = "End_date"
        DataGrid1.Columns(3).DataField = "duration"
        DataGrid1.Columns(4).DataField = "Category"
        DataGrid1.ReBind
        DataGrid1.Refresh
    End If
End Sub
Private Sub Option1_Click(index As Integer)
    Weekend = Option1(index).Caption
End Sub
