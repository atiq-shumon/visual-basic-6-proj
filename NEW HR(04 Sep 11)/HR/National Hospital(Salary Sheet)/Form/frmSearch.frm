VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSearch 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Search..."
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   Icon            =   "frmSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form21"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraResult 
      BackColor       =   &H00FFFFFF&
      Height          =   6540
      Left            =   3150
      TabIndex        =   2
      Top             =   0
      Width           =   8745
      Begin MSDataGridLib.DataGrid dtgResult 
         Height          =   5595
         Left            =   90
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "   Select search option by pressing  Ctrl + underlined character  "
         Top             =   135
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   9869
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BorderStyle     =   0
         ForeColor       =   13238272
         HeadLines       =   1
         RowHeight       =   19
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
      Begin VB.CommandButton cmdResize 
         BackColor       =   &H00FFEEEC&
         Height          =   400
         Index           =   0
         Left            =   135
         Picture         =   "frmSearch.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "  Ctrl + Left Arrow  "
         Top             =   5895
         Width           =   915
      End
      Begin VB.CommandButton cmdResize 
         BackColor       =   &H00FFEEEC&
         Height          =   400
         Index           =   1
         Left            =   135
         Picture         =   "frmSearch.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "  Ctrl + Right Arrow  "
         Top             =   5895
         Width           =   915
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   6570
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   6540
      Index           =   0
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "    Select search option by pressing  Ctrl + underlined character  "
      Top             =   0
      Width           =   3165
      Begin VB.CommandButton cmdSearch 
         Height          =   400
         Left            =   180
         Picture         =   "frmSearch.frx":114E
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "  Ctrl + S  "
         Top             =   5895
         Width           =   1050
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FCF8F8&
         BorderStyle     =   0  'None
         Height          =   4830
         Left            =   45
         TabIndex        =   12
         Top             =   135
         Width           =   2985
         Begin VB.OptionButton optSearch 
            BackColor       =   &H00FCF8F8&
            Caption         =   "Employee &Name"
            ForeColor       =   &H00C000C0&
            Height          =   285
            Index           =   0
            Left            =   270
            TabIndex        =   25
            Top             =   405
            Width           =   1995
         End
         Begin VB.OptionButton optSearch 
            BackColor       =   &H00FCF8F8&
            Caption         =   "&Joining Date"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   270
            TabIndex        =   24
            Top             =   1047
            Width           =   1995
         End
         Begin VB.OptionButton optSearch 
            BackColor       =   &H00FCF8F8&
            Caption         =   "A&ge (more/equal)"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   3
            Left            =   270
            TabIndex        =   23
            Top             =   1368
            Width           =   1995
         End
         Begin VB.OptionButton optSearch 
            BackColor       =   &H00FCF8F8&
            Caption         =   "&Unit"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   4
            Left            =   270
            TabIndex        =   22
            Top             =   1689
            Width           =   1995
         End
         Begin VB.OptionButton optSearch 
            BackColor       =   &H00FCF8F8&
            Caption         =   "&Designation"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   6
            Left            =   270
            TabIndex        =   21
            Top             =   2010
            Width           =   1995
         End
         Begin VB.OptionButton optSearch 
            BackColor       =   &H00FCF8F8&
            Caption         =   "Job &Type"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   7
            Left            =   270
            TabIndex        =   20
            Top             =   2340
            Width           =   1995
         End
         Begin VB.OptionButton optSearch 
            BackColor       =   &H00FCF8F8&
            Caption         =   "Job Du&ration"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   8
            Left            =   270
            TabIndex        =   19
            Top             =   2655
            Width           =   1995
         End
         Begin VB.OptionButton optSearch 
            BackColor       =   &H00FCF8F8&
            Caption         =   "Job &Ending"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   9
            Left            =   270
            TabIndex        =   18
            Top             =   2985
            Width           =   1995
         End
         Begin VB.OptionButton optSearch 
            BackColor       =   &H00FCF8F8&
            Caption         =   "Acc&omodation Type"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   270
            TabIndex        =   17
            Top             =   3300
            Width           =   1995
         End
         Begin VB.OptionButton optSearch 
            BackColor       =   &H00FCF8F8&
            Caption         =   "&Basic Salary"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   11
            Left            =   270
            TabIndex        =   16
            Top             =   3615
            Width           =   2220
         End
         Begin VB.OptionButton optSearch 
            BackColor       =   &H00FCF8F8&
            Caption         =   "&Pay Scale"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   12
            Left            =   270
            TabIndex        =   15
            Top             =   3945
            Width           =   1995
         End
         Begin VB.OptionButton optSearch 
            BackColor       =   &H00FCF8F8&
            Caption         =   "Payment &Mode"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   13
            Left            =   270
            TabIndex        =   14
            Top             =   4275
            Width           =   1995
         End
         Begin VB.OptionButton optSearch 
            BackColor       =   &H00FCF8F8&
            Caption         =   "&Home district"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   270
            TabIndex        =   13
            Top             =   726
            Width           =   1995
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search Option"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   135
            TabIndex        =   26
            Top             =   45
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdExport 
         DisabledPicture =   "frmSearch.frx":2AE0
         Enabled         =   0   'False
         Height          =   400
         Left            =   1845
         Picture         =   "frmSearch.frx":4472
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "  Alt + X  "
         Top             =   5895
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker dtpDtFrom 
         Height          =   300
         Left            =   585
         TabIndex        =   7
         Top             =   5400
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   12582912
         CalendarTrailingForeColor=   16576
         CustomFormat    =   "dd/MM/yy"
         Format          =   65470467
         CurrentDate     =   37722
      End
      Begin MSComCtl2.DTPicker dtpDtTo 
         Height          =   300
         Left            =   1935
         TabIndex        =   8
         Top             =   5400
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   12582912
         CalendarTrailingForeColor=   16576
         CustomFormat    =   "dd/MM/yy"
         Format          =   65470467
         CurrentDate     =   37722
      End
      Begin VB.Label lblFrom_To 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "From                          To"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   5430
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label lblSearchOption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   5085
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim Opt_Index As Integer
'Dim Search_Pattern As String
'Dim Src_Dt1 As String
'Dim Src_Dt2 As String
'Dim Search_Rs As ADODB.Recordset
'
''dtgResult column width= field length*ln+Fine_Tune
'
'Const Ln As Long = 120
'Const Fine_Tune As Long = 100
'
'Const lngFraWd1 As Long = 8745  'initial width of fraResult
'Const lngFraWd2 As Long = 11850 'changed width of fraResult
'Const lngFraLft1 As Long = 3150 'initial Left of fraResult
'Const lngFraLft2 As Long = 45   'changed Left of fraResult
'
'
'Private Sub cmdResize_Click(index As Integer)
'
'Select Case index
'    Case 0  'enlarge
'        With fraResult
'            .Width = lngFraWd2
'            .Left = lngFraLft2
'        End With
'
'        dtgResult.Width = lngFraWd2 - 180
'
'    cmdResize(1).Visible = True     '--> Button visible true
'    Case 1  'original size
'        With fraResult
'            .Width = lngFraWd1
'            .Left = lngFraLft1
'        End With
'        dtgResult.Width = lngFraWd1 - 180
'    cmdResize(0).Visible = True     '<-- Button visible true
'End Select
'
'cmdResize(index).Visible = False
'
'End Sub
'
'Private Sub cmdSearch_Click()
'
'  On Error Resume Next
'
'    Dim Conn As New Connection
'    Dim cmd As New Command
'    Dim Rs As New ADODB.Recordset
'
'
'
'    Select Case Opt_Index
'
'         Case 0, 1, 3, 8, 11
'
'            Search_Pattern = Trim(txtSearch)
'
'        Case 2, 9
'
'            Src_Dt1 = Valid_Dt(dtpDtFrom.Value)
'            Src_Dt2 = Valid_Dt(dtpDtTo.Value)
'
'        Case 4, 6, 7, 10, 12, 13
'
'            Search_Pattern = Trim(cboSearch)
'
'
'End Select
''-------------------------------------------------'
'
'    If Search_Pattern = Empty Then
'      'Msgbox "Please "
'        Exit Sub
'    End If
'
'
'  'MsgBox Opt_Index
'  'MsgBox Src_Dt1 + "   " + Src_Dt2
'
'    Conn.Open strCN.Connection_String
'
'    Set cmd.ActiveConnection = Conn
'    cmd.CommandType = adCmdText
'    cmd.CommandText = "GetSearch_Result " & Opt_Index & ",'" + Search_Pattern + "','" + Src_Dt1 + "','" + Src_Dt2 + "'"
'    Rs.CursorLocation = adUseClient
'    Rs.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
'
'    'If Not (Rs.EOF Or Rs.BOF) Then
'        Set Search_Rs = Rs
'    'End If
'
'    Call Show_Data
'
'    Search_Pattern = Empty
'
'End Sub
'
'Private Sub cmdExport_Click()
'
'    Dim strFile_Nm As String
'
'    With Form1.CommonDialog1
'            .Filter = "MS Excel Files (*.xls)|*.xls"    ' Set filters.
'            .DialogTitle = "Save Result As" + Space(65) + "Carew & Company (Bangladesh) Limited"
'            .filename = "Search result (" + Replace(lblSearchOption, "(more/equal)", "") + " wise)"
'            .Action = 2                                 ' Display the save as dialog box.
'       strFile_Nm = .filename                           ' assign file name
'    End With
'
'
'    'MsgBox strFile_Nm
'
'    'If dialog box is closed by pressing 'X' button then an error
'    'likely to be occured. To prevent this error ":\" checking is implied
'
'    If Not InStr(1, strFile_Nm, ":\") = 0 Then
'
'
'    ' MsgBox strFile_Nm
'
'    ' checking records in the recordset
'
'            If Search_Rs.RecordCount >= 1 Then
'                'calling subroutine to export data to Excel worksheet
'                Call SaveAsExcel(Search_Rs, strFile_Nm, xlWorkbookNormal, True)
'            End If
'
'    End If
'
'End Sub
'
'Private Sub Form_Initialize()
'On Error Resume Next
'    Dim i As Integer
'    Opt_Index = 0
'    StatusBar1.Panels(1).Width = 3150
'    StatusBar1.Panels(2).Width = 3800
'    StatusBar1.Panels(3).Width = 4900
'
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
''MsgBox KeyCode
'
' If KeyCode = vbKeyEscape Then Unload Me
'
'    'if key is pressed along  with control key
'
'    If (Shift And vbCtrlMask) > 0 Then
'
'        Select Case KeyCode
'
'            Case vbKeyN: optSearch(0).Value = True
'            Case vbKeyH: optSearch(1).Value = True
'            Case vbKeyJ: optSearch(2).Value = True
'            Case vbKeyG: optSearch(3).Value = True
'            Case vbKeyU: optSearch(4).Value = True
''           Case vbKeyC: optSearch(5).Value = True
'            Case vbKeyD: optSearch(6).Value = True
'            Case vbKeyT: optSearch(7).Value = True
'            Case vbKeyR: optSearch(8).Value = True
'            Case vbKeyE: optSearch(9).Value = True
'            Case vbKeyO: optSearch(10).Value = True
'            Case vbKeyB: optSearch(11).Value = True
'            Case vbKeyP: optSearch(12).Value = True
'            Case vbKeyM: optSearch(13).Value = True
'            Case vbKeyLeft:  cmdResize_Click (0)
'            Case vbKeyRight:  cmdResize_Click (1)
'            Case vbKeyS: cmdSearch_Click
'
'
'
'        End Select
'    End If
'
'    'if key is pressed along  with 'Alt' key
'
'    If (Shift And vbAltMask) > 0 Then
'
'        Select Case KeyCode
'
'            Case vbKeyX
'                If cmdExport.Enabled = True Then
'                        cmdExport_Click
'                End If
'        End Select
'    End If
'End Sub
'
'Private Sub Form_Load()
'    Screen_Position Me
'
'
'End Sub
'
'Private Sub optSearch_Click(index As Integer)
'   On Error Resume Next
'
'    Dim i As Integer
'
'    txtSearch = ""
'
'    Opt_Index = index
'
'    For i = 0 To optSearch.UBound
'        optSearch(i).ForeColor = &H800000
'    Next
'
'    optSearch(index).ForeColor = &HC000C0             '&HFF0000
'
'    lblSearchOption = Replace(optSearch(index).Caption, "&", "")
'
'
'
'    Select Case index
'
'
'        Case 4                  'Unit
'            Load_UnitNm Me
'
'        Case 6                  'Designation
'
'            Load_Desig Me
'        Case 7                  'Job Type
'
'            Load_JbType Me
'        Case 10                 'Accomodation Type
'
'            With cboSearch
'                .Clear
'                .AddItem "Proper"
'                .AddItem "Improvised"
'                .AddItem "Single"
'                .AddItem "Below Standard"
'            End With
'
'        Case 12                 'Pay Scale
'
'            Load_PScale Me
'
'        Case 13                 'Payment Mode
'
'            With cboSearch
'                .Clear
'                .AddItem "Cash"
'                .AddItem "Bank"
'                .AddItem "Draft"
'            End With
'
'
'    End Select
'
'        Call Screen_Rearrange(index)
'
'
'End Sub
'
'Public Sub Screen_Rearrange(index As Integer)
'
'
'    Select Case index
'
'         Case 0, 1, 3, 8, 11
'
'            With txtSearch
'                .Height = 285
'                .Width = 2850
'                .Top = 5400
'                .Left = 180
'                .Visible = True
'                .SetFocus
'            End With
'
'
'            dtpDtFrom.Visible = False
'            dtpDtTo.Visible = False
'            lblFrom_To.Visible = False
'            cboSearch.Visible = False
'        Case 2, 9
'            txtSearch.Visible = False
'            cboSearch.Visible = False
'            lblFrom_To.Visible = True
'
'            With dtpDtTo
'                .Height = 300
'                .Width = 1095
'                .Top = 5400
'                .Left = 1935
'                .Value = Date + 1
'                .Visible = True
'            End With
'            With dtpDtFrom
'                .Height = 300
'                .Width = 1095
'                .Top = 5400
'                .Left = 585
'                .Value = Date
'                .Visible = True
'                .SetFocus
'            End With
'
'        Case 4, 6, 7, 10, 12, 13
'
'            txtSearch.Visible = False
'            dtpDtFrom.Visible = False
'            dtpDtTo.Visible = False
'            lblFrom_To.Visible = False
'
'            With cboSearch
'                .Height = 285
'                .Width = 2850
'                .Top = 5400
'                .Left = 180
'                .Visible = True
'                .SetFocus
'            End With
'    End Select
'
'End Sub
'
'Public Sub Show_Data()
'
'    Dim Panel_Des As String
'    Dim i As Integer
'
'
'    If Not (Search_Rs.EOF Or Search_Rs.BOF) Then
'        Set dtgResult.DataSource = Search_Rs
'
'        'grid columns resizing as per field length
'
'        For i = 0 To Search_Rs.Fields.Count - 1
'              dtgResult.Columns(i).Width = Ln * Len(Search_Rs.Fields(i)) + Fine_Tune
'              dtgResult.Columns(i).Locked = True
'        Next
'
'        With StatusBar1
'                .Panels(1).Text = "Search option: " + lblSearchOption
'                .Panels(2).Text = "Search pattern  '" + Search_Pattern + "'"
'                .Panels(3).Text = Search_Rs.RecordCount & " record(s) found"
'                cmdExport.Enabled = True
'        End With
'
'    Else
'
'
'        For i = 0 To 4
'              dtgResult.Columns(i).Width = 1605
'        Next
'
'        Set dtgResult.DataSource = Nothing
'
'        With StatusBar1
'                .Panels(1).Text = "Search option: " + lblSearchOption
'                .Panels(2).Text = "Search pattern  '" + Search_Pattern + "'"
'                .Panels(3).Text = "Search is complete. There is no result to display"
'                 cmdExport.Enabled = False
'        End With
'    End If
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    Destroy Me
'End Sub
'
'
'Private Sub txtSearch_KeyPress(KeyAscii As MSForms.ReturnInteger)
'
'     Select Case Opt_Index
'        Case 3, 8, 11
'            KeyAscii = IsNum(KeyAscii)         'Accept numeric only
'     End Select
'End Sub
