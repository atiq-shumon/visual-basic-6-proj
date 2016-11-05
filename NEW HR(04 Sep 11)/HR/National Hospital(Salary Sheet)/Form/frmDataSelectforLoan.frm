VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataSelectforLoan 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3615
   Icon            =   "frmDataSelectforLoan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleMode       =   0  'User
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   3500
      _ExtentX        =   6165
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      ForeColor       =   0
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
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   600
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelectforLoan.frx":030A
            Key             =   "Bell"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelectforLoan.frx":04A4
            Key             =   "Sort Ascending"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelectforLoan.frx":05B6
            Key             =   "Misc08"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelectforLoan.frx":08D0
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelectforLoan.frx":09E2
            Key             =   "Top"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelectforLoan.frx":0F24
            Key             =   "Prior"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelectforLoan.frx":102E
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelectforLoan.frx":1140
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelectforLoan.frx":124A
            Key             =   "Bottom"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelectforLoan.frx":178C
            Key             =   "Spell Check"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelectforLoan.frx":189E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelectforLoan.frx":1CF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelectforLoan.frx":200A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      HotImageList    =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Select"
            Object.ToolTipText     =   "Select"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Top"
            Object.ToolTipText     =   "First Record"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Prior"
            Object.ToolTipText     =   "Previous Record"
            ImageKey        =   "Prior"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next"
            Object.ToolTipText     =   "Next Record"
            ImageKey        =   "Next"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bottom"
            Object.ToolTipText     =   "Last Record"
            ImageKey        =   "Bottom"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmDataSelectforLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intPutSel, intputsel1, intputsel2, colval As Integer
Public adoRecordset As New Recordset
Public OwnerForm As Form
'Public intCaption As Integer
Dim frmRecordset As New Recordset
Dim mbCtrlKey As Integer
Dim mbSortCol As String
Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
On Error GoTo Errdes
    If mbCtrlKey Then
        msSortCol = "[" & adoRecordset(ColIndex).Name & "] desc"
        mbCtrlKey = 0 'reset it
    Else
        msSortCol = "[" & adoRecordset(ColIndex).Name & "]"
    End If
    adoRecordset.Sort = msSortCol
    msSortCol = vbNullString 'reset it
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub Form_Load()
On Error GoTo Errdes
    Set frmRecordset = adoRecordset
    Set grdDataGrid.DataSource = frmRecordset
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> 1 Then
        grdDataGrid.Height = Me.Height - (950)
        grdDataGrid.Width = Me.Width - (250)
    End If
End Sub

Private Sub grdDataGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        DataSelect
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Select"
            DataSelect
        Case "Bottom"
            LastRecord
        Case "Next"
            NextRecord
        Case "Find"
            DataSearch
        Case "Prior"
            PreviousRecord
        Case "Top"
            FirstRecord
   End Select
End Sub

Public Sub FirstRecord()
On Error GoTo Errdes
    If adoRecordset.RecordCount <> 0 Then
        adoRecordset.MoveFirst
    End If
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Public Sub LastRecord()
On Error GoTo Errdes
    If adoRecordset.RecordCount <> 0 Then
        adoRecordset.MoveLast
    End If
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Public Sub NextRecord()
On Error GoTo Errdes
    If adoRecordset.RecordCount <> 0 Then
        If adoRecordset.EOF Then
            adoRecordset.MoveLast
        Else
            adoRecordset.MoveNext
        End If
    End If
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Public Sub PreviousRecord()
On Error GoTo Errdes
    If adoRecordset.RecordCount <> 0 Then
        If adoRecordset.BOF Then
            adoRecordset.MoveFirst
        Else
            adoRecordset.MovePrevious
        End If
    End If
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Public Sub DataSearch()
On Error GoTo SearchError
Dim txtSearch As String
    If adoRecordset.RecordCount <> 0 Then
        txtSearch = InputBox("Enter Data Search Expression", "Data Search")
        If Len(Trim(txtSearch)) = 0 Then Exit Sub
        adoRecordset.MoveFirst
        adoRecordset.Find txtSearch
        If (adoRecordset.EOF Or adoRecordset.BOF) Then
            MsgBox "Search Fail"
        End If
    End If
    Exit Sub
SearchError:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Public Sub DataSelect()
On Error GoTo DataSelectError
    Dim varData As Variant
    Dim showtexData(5) As Variant
    
    If (adoRecordset.RecordCount <> 0) Or (Not (adoRecordset.EOF Or adoRecordset.BOF)) Then
        varData = adoRecordset(0)
        OwnerForm.Combo1(intPutSel) = varData
    
       If colval = 2 Then
        showtexData(0) = adoRecordset(1)
        ElseIf colval = 3 Then
        showtexData(0) = adoRecordset(1)
        showtexData(1) = adoRecordset(2)
       End If
        If showtexData(0) <> "" Then
        OwnerForm.Txtshow(intputsel1) = showtexData(0)
       End If
        
        If showtexData(1) <> "" Then
         OwnerForm.Txtshow(intputsel2) = showtexData(1)
        End If
        
        Unload Me
        OwnerForm.Combo1(intPutSel).SetFocus
    End If
    colval = 0
    Exit Sub
DataSelectError:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
