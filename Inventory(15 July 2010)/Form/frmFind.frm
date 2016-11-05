VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Application User"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   75
      TabIndex        =   1
      Top             =   300
      Width           =   3915
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2415
      Left            =   75
      TabIndex        =   2
      Top             =   675
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
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
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Find"
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   300
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objcom As New DSLComFram.clsCommon

Public objFindRS As New ADODB.Recordset
Public strfrmCaption As String
Public intInputsel As Integer
'Public intInputsel1 As Integer
'Public intInputsel2 As Integer
'Public intInputsel3 As Integer
'Public intInputsel4 As Integer
'Public intInputsel5 As Integer

Public OwnerForm As Form

Private Sub DataGrid1_DblClick()
OwnerForm.txtFields(intInputsel) = DataGrid1.Columns(0).Text
'If Not IsEmpty(intInputsel1) Then OwnerForm.txtFields(intInputsel1) = DataGrid1.Columns(1).Text
'If Not IsEmpty(intInputsel2) Then OwnerForm.txtFields(intInputsel2) = DataGrid1.Columns(2).Text
'If Not IsEmpty(intInputsel3) Then OwnerForm.txtFields(intInputsel1) = DataGrid1.Columns(3).Text
'If Not IsEmpty(intInputsel4) Then OwnerForm.txtFields(intInputsel2) = DataGrid1.Columns(4).Text
'If Not IsEmpty(intInputsel5) Then OwnerForm.txtFields(intInputsel1) = DataGrid1.Columns(5).Text


Unload Me
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    DataGrid1_DblClick
End If
End Sub

Private Sub Form_Load()
Set DataGrid1.DataSource = objFindRS
DataGrid1.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmFind = Nothing
End Sub

Private Sub txtFind_Change()
Dim objRs As New ADODB.Recordset
Set objRs = objFindRS 'objcom.Get_RS("select CustomerId,CustomerName,Address,Phone,Fax,City,Country from CustomerInfo where CustomerId='" & txtFields(0).Text & "'", objmyCon)
                                
Set DataGrid1.DataSource = objRs
DataGrid1.Refresh
Set objRs = Nothing

End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    DataGrid1.SetFocus
End If
End Sub
