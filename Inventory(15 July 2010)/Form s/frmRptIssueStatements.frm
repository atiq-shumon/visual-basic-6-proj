VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptIssueStatements 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Issue Statements"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7575
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Statements"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1680
         TabIndex        =   11
         Top             =   180
         Width           =   2865
      End
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmRptIssueStatements.frx":0000
      Left            =   1350
      List            =   "frmRptIssueStatements.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2280
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmRptIssueStatements.frx":0015
      Left            =   1350
      List            =   "frmRptIssueStatements.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1612
      Width           =   4245
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4875
      TabIndex        =   5
      ToolTipText     =   "Click to Save"
      Top             =   3555
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H8000000C&
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3645
      TabIndex        =   4
      ToolTipText     =   "Click to View Report"
      Top             =   3555
      Width           =   1215
   End
   Begin MSMask.MaskEdBox DemandDate 
      Height          =   405
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   975
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox DemandDate 
      Height          =   405
      Index           =   1
      Left            =   3780
      TabIndex        =   1
      Top             =   975
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      TabIndex        =   9
      Top             =   3420
      Width           =   7575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Issye Type"
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   2378
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Demand Period"
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Category Name"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   1710
      Width           =   1095
   End
End
Attribute VB_Name = "frmRptIssueStatements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objcom As New DSLComFram.clsCommon
Dim objRs As New ADODB.Recordset

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
If Combo1 = "" Then
    MsgBox "Category Required", vbInformation, App.title
    Combo1.SetFocus
    Exit Sub
End If
If Combo2 = "" Then
    MsgBox "Type Required", vbInformation, App.title
    Combo2.SetFocus
    Exit Sub
End If
StDate = Format(DemandDate(0).Text, "dd mmm yyyy")
EdDate = Format(DemandDate(1).Text, "dd mmm yyyy")
If Combo1 = "All Category" Then
    CatCode = "All"
    CatName = Combo1
Else
    CatCode = Get_Code(Combo1)
    CatName = Get_Description(Combo1)
End If

If Combo2 = "All Type" Then
    CustCode = "All"
    CustName = Combo2
    If CustName = "All Type" Then
       CustName = "All"
    End If
Else
    CustCode = Get_Code(Combo2)
    CustName = Get_Description(Combo2)
End If

rptmode = 5
rptViewer.Show 1
End Sub

Private Sub DemandDate_GotFocus(Index As Integer)
DemandDate(Index).SelLength = Len(DemandDate(Index))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     SendKeys (Chr(9))
  End If
End Sub

Private Sub Form_Load()
Set objRs = objcom.Get_RS("SELECT cate_code,cate_name from item_cate_info", objmyCon)
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       Combo1.AddItem objRs(1) + "~" + objRs(0)
       objRs.MoveNext
    Loop
End If

Set objRs = objcom.Get_RS("SELECT type_code,type_name from item_issue_type", objmyCon)
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       Combo2.AddItem objRs(1) + "~" + objRs(0)
       objRs.MoveNext
    Loop
 End If
 
DemandDate(0).Text = Format(Date, "dd/mm/yy")
DemandDate(1).Text = Format(Date, "dd/mm/yy")
End Sub

