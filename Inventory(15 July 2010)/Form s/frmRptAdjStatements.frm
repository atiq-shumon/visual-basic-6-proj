VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptAdjStatements 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Adjustment Statements"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmRptAdjStatements.frx":0000
      Left            =   1350
      List            =   "frmRptAdjStatements.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   3075
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
      Height          =   375
      Left            =   1035
      TabIndex        =   1
      ToolTipText     =   "Click to Save"
      Top             =   1215
      Width           =   945
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
      Height          =   375
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Click to View Report"
      Top             =   1215
      Width           =   945
   End
   Begin MSMask.MaskEdBox DemandDate 
      Height          =   285
      Index           =   0
      Left            =   1350
      TabIndex        =   5
      Top             =   135
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      Format          =   "dd-mmm-yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox DemandDate 
      Height          =   285
      Index           =   1
      Left            =   3090
      TabIndex        =   6
      Top             =   135
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      Format          =   "dd-mmm-yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Demand Period"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Category Name"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   540
      Width           =   1095
   End
End
Attribute VB_Name = "frmRptAdjStatements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objcom As New DSLComFram.clsCommon
Dim objrs As New ADODB.Recordset

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
If Combo1 = "" Then
    MsgBox "Category Required", vbInformation, App.title
    Combo1.SetFocus
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

rptmode = 8
rptViewer.Show 1
End Sub

Private Sub DemandDate_GotFocus(Index As Integer)
DemandDate(Index).SelLength = Len(DemandDate(Index))
End Sub

Private Sub Form_Load()
Set objrs = objcom.Get_RS("SELECT cate_code,cate_name from item_cate_info", objmyCon)
 If Not objrs.EOF Then
    objrs.MoveFirst
    Do Until objrs.EOF
       Combo1.AddItem objrs(1) + "~" + objrs(0)
       objrs.MoveNext
    Loop
End If


 
DemandDate(0).Text = Format(Date, "dd/mm/yy")
DemandDate(1).Text = Format(Date, "dd/mm/yy")
End Sub

