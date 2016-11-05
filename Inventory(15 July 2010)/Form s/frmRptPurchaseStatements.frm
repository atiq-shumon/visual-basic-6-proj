VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptPurchaseStatements 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Purchase Statements"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   -30
      TabIndex        =   9
      Top             =   -30
      Width           =   8085
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Statement"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1980
         TabIndex        =   10
         Top             =   270
         Width           =   3570
      End
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "frmRptPurchaseStatements.frx":0000
      Left            =   1560
      List            =   "frmRptPurchaseStatements.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2670
      Width           =   4305
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "frmRptPurchaseStatements.frx":0019
      Left            =   1560
      List            =   "frmRptPurchaseStatements.frx":0020
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2040
      Width           =   4335
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5130
      TabIndex        =   5
      ToolTipText     =   "Click to Close"
      Top             =   3630
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H8000000C&
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3855
      TabIndex        =   4
      ToolTipText     =   "Click to View Report"
      Top             =   3630
      Width           =   1245
   End
   Begin MSMask.MaskEdBox DemandDate 
      Height          =   465
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   1230
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   820
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
      Height          =   465
      Index           =   1
      Left            =   4080
      TabIndex        =   1
      Top             =   1245
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   820
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
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   -60
      TabIndex        =   11
      Top             =   3450
      Width           =   8175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3540
      TabIndex        =   12
      Top             =   1350
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Supplier Name"
      Height          =   195
      Left            =   30
      TabIndex        =   8
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Demand Period"
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   1290
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Category Name"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   2100
      Width           =   1095
   End
End
Attribute VB_Name = "frmRptPurchaseStatements"
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
    MsgBox "Supplier Required", vbInformation, App.title
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

If Combo2 = "All Supplier" Then
    SuppCode = "All"
    SuppName = Combo2
Else
    SuppCode = Get_Code(Combo2)
    SuppName = Get_Description(Combo2)
End If

rptmode = 4
rptViewer.Show 1
End Sub

Private Sub DemandDate_GotFocus(Index As Integer)
DemandDate(Index).SelLength = Len(DemandDate(Index))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
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
DemandDate(0).Text = Format(Date, "dd/mm/yy")
DemandDate(1).Text = Format(Date, "dd/mm/yy")
End Sub

