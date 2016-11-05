VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptItemLedger 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Item Ledger"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      TabIndex        =   8
      Top             =   -30
      Width           =   7575
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Ledger "
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
         Left            =   2460
         TabIndex        =   9
         Top             =   180
         Width           =   2085
      End
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
      ItemData        =   "frmRptItemLedger.frx":0000
      Left            =   1350
      List            =   "frmRptItemLedger.frx":0002
      TabIndex        =   2
      Top             =   1740
      Width           =   5325
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
      Height          =   495
      Left            =   5805
      TabIndex        =   4
      ToolTipText     =   "Click to Save"
      Top             =   3090
      Width           =   1125
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
      Height          =   495
      Left            =   4620
      TabIndex        =   3
      ToolTipText     =   "Click to View Report"
      Top             =   3090
      Width           =   1125
   End
   Begin MSMask.MaskEdBox DemandDate 
      Height          =   375
      Index           =   0
      Left            =   1350
      TabIndex        =   0
      Top             =   1020
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
      Height          =   375
      Index           =   1
      Left            =   3930
      TabIndex        =   1
      Top             =   1020
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      TabIndex        =   10
      Top             =   2970
      Width           =   7425
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3330
      TabIndex        =   7
      Top             =   1080
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Demand Period"
      Height          =   225
      Left            =   90
      TabIndex        =   6
      Top             =   1050
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Item Name"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1860
      Width           =   765
   End
End
Attribute VB_Name = "frmRptItemLedger"
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

StDate = Format(DemandDate(0).Text, "dd mmm yyyy")
EdDate = Format(DemandDate(1).Text, "dd mmm yyyy")
If Combo1 = "All Category" Then
    CatCode = "All"
    CatName = Combo1
Else
    CatCode = Get_Code(Combo1)
    CatName = Get_Description(Combo1)
End If

rptmode = 9
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
Set objRs = objcom.Get_RS("SELECT item_code,item_name from item_info where cate_code='" & CategoryCode & "'", objmyCon)
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

