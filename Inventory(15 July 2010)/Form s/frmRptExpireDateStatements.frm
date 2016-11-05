VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptExpireDateStatements 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Expired Date Statements"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   -90
      TabIndex        =   6
      Top             =   -60
      Width           =   6225
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expire Date Statement"
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
         Left            =   990
         TabIndex        =   7
         Top             =   210
         Width           =   3705
      End
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
      ItemData        =   "frmRptExpireDateStatements.frx":0000
      Left            =   1380
      List            =   "frmRptExpireDateStatements.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1500
      Width           =   4005
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
      Left            =   4485
      TabIndex        =   3
      ToolTipText     =   "Click to Save"
      Top             =   2400
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
      Left            =   3480
      TabIndex        =   2
      ToolTipText     =   "Click to View Report"
      Top             =   2400
      Width           =   945
   End
   Begin MSMask.MaskEdBox DemandDate 
      Height          =   375
      Index           =   0
      Left            =   1380
      TabIndex        =   0
      Top             =   885
      Width           =   1965
      _ExtentX        =   3466
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
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      TabIndex        =   8
      Top             =   2100
      Width           =   6105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Date Upto :"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   930
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Category Name"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1650
      Width           =   1095
   End
End
Attribute VB_Name = "frmRptExpireDateStatements"
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
If Combo1 = "All Category" Then
    CatCode = "All"
    CatName = Combo1
Else
    CatCode = Get_Code(Combo1)
    CatName = Get_Description(Combo1)
End If

rptmode = 6
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
DemandDate(0).Text = Format(Date, "dd/mm/yy")
End Sub

