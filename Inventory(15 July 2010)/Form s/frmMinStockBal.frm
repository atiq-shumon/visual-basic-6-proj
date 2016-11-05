VERSION 5.00
Begin VB.Form frmMinStockBal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Low Stock Items Report"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      ItemData        =   "frmMinStockBal.frx":0000
      Left            =   1500
      List            =   "frmMinStockBal.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1275
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
      Height          =   435
      Left            =   5130
      TabIndex        =   1
      ToolTipText     =   "Click to Close"
      Top             =   2385
      Width           =   1005
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
      Height          =   435
      Left            =   4125
      TabIndex        =   0
      ToolTipText     =   "Click to View Report"
      Top             =   2385
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   -60
      TabIndex        =   4
      Top             =   -30
      Width           =   6705
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Low Stock items statements"
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
         Left            =   810
         TabIndex        =   5
         Top             =   180
         Width           =   4605
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      TabIndex        =   6
      Top             =   2130
      Width           =   6735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Category Name"
      Height          =   195
      Left            =   165
      TabIndex        =   3
      Top             =   1350
      Width           =   1095
   End
End
Attribute VB_Name = "frmMinStockBal"
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
If Combo1.Text = "All Category" Then
    CatCode = "All"
    CatName = Combo1
    rptmode = 10
    rptViewer.Show 1
Else
    CatCode = Get_Code(Combo1)
    CatName = Get_Description(Combo1)
    rptmode = 10
    rptViewer.Show 1
End If
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
End Sub
