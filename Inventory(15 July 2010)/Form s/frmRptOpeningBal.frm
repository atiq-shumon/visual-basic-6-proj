VERSION 5.00
Begin VB.Form frmRptOpeningBal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Opening Balance"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   -180
      TabIndex        =   5
      Top             =   -60
      Width           =   6105
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Balance Statement"
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
         TabIndex        =   6
         Top             =   180
         Width           =   4545
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmRptOpeningBal.frx":0000
      Left            =   1740
      List            =   "frmRptOpeningBal.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   2805
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
      Left            =   4500
      TabIndex        =   1
      ToolTipText     =   "Click to Save"
      Top             =   2535
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
      Left            =   3525
      TabIndex        =   0
      ToolTipText     =   "Click to View Report"
      Top             =   2535
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   30
      TabIndex        =   4
      Top             =   2400
      Width           =   5715
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Category Name"
      Height          =   195
      Left            =   300
      TabIndex        =   3
      Top             =   1290
      Width           =   1095
   End
End
Attribute VB_Name = "frmRptOpeningBal"
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
    rptmode = 2
    rptViewer.Show 1
Else
    CatCode = Get_Code(Combo1)
    CatName = Get_Description(Combo1)
    rptmode = 3
    rptViewer.Show 1
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
