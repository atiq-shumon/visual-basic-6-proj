VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptStockStatements 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stock/Valuation Statements"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   -150
      TabIndex        =   15
      Top             =   -60
      Width           =   8115
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock/Value Statement"
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
         Left            =   2010
         TabIndex        =   16
         Top             =   300
         Width           =   3840
      End
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      Caption         =   "Stock Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   1
      Left            =   3960
      TabIndex        =   13
      Top             =   1020
      Width           =   2325
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      Caption         =   "Stock(Qty) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   0
      Left            =   1260
      TabIndex        =   12
      Top             =   1020
      Value           =   -1  'True
      Width           =   1965
   End
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "frmRptStockStatements.frx":0000
      Left            =   1530
      List            =   "frmRptStockStatements.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3465
      Width           =   4965
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
      ItemData        =   "frmRptStockStatements.frx":0019
      Left            =   1530
      List            =   "frmRptStockStatements.frx":0020
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2865
      Width           =   4965
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
      ItemData        =   "frmRptStockStatements.frx":0032
      Left            =   1530
      List            =   "frmRptStockStatements.frx":0039
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2280
      Width           =   4965
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6015
      TabIndex        =   6
      ToolTipText     =   "Click to Save"
      Top             =   4245
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H8000000C&
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4695
      TabIndex        =   5
      ToolTipText     =   "Click to View Report"
      Top             =   4245
      Width           =   1275
   End
   Begin MSMask.MaskEdBox DemandDate 
      Height          =   405
      Index           =   0
      Left            =   1530
      TabIndex        =   0
      Top             =   1635
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
   Begin MSMask.MaskEdBox DemandDate 
      Height          =   405
      Index           =   1
      Left            =   4680
      TabIndex        =   1
      Top             =   1635
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   -120
      TabIndex        =   14
      Top             =   4080
      Width           =   8085
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      Height          =   555
      Left            =   -60
      Top             =   930
      Width           =   8115
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Demand Period"
      Height          =   195
      Left            =   300
      TabIndex        =   11
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3750
      TabIndex        =   10
      Top             =   1665
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Group Name"
      Height          =   195
      Left            =   270
      TabIndex        =   9
      Top             =   3570
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Type Name"
      Height          =   195
      Left            =   300
      TabIndex        =   8
      Top             =   2970
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Category Name"
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "frmRptStockStatements"
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
If Option1(0).value = True Then
    stockOrValue = 1
ElseIf Option1(1).value = True Then
    stockOrValue = 2
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
If stkFormMOde = 1 Then
    rptmode = 7
    rptViewer.Show 1
ElseIf stkFormMOde = 2 Then
   rptmode = 14
    rptViewer.Show 1
End If
End Sub

Private Sub Combo1_Click()
  load_data (1)
End Sub

Private Sub Combo2_Click()
   load_data (2)
End Sub

Private Sub DemandDate_GotFocus(Index As Integer)
   DemandDate(Index).SelStart = 0
   DemandDate(Index).SelLength = Len(DemandDate(Index))
End Sub
Private Sub load_data(mode As Integer)
Dim combo_rs1 As New ADODB.Recordset
Dim combo_rs2 As New ADODB.Recordset
 Select Case mode
        Case 1 ''''''''''type
           Combo2.Clear
           Combo3.Clear
           Set combo_rs1 = objcom.Get_RS("SELECT type_code,type_name from item_type_info where cate_code='" & Get_Code(Combo1) & "'", objmyCon)
           If Not combo_rs1.EOF Then
             combo_rs1.MoveFirst
             Do Until combo_rs1.EOF
                Combo2.AddItem combo_rs1(1) + "~" + combo_rs1(0)
                combo_rs1.MoveNext
             Loop
         End If
      Case 2
           Combo3.Clear
           Set combo_rs2 = objcom.Get_RS("SELECT group_code,group_name from item_group_info where  type_code='" & Get_Code(Combo2.Text) & "'", objmyCon)
          If Not combo_rs2.EOF Then
             combo_rs2.MoveFirst
             Do Until combo_rs2.EOF
                Combo3.AddItem combo_rs2(1) + "~" + combo_rs2(0)
                combo_rs2.MoveNext
             Loop
         End If

End Select
Set combo_rs1 = Nothing
Set combo_rs2 = Nothing
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
    Unload Me
 End If
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
DemandDate(1).Text = Format(Date, "dd/mm/yy")
End Sub

