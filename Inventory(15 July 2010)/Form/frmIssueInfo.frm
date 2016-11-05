VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmIssue 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   2520
      Left            =   30
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   4445
      _Version        =   393216
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14737632
      ForeColor       =   8388608
      BackColorFixed  =   14737632
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483635
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtfields 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   7470
      MaxLength       =   15
      TabIndex        =   12
      Text            =   "0"
      ToolTipText     =   "Insert Quantity"
      Top             =   2670
      Width           =   825
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   1695
      Left            =   -30
      TabIndex        =   32
      Top             =   720
      Width           =   11955
      Begin VB.CommandButton Command1 
         Caption         =   ":::"
         Height          =   285
         Index           =   2
         Left            =   11520
         TabIndex        =   48
         ToolTipText     =   "Search Employee"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   8
         Left            =   5460
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   47
         ToolTipText     =   "Insert Indent No"
         Top             =   600
         Width           =   6045
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   4380
         MaxLength       =   30
         TabIndex        =   6
         ToolTipText     =   "Insert Issue To"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   10410
         MaxLength       =   30
         TabIndex        =   4
         ToolTipText     =   "Insert Reg. No"
         Top             =   165
         Width           =   1455
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   6510
         MaxLength       =   30
         TabIndex        =   2
         ToolTipText     =   "Insert Indent No"
         Top             =   165
         Width           =   1155
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1140
         MaxLength       =   12
         TabIndex        =   0
         ToolTipText     =   "Purchase Serial"
         Top             =   165
         Width           =   1875
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   2
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   7
         ToolTipText     =   "Insert Remarks"
         Top             =   1095
         Width           =   8505
      End
      Begin VB.CommandButton Command1 
         Caption         =   ":::"
         Height          =   315
         Index           =   0
         Left            =   3030
         TabIndex        =   33
         Top             =   150
         Width           =   375
      End
      Begin VB.ComboBox CboPurType 
         Height          =   315
         ItemData        =   "frmIssueInfo.frx":0000
         Left            =   1140
         List            =   "frmIssueInfo.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   2265
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   285
         Left            =   4380
         TabIndex        =   1
         Top             =   165
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin LVbuttons.LaVolpeButton CmdGenerate 
         Height          =   435
         Left            =   9660
         TabIndex        =   8
         ToolTipText     =   "Press to Generate Purchase Serial"
         Top             =   1080
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "Generate Serial"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   14215660
         FCOL            =   192
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmIssueInfo.frx":0030
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   285
         Left            =   8700
         TabIndex        =   3
         Top             =   165
         Width           =   1095
         _ExtentX        =   1931
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
         BackStyle       =   0  'Transparent
         Caption         =   "Reg #"
         Height          =   195
         Index           =   6
         Left            =   9840
         TabIndex        =   46
         Top             =   195
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Indent Date "
         Height          =   195
         Index           =   5
         Left            =   7770
         TabIndex        =   45
         Top             =   180
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Indent No #"
         Height          =   195
         Index           =   4
         Left            =   5640
         TabIndex        =   44
         Top             =   195
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue  Serial#"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   38
         Top             =   210
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   1140
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Date "
         Height          =   195
         Index           =   1
         Left            =   3510
         TabIndex        =   36
         Top             =   210
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Type"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   35
         Top             =   660
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue To "
         Height          =   195
         Index           =   3
         Left            =   3510
         TabIndex        =   34
         Top             =   630
         Width           =   660
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000B&
      Height          =   705
      Left            =   0
      TabIndex        =   26
      Top             =   2310
      Width           =   11925
      Begin VB.TextBox txtItemTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAF2C8&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   360
         Width           =   4635
      End
      Begin VB.TextBox cboItem 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   30
         TabIndex        =   9
         Top             =   360
         Width           =   1425
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0097C8E6&
         Caption         =   ">>"
         Height          =   345
         Index           =   1
         Left            =   11460
         MaskColor       =   &H00C0C000&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   405
      End
      Begin VB.ComboBox CboPurId 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6090
         TabIndex        =   11
         Top             =   360
         Width           =   1395
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   10530
         TabIndex        =   14
         Top             =   360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   315
         Index           =   3
         Left            =   9480
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   16
         Text            =   "0"
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   4
         Left            =   8310
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   15
         Text            =   "0"
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   2100
         TabIndex        =   42
         Top             =   120
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   41
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Id"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   6330
         TabIndex        =   39
         Top             =   120
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp. Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   10530
         TabIndex        =   30
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   9510
         TabIndex        =   29
         Top             =   120
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   7530
         TabIndex        =   28
         Top             =   120
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Rate"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   8340
         TabIndex        =   27
         Top             =   120
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   615
      Left            =   0
      TabIndex        =   23
      Top             =   7440
      Width           =   11925
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000C&
         Caption         =   "Print"
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
         Left            =   8940
         TabIndex        =   43
         ToolTipText     =   "Click to insert new information"
         Top             =   150
         Width           =   945
      End
      Begin VB.TextBox txttrackid 
         Height          =   285
         Left            =   1200
         TabIndex        =   31
         Top             =   180
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H8000000C&
         Caption         =   "Edit"
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
         Left            =   7980
         TabIndex        =   24
         ToolTipText     =   "Click to Edit Information"
         Top             =   150
         Width           =   945
      End
      Begin VB.CommandButton cmdnew 
         BackColor       =   &H8000000C&
         Caption         =   "New"
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
         Left            =   6030
         TabIndex        =   18
         ToolTipText     =   "Click to insert new information"
         Top             =   150
         Width           =   945
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H8000000C&
         Caption         =   "Save"
         Enabled         =   0   'False
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
         Left            =   7005
         TabIndex        =   17
         ToolTipText     =   "Click to Save"
         Top             =   150
         Width           =   945
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H8000000C&
         Caption         =   "Delete"
         Enabled         =   0   'False
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
         Left            =   9915
         TabIndex        =   19
         ToolTipText     =   "Click to Delete"
         Top             =   150
         Width           =   945
      End
      Begin VB.CommandButton cmdexit 
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
         Left            =   10890
         TabIndex        =   20
         ToolTipText     =   "Click to Close"
         Top             =   150
         Width           =   945
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   6960
         Top             =   120
         Width           =   4905
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4455
      Left            =   0
      TabIndex        =   22
      Top             =   3000
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   7858
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   13627123
      ForeColor       =   12582912
      ForeColorFixed  =   0
      BackColorSel    =   12640511
      ForeColorSel    =   8421631
      BackColorBkg    =   -2147483637
      GridColor       =   -2147483637
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      SelectionMode   =   1
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000001&
      FillColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   -30
      ScaleHeight     =   795
      ScaleWidth      =   11895
      TabIndex        =   21
      Top             =   -30
      Width           =   11955
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Issue Entry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEF0F7&
         Height          =   405
         Left            =   4860
         TabIndex        =   25
         Top             =   120
         Width           =   2850
      End
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mnuDL 
         Caption         =   "Delete"
      End
      Begin VB.Menu fdsafdsa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRF 
         Caption         =   "Refresh"
      End
      Begin VB.Menu gfdsgds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objcom As New DSLComFram.clsCommon
Dim objRs As New ADODB.Recordset
Dim newBalance As Double
Private Sub Cbocategory_Click()
    load_item
End Sub
Private Sub CboItem_Click()
 On Error Resume Next
    load_purid
End Sub
Private Sub load_qty()
  Set objRs = objcom.Get_RS("SELECT (PurQty-(UsedQty+ReturnQty)) as PurQty,URate,exp_date from PurchaseSub  WHERE (PurId= '" & Trim(CboPurId) & "' and to_number(itemId)=to_number('" & Trim(Get_Code(cboItem)) & "'))", objmyCon)
       If Not objRs.EOF Then
          txtfields(1) = objRs(0)
          txtfields(4) = objRs(1)
          If Not IsNull(objRs(2)) Then
            MaskEdBox1 = Format(objRs(2), "dd/mm/yy")
          Else
            MaskEdBox1.Text = "__/__/__"
          End If
      End If
End Sub
Private Sub getItemCode(title As String)
  On Error GoTo err_loop
    MSFlexGrid2.Clear
    MSFlexGrid2.Rows = 0
      
    MSFlexGrid2.ColWidth(0) = "600"
    MSFlexGrid2.ColAlignment(0) = 1
    MSFlexGrid2.ColWidth(1) = "8100"
    MSFlexGrid2.ColWidth(2) = "2800"
    
    Set objRs = objcom.Get_RS("select a.item_code,a.item_name,b.group_name from item_info a,item_group_info b where a.group_code=b.group_code and Upper(a.item_name) like '" & Trim(UCase(title)) & "%' AND a.cate_code='" & CategoryCode & "'", objmyCon)
    If Not objRs.EOF Then
    i = 0
    With MSFlexGrid2
        Do Until objRs.EOF
            MSFlexGrid2.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = Trim(objRs(0))
                .TextMatrix(i, 1) = Trim(objRs(1))
                .TextMatrix(i, 2) = objRs(2)
                i = i + 1
            objRs.MoveNext
        Loop
    End With
Else
    MSFlexGrid2.Rows = 50
 End If
    MSFlexGrid2.Visible = True
'    MSFlexGrid2.TabIndex = cboItem.TabIndex
    MSFlexGrid2.SetFocus
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

Private Sub cboItem_GotFocus()
  cboItem.SelStart = 0
  cboItem.SelLength = Len(cboItem)
End Sub

Private Sub CboItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If cboItem = "" Then Exit Sub
 
If IsNumeric(cboItem) = True Then
 Set objRs = objcom.Get_RS("select item_name from item_info where to_number(item_code)='" & Trim(cboItem.Text) & "'", objmyCon)
   If Not objRs.EOF Then
     txtItemTitle.Text = objRs(0)
     CboPurId.SetFocus
    End If
Else
  MSFlexGrid2.Left = cboItem.Left
  '       MSFlexGrid2.Top = cboItem.Top
 MSFlexGrid2.TabIndex = cboItem.TabIndex
  Call getItemCode(cboItem)
  Exit Sub
End If
End If
End Sub

Private Sub cboItem_LostFocus()
'     If Len(Trim(cboItem.Text)) = 0 Then
'        txtItemTitle = ""
'        Exit Sub
'     Else
'            If IsNumeric(cboItem) Then
'               Set objRs = objcom.Get_RS("select item_name from item_info where to_number(item_code)='" & Trim(cboItem.Text) & "'", objmyCon)
'                 If Not objRs.EOF Then
'                   txtItemTitle.Text = objRs(0)
'                 End If
'
'            Else
'
'               MSFlexGrid2.Left = cboItem.Left
'       '       MSFlexGrid2.Top = cboItem.Top
'               MSFlexGrid2.TabIndex = cboItem.TabIndex
'               Call getItemCode(cboItem)
'            Exit Sub
'          End If
'  End If
End Sub

Private Sub CboPurId_Click()
    load_qty
End Sub

Private Sub CboPurId_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      txtfields(1).SetFocus
  End If
End Sub

Private Sub CboPurType_Click()
      IssueSearchMode = CboPurType.ListIndex
End Sub
Private Sub CboPurType_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     txtfields(7).SetFocus
  End If
End Sub
Private Sub CboSupplier_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     txtfields(2).SetFocus
  End If
End Sub
Private Sub cmdDelete_Click()
           If Len(txtfields(0)) = 0 Then
                MsgBox "Issue Serial Mandatory", vbInformation, App.title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(cboItem) = 0 Then
                MsgBox "Please Select an Item", vbInformation, App.title
                cboItem.SetFocus
                Exit Sub
            End If
            
            If Len(CboPurId) = 0 Then
                MsgBox "Please Select an Purchase ID", vbInformation, App.title
                CboPurId.SetFocus
                Exit Sub
            End If
            
           
            
             If Val(txtfields(3)) = 0 Then
                MsgBox "Amount Must be More than One", vbInformation, App.title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
           
            Set objRs = objcom.Get_RS("SELECT IssueId from IssueMain  WHERE (IssueId= '" & txtfields(0) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "Invalid Issue Serial..Please Verify.", vbInformation, cmp
               txtfields(0).SelLength = Len(txtfields(0))
               txtfields(0).SetFocus
               Exit Sub
            End If
          If MsgBox("Are you sure to Delete the Whole Issue Serial Information ", vbYesNo + vbInformation, cmp) = vbYes Then
                  delete
          Else
            Exit Sub
          End If

       MsgBox "Deleted successfully.", vbInformation, cmp
       Call ShowFlexData
       cmdnew.SetFocus
End Sub
Private Sub CmdEdit_Click()
  If Len(txtfields(0)) = 0 Then
       MsgBox "Issue Serial Mandatory..Please Blank the field to Generate", vbInformation, cmp
       txtfields(0).SetFocus
       Exit Sub
     End If
     
     If MaskEdBox2.Text = "__/__/__" Then
        MsgBox "Issue Date Mandatory", vbInformation, cmp
        MaskEdBox2.SetFocus
        Exit Sub
    End If
     
     If Len(CboPurType) = 0 Then
        MsgBox "Issue Type Required..", vbInformation, cmp
        CboPurType.SetFocus
        Exit Sub
    End If
        
'    If Len(CboSupplier) = 0 Then
'        MsgBox "Issue to whom Required..", vbInformation, cmp
'        CboSupplier.SetFocus
'        Exit Sub
'    End If
    Set objRs = objcom.Get_RS("SELECT IssueId from IssueMain  WHERE (IssueId= '" & txtfields(0) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "Invalid Issue Serial..Please Verify.", vbInformation, cmp
               txtfields(0).SelLength = Len(txtfields(0))
               txtfields(0).SetFocus
               Exit Sub
   End If
            
   
   Set objRs = objcom.Get_RS("SELECT  max(PurDate) from purchasemain  WHERE PurId in(select PurchaseId from IssueSub where issueid='" & txtfields(0).Text & "')", objmyCon)
            
     If Not objRs.EOF Then
      If Not IsNull(objRs(0)) Then
        If CDate(Format(MaskEdBox2.Text, "dd mmm yyyy")) < CDate(Format(objRs(0), "dd mmm yyyy")) Then
            MsgBox "Issued Date Can't be Less than Purchase Date ( " & (objRs(0)) & "  ),Please Verify", vbInformation, cmp
             Exit Sub
           End If
        End If
      End If
     
   generate (2)
   cboItem.SetFocus
   MsgBox "Updated SuccessfullY", vbInformation, cmp

End Sub
Private Sub cmdExit_Click()
      Unload Me
End Sub
Private Sub CmdGenerate_Click()
    If Len(txtfields(0)) > 0 Then
       MsgBox "Serial No already exists..Please Blank the field to Generate", vbInformation, cmp
       txtfields(0).SetFocus
       Exit Sub
     End If
     
     If MaskEdBox2.Text = "__/__/__" Then
        MsgBox "Issue Date Mandatory", vbInformation, cmp
        MaskEdBox2.SetFocus
        Exit Sub
    End If
     
     If Len(CboPurType) = 0 Then
        MsgBox "Issue Type Required..", vbInformation, cmp
        CboPurType.SetFocus
        Exit Sub
    End If
        
    If Len(txtfields(5)) = 0 Then
        MsgBox "Indent No Required...Please verify.", vbInformation, cmp
        txtfields(5).SetFocus
        Exit Sub
    End If
     
     generate (1)
   cboItem.SetFocus
End Sub
Private Sub generate(mode As Integer)
    Dim RS As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    Dim Param10 As New Parameter
     Dim Param11 As New Parameter
    
    
    If mode = 1 Then
        Set objRs = objcom.Get_RS("select to_char(nvl(max(to_number(substr(issueId,6,6))),0)+1,'000000')  from IssueMain where  to_char(substr(issueId,3,2))='" & CategoryCode & "'", objmyCon)
         If Not objRs.EOF Then
            txtfields(0) = "I-" + CategoryCode + "-" + Trim(objRs(0))
         End If
    End If
    
    
    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, mode)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 12, Trim(txtfields(0)))
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, MaskEdBox2.Text)
    cmd.Parameters.Append Param3
   
           
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 10, Get_Code(Trim(CboPurType)))
    cmd.Parameters.Append Param4
    
     
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 10, txtfields(7).Text)
    cmd.Parameters.Append Param5
     
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 150, IIf(Len(txtfields(2).Text) = 0, " ", txtfields(2).Text))
    cmd.Parameters.Append Param6
    
     Set Param7 = cmd.CreateParameter("param7", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param7
   
     
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param8
     
    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 30, txtfields(5).Text)
    cmd.Parameters.Append Param9
     
    Set Param10 = cmd.CreateParameter("param10", adDate, adParamInput, 10, IIf(MaskEdBox3.Text = "__/__/__", Null, MaskEdBox3.Text))
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 30, txtfields(6).Text)
    cmd.Parameters.Append Param11
   
   
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_issue_info_main(?,?,?,?,?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
     cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

End Sub
Private Sub edit_main()
   Dim RS As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    Dim Param10 As New Parameter
    
    

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 2)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 7, Trim(txtfields(0)))
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, MaskEdBox2.Text)
    cmd.Parameters.Append Param3
   
           
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 10, Get_Code(Trim(CboPurType)))
    cmd.Parameters.Append Param4
    
     
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 10, Get_Code(Trim(CboSupplier)))
    cmd.Parameters.Append Param5
     
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 150, txtfields(2).Text)
    cmd.Parameters.Append Param6
    
     Set Param7 = cmd.CreateParameter("param7", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param7
   
     
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param8
     
    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 10, txtfields(5).Text)
    cmd.Parameters.Append Param9
    
    Set Param10 = cmd.CreateParameter("param10", adDate, adParamInput, 10, IIf(MaskEdBox3.Text = "__/__/__", Null, MaskEdBox3.Text))
    cmd.Parameters.Append Param10
   
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_issue_info_main(?,?,?,?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
     cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

End Sub
Private Sub cmdnew_Click()
    txtfields(0) = ""
    txtfields(2) = ""
    cboItem.Text = ""
    txtItemTitle = ""
    txtfields(1) = 0
    txtfields(4) = 0
    txtfields(3) = 3
    txtfields(5) = ""
    txtfields(6) = ""
    txtfields(7) = ""
    txtfields(8) = ""
'    txtfields(7) = ""
    MaskEdBox2 = Format(Date, "dd/mm/yy")
    MaskEdBox3 = "__/__/__"
    txttrackid = ""
   MaskEdBox2.SetFocus
End Sub

Private Sub save(mode As Integer)
    Dim RS As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, mode)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 12, Trim(txtfields(0)))
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 12, Trim(CboPurId))
    cmd.Parameters.Append Param3
   
       
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 5, Get_Code(cboItem))
    cmd.Parameters.Append Param4
   
    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 18, Val(txtfields(1)))
    cmd.Parameters.Append Param5
   
        
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 10, Val(txtfields(4)))
    cmd.Parameters.Append Param6
    
       
    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 10, Val(txttrackid))
    cmd.Parameters.Append Param7
 
      
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_issue_info_S_P(?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
     cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
   
End Sub
Private Sub edit()
 Dim RS As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    Dim Param10 As New Parameter
    Dim Param11 As New Parameter
    Dim Param12 As New Parameter
    Dim Param13 As New Parameter
    Dim Param14 As New Parameter
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 2)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, Trim(txtfields(0)))
    cmd.Parameters.Append Param2
    
     Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, Trim(CboPurId))
    cmd.Parameters.Append Param3
   
    
    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, MaskEdBox2.Text)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 2, Get_Code(CboPurType))
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 4, "")
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 5, Get_Code(cboItem))
    cmd.Parameters.Append Param7
   
    Set Param8 = cmd.CreateParameter("param8", adDouble, adParamInput, 18, Val(txtfields(1)))
    cmd.Parameters.Append Param8
   
        
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 10, Val(txtfields(4)))
    cmd.Parameters.Append Param9
    
    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 150, Trim(txtfields(2)))
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("param11", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param12
    
    Set Param13 = cmd.CreateParameter("param13", adDouble, adParamInput, 10, Val(txttrackid))
    cmd.Parameters.Append Param13
 
      
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_issue_info_S_P(?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
     cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
      
End Sub
Private Sub delete()
     Dim RS As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    Dim Param10 As New Parameter
    Dim Param11 As New Parameter
    Dim Param12 As New Parameter
    Dim Param13 As New Parameter
    Dim Param14 As New Parameter
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 3)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, Trim(txtfields(0)))
    cmd.Parameters.Append Param2
    
     Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, Trim(CboPurId))
    cmd.Parameters.Append Param3
   
    
    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, MaskEdBox2.Text)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 2, Get_Code(CboPurType))
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 4, Trim(txtfields(7).Text))
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 5, Get_Code(cboItem))
    cmd.Parameters.Append Param7
   
    Set Param8 = cmd.CreateParameter("param8", adDouble, adParamInput, 18, Val(txtfields(1)))
    cmd.Parameters.Append Param8
   
        
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 10, Val(txtfields(4)))
    cmd.Parameters.Append Param9
    
    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 150, Trim(txtfields(2)))
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("param11", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param12
    
    Set Param13 = cmd.CreateParameter("param13", adDouble, adParamInput, 10, Val(txttrackid))
    cmd.Parameters.Append Param13
 
      
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_issue_info_S_P(?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
     cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
      
      
End Sub
Private Sub popup_delete()
    Dim RS As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 4)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 12, Trim(txtfields(0)))
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 12, Trim(CboPurId))
    cmd.Parameters.Append Param3
   
       
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 5, Get_Code(cboItem))
    cmd.Parameters.Append Param4
   
    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 18, Val(txtfields(1)))
    cmd.Parameters.Append Param5
   
        
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 10, Val(txtfields(4)))
    cmd.Parameters.Append Param6
    
       
    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 10, Val(txttrackid))
    cmd.Parameters.Append Param7
 
      
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_issue_info_S_P(?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
     cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
         
End Sub
Private Sub dtpic_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdsave.SetFocus
End If
End Sub
Private Sub Command1_Click(Index As Integer)
Dim RS As New ADODB.Recordset
Select Case Index
    Case 0
         form33.Show 1
    Case 1
       If txtfields(1).Text = "" Then
          txtfields(1).Text = 0
       End If
       If Len(txtfields(0)) = 0 Then
                MsgBox "Issue Serial Mandatory", vbInformation, App.title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(cboItem) = 0 Then
                MsgBox "Please Select an Item", vbInformation, App.title
                cboItem.SetFocus
                Exit Sub
            End If
            
            If Len(CboPurId) = 0 Then
                MsgBox "Please Select an Purchase ID", vbInformation, App.title
                CboPurId.SetFocus
                Exit Sub
            End If
            
'            If Len(CboSupplier) = 0 Then
'                MsgBox "Issue To Required...", vbInformation, App.title
'                CboSupplier.SetFocus
'                Exit Sub
'            End If
'
                      
            
             If Val(txtfields(1)) = 0 Then
                MsgBox "Quantity Must be More than One", vbInformation, App.title
                CboPurId.SetFocus
                Exit Sub
            End If
            
                           
            
            
            Set objRs = objcom.Get_RS("SELECT  PurDate from PurchaseMain  WHERE (PurId= '" & Trim(CboPurId) & "')", objmyCon)
            
           If Not objRs.EOF Then
              If CDate(Format(MaskEdBox2.Text, "dd mmm yyyy")) < CDate(Format(objRs(0), "dd mmm yyyy")) Then
                 MsgBox "Issued Date Can't be Less than Purchase Date ( " & (objRs(0)) & "  ),Please Verify", vbInformation, cmp
                 Exit Sub
              End If
           End If
            
            Set objRs = objcom.Get_RS("SELECT purId from purchasesub  WHERE (purId= '" & Trim(CboPurId) & "') and to_number(itemId)= to_number('" & Trim(Get_Code(cboItem)) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "Invalid Purchase ID...Please Verify.", vbInformation, cmp
               CboPurId.SetFocus
               Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT exp_date,sysdate  from PurchaseSub   WHERE (PurId= '" & CboPurId & "' and to_number(itemId)=to_number('" & Get_Code(cboItem) & "') )", objmyCon)
            
            If Not objRs.EOF Then
             If Not IsNull(objRs(0)) Then
              If Format(objRs(0), "dd/mmm/yy") < CDate(Format(objRs(1), "dd/mmm/yy")) Then
                 MsgBox "Date Expired..Please Verify", vbInformation, cmp
                 CboPurId.SetFocus
                 Exit Sub
              End If
            End If
                 
            End If
            
         
           
            Set objRs = objcom.Get_RS("SELECT IssueId from IssueMain  WHERE (IssueId= '" & txtfields(0) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "Invalid Issue Serial...Please Verify.", vbInformation, cmp
               txtfields(0).SelLength = Len(txtfields(0))
               txtfields(0).SetFocus
               Exit Sub
            End If
            
          
            
'          Set objRs = objcom.Get_RS("SELECT PurchaseId from issueSub  WHERE (PurchaseId= '" & Trim(CboPurId) & "' and issueId= '" & Trim(txtfields(0)) & "' and itemcode='" & Trim(Get_Code(cboItem)) & "')", objmyCon)
'          If Not objRs.EOF Then
'              Exit Sub
'           End If
'
        Set objRs = objcom.Get_RS("SELECT ItemCode from IssueSub  WHERE (IssueId= '" & txtfields(0) & "') and to_number(ItemCode)=to_number('" & cboItem & "') and PurchaseId='" & CboPurId & "'", objmyCon)
         
        If objRs.EOF Then
              Set objRs = objcom.Get_RS("SELECT (PurQty-(UsedQty+ReturnQty)) as PurQty from PurchaseSub  WHERE (PurId= '" & Trim(CboPurId) & "' and to_number(itemId)=to_number('" & Trim(Get_Code(cboItem)) & "'))", objmyCon)
            If Not objRs.EOF Then
               If Val(objRs(0)) < Val(txtfields(1).Text) Then
                  MsgBox "Insufficient Quantity...Please Verify", vbInformation, cmp
                  txtfields(1).SelLength = Len(txtfields(1).Text)
                  CboPurId.SetFocus
                  Exit Sub
               End If
          End If
             save (1)
        Else
           If Len(txttrackid.Text) = 0 Then
              MsgBox "Please Select an item from the Grid below to Edit", vbCritical, cmp
              MSFlexGrid1.SetFocus
              Exit Sub
           End If
           Dim balance As Double
           Dim preqty As Double
  
          Set objRs = objcom.Get_RS("SELECT (PurQty-(UsedQty+ReturnQty)) as balance from PurchaseSub  WHERE (PurId= '" & Trim(CboPurId) & "' and itemId='" & Trim(Get_Code(cboItem)) & "')", objmyCon)
  
            If Not objRs.EOF Then
               balance = objRs(0)
            End If
            
          Set objRs = objcom.Get_RS("SELECT qty as preqty from IssueSub  WHERE TrackId = " & Val(txttrackid) & "", objmyCon)
          If Not objRs.EOF Then
               preqty = objRs(0)
            End If
            
         If (balance + preqty) < Val(txtfields(1).Text) Then
            MsgBox "Insufficient Balance..Please Verify.", vbInformation, cmp
            CboPurId.SetFocus
            Exit Sub
        End If
          save (2)
    End If
       load_purid
       Call ShowFlexData
       txtfields(1).Text = 0
       txtfields(4).Text = 0
       txttrackid = ""
       cboItem.SetFocus
   Case 2
        form34.Show 1
   End Select

End Sub

Private Sub Command2_Click()
 If Len(txtfields(0)) = 0 Then
    MsgBox "Issue Serial Mandatory", vbInformation, App.title
    txtfields(0).SetFocus
    Exit Sub
End If
 
Set objRs = objcom.Get_RS("SELECT IssueId from IssueMain  WHERE (IssueId= '" & txtfields(0) & "')", objmyCon)
            
 If objRs.EOF Then
   MsgBox "Invalid Issue Serial...Please Verify.", vbInformation, cmp
   txtfields(0).SelLength = Len(txtfields(0))
   txtfields(0).SetFocus
   Exit Sub
End If
            
  rptmode = 11
 rptViewer.Show 1

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
'      SendKeys (Chr(9))
   ElseIf KeyAscii = 27 Then
          Unload Me
          
   ElseIf KeyAscii = 14 Then
      cmdnew_Click
  End If
       
   
  
   
End Sub
Private Sub Form_Load()

load_issue_type


MaskEdBox2.Text = Format(Date, "dd/mm/yy")
Dim RS As New ADODB.Recordset

With MSFlexGrid1
    .Rows = 1
    .Cols = 9
    .Col = 0: .Text = "CatCode"
    .Col = 1: .Text = "Type"
    .Col = 2: .Text = " Code"
    .Col = 3: .Text = " Title"
    .Col = 4: .Text = "Purchase Id"
    .Col = 5: .Text = " Quantity"
    .Col = 6: .Text = "Unit Rate"
    .Col = 7: .Text = "Total"
    .Col = 8: .Text = "Serail no"
   
   
    .ColWidth(0) = 0
    .ColWidth(1) = 1400
    .ColWidth(2) = 0
    .ColWidth(3) = 4700
    .ColWidth(4) = 1350
    .ColWidth(5) = 850
    .ColWidth(6) = 1150
    .ColWidth(7) = 2150
    .ColWidth(8) = 0
    

    
End With
Call ShowFlexData
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title

End Sub
Private Sub load_issueto(Index As Integer)
 
 Select Case Index
        Case 0
           Set objRs = objcom.Get_RS("SELECT emp_id,emp_nm  from payroll.emp_info where upper(emp_id) not like upper('M%') or upper(emp_id) not like upper('c%') order by emp_id", objmyCon)
           CboSupplier.Clear
           If Not objRs.EOF Then
              objRs.MoveFirst
              Do Until objRs.EOF
                 CboSupplier.AddItem Trim(objRs(1)) + "~" + Trim(objRs(0))
                 objRs.MoveNext
              Loop
           End If
        Case 1
           Set objRs = objcom.Get_RS("SELECT emp_id,emp_nm  from payroll.emp_info where upper(emp_id)  like upper('M%')  order by emp_id", objmyCon)
           CboSupplier.Clear
           If Not objRs.EOF Then
              objRs.MoveFirst
              Do Until objRs.EOF
                 CboSupplier.AddItem Trim(objRs(1)) + "~" + Trim(objRs(0))
                 objRs.MoveNext
              Loop
           End If
       Case 2
           Set objRs = objcom.Get_RS("SELECT bed_no,bed_type,BED_EXT_COL  from hospital_billing.bed_info  order by bed_type", objmyCon)
           CboSupplier.Clear
           If Not objRs.EOF Then
              objRs.MoveFirst
              Do Until objRs.EOF
                 CboSupplier.AddItem Trim(objRs(1)) + "~" + Trim(objRs(2)) + "-" + Trim(objRs(0))
                 objRs.MoveNext
              Loop
           End If
                
'       Case 2
'           Set objRs = objcom.Get_RS("SELECT emp_id,emp_nm  from payroll.emp_info where upper(emp_id)  like upper('c%')  order by emp_id", objmyCon)
'           CboSupplier.Clear
'           If Not objRs.EOF Then
'              objRs.MoveFirst
'              Do Until objRs.EOF
'                 CboSupplier.AddItem Trim(objRs(1)) + "~" + Trim(objRs(0))
'                 objRs.MoveNext
'              Loop
'           End If
'
         Case 3
           Set objRs = objcom.Get_RS("SELECT refer_code,doc_dept  from hospital_billing.doctor_info  order by refer_code", objmyCon)
           CboSupplier.Clear
           If Not objRs.EOF Then
              objRs.MoveFirst
              Do Until objRs.EOF
                 CboSupplier.AddItem Trim(objRs(1)) + "~" + Trim(objRs(0))
                 objRs.MoveNext
              Loop
           End If
           
           
      Case 4
'           Set objRs = objcom.Get_RS("SELECT refer_code,doc_dept  from hospital_billing.doctor_info  order by refer_code", objmyCon)
           CboSupplier.Clear
'           If Not objRs.EOF Then
'              objRs.MoveFirst
'              Do Until objRs.EOF
'                 CboSupplier.AddItem Trim(objRs(1)) + "~" + Trim(objRs(0))
'                 objRs.MoveNext
'              Loop
'           End If
           
  End Select
  
End Sub
Private Sub load_issue_to_specific(Index As Integer, value As String)
  Select Case Index
        Case 0
           Set objRs = objcom.Get_RS("SELECT emp_id,emp_nm  from payroll.emp_info where upper(emp_id)=upper(value) order by emp_id", objmyCon)
           CboSupplier.Clear
           If Not objRs.EOF Then
                 CboSupplier.AddItem Trim(objRs(1)) + "~" + Trim(objRs(0))
           End If
        Case 1
           Set objRs = objcom.Get_RS("SELECT emp_id,emp_nm  from payroll.emp_info where upper(emp_id)=upper('" & value & "')  order by emp_id", objmyCon)
         
           If Not objRs.EOF Then
                CboSupplier.Text = (objRs(1)) + "~" + Trim(objRs(0))
            End If
       Case 2
           Set objRs = objcom.Get_RS("SELECT bed_no,bed_type,BED_EXT_COL  from hospital_billing.bed_info  order by bed_type", objmyCon)
           CboSupplier.Clear
           If Not objRs.EOF Then
              objRs.MoveFirst
              Do Until objRs.EOF
                 CboSupplier.AddItem Trim(objRs(1)) + "~" + Trim(objRs(2)) + "-" + Trim(objRs(0))
                 objRs.MoveNext
              Loop
           End If
                
'       Case 2
'           Set objRs = objcom.Get_RS("SELECT emp_id,emp_nm  from payroll.emp_info where upper(emp_id)  like upper('c%')  order by emp_id", objmyCon)
'           CboSupplier.Clear
'           If Not objRs.EOF Then
'              objRs.MoveFirst
'              Do Until objRs.EOF
'                 CboSupplier.AddItem Trim(objRs(1)) + "~" + Trim(objRs(0))
'                 objRs.MoveNext
'              Loop
'           End If
'
         Case 3
           Set objRs = objcom.Get_RS("SELECT refer_code,doc_dept  from hospital_billing.doctor_info  order by refer_code", objmyCon)
           CboSupplier.Clear
           If Not objRs.EOF Then
              objRs.MoveFirst
              Do Until objRs.EOF
                 CboSupplier.AddItem Trim(objRs(1)) + "~" + Trim(objRs(0))
                 objRs.MoveNext
              Loop
           End If
           
           
      Case 4
'           Set objRs = objcom.Get_RS("SELECT refer_code,doc_dept  from hospital_billing.doctor_info  order by refer_code", objmyCon)
           CboSupplier.Clear
'           If Not objRs.EOF Then
'              objRs.MoveFirst
'              Do Until objRs.EOF
'                 CboSupplier.AddItem Trim(objRs(1)) + "~" + Trim(objRs(0))
'                 objRs.MoveNext
'              Loop
'           End If
           
  End Select
 
End Sub
Private Sub load_item()
 Set objRs = objcom.Get_RS("SELECT item_code ,item_name from item_info where group_code='" & Get_Code(CategoryCode) & "' order by item_code", objmyCon)
 cboItem.Clear
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       cboItem.AddItem Trim(objRs(1)) + "~" + Trim(objRs(0))
       objRs.MoveNext
    Loop
 End If
End Sub
Private Sub load_issue_type()
 Set objRs = objcom.Get_RS("SELECT type_code,type_name from item_issue_type order by type_code", objmyCon)
 CboPurType.Clear
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       CboPurType.AddItem objRs(1) + "~" + objRs(0)
       objRs.MoveNext
    Loop
 End If
  
End Sub
Private Sub load_group()
 Set objRs = objcom.Get_RS("SELECT group_code,group_name from item_group_info order by group_code", objmyCon)
 CboGroup.Clear
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       CboGroup.AddItem objRs(1) + "~" + objRs(0)
       objRs.MoveNext
    Loop
 End If
End Sub
Private Sub load_purid()

 Set objRs = objcom.Get_RS("select PurId  from PurchaseSub  where to_number(itemId) = to_number('" & Get_Code(cboItem) & "') and (PurQty-(UsedQty+ReturnQty)) > 0 order by PurId", objmyCon)
 CboPurId.Clear
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       CboPurId.AddItem objRs(0)
       objRs.MoveNext
    Loop
  End If

End Sub
Private Sub MaskEdBox1_GotFocus()
  MaskEdBox1.SetFocus
  MaskEdBox1.SelLength = Len(MaskEdBox1)
End Sub
Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    If MaskEdBox1 <> "__/__/__" Then
            If Check_ValidDate(MaskEdBox1) = False Then
                MaskEdBox1.Text = "__/__/__"
                MaskEdBox1.SetFocus
                Exit Sub
            End If
      End If
    
End If
End Sub
Private Sub MaskEdBox2_GotFocus()
    MaskEdBox2.SelLength = Len(MaskEdBox2)
End Sub
Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If MaskEdBox2 <> "__/__/__" Then
            If Check_ValidDate(MaskEdBox2) = False Then
                MaskEdBox2.Text = "__/__/__"
                MaskEdBox2.SetFocus
                Exit Sub
            Else
               txtfields(5).SetFocus
            End If
    Else
       CboPurType.SetFocus
    End If
 End If
End Sub

Private Sub MaskEdBox3_GotFocus()
   MaskEdBox3.SelLength = Len(MaskEdBox3)
End Sub

Private Sub MaskEdBox3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  If MaskEdBox3 <> "__/__/__" Then
            If Check_ValidDate(MaskEdBox3) = False Then
                MaskEdBox3.Text = "__/__/__"
                MaskEdBox3.SetFocus
                Exit Sub
            Else
              txtfields(6).SetFocus
            End If
    Else
       txtfields(6).SetFocus
    End If
 End If
End Sub

Private Sub mnuClose_Click()
  Unload Me
End Sub

Private Sub mnuDL_Click()
         If Len(txtfields(0)) = 0 Then
                MsgBox "Issue Serial Mandatory", vbInformation, App.title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(cboItem) = 0 Then
                MsgBox "Please Select an Item", vbInformation, App.title
                cboItem.SetFocus
                Exit Sub
            End If
            
            If Len(CboPurId) = 0 Then
                MsgBox "Please Select an Purchase ID", vbInformation, App.title
                CboPurId.SetFocus
                Exit Sub
            End If
            
           
            
             If Val(txtfields(3)) = 0 Then
                MsgBox "Amount Must be More than One", vbInformation, App.title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
           
            Set objRs = objcom.Get_RS("SELECT IssueId from IssueMain  WHERE (IssueId= '" & txtfields(0) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "Invalid Issue Serial..Please Verify.", vbInformation, cmp
               txtfields(0).SelLength = Len(txtfields(0))
               txtfields(0).SetFocus
               Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT trackid from Issuesub  WHERE (IssueId= '" & txtfields(0) & "') and trackid='" & txttrackid & "' ", objmyCon)
            
            If objRs.EOF Then
               MsgBox "Invalid Issue..Please Verify.", vbInformation, cmp
               Exit Sub
            End If
            
            popup_delete

      
       Call ShowFlexData
       cmdnew.SetFocus
  
End Sub

Private Sub mnuUpdate_Click()

    If Len(txtfields(0)) = 0 Then
        MsgBox "Issue Serial Mandatory", vbInformation, App.title
        txtfields(0).SetFocus
        Exit Sub
    End If
            
            
            If Len(cboItem) = 0 Then
                MsgBox "Please Select an Item", vbInformation, App.title
                cboItem.SetFocus
                Exit Sub
            End If
            
            If Len(CboPurId) = 0 Then
                MsgBox "Please Select an Purchase ID", vbInformation, App.title
                CboPurId.SetFocus
                Exit Sub
            End If
            
           
            
             If Val(txtfields(3)) = 0 Then
                MsgBox "Amount Must be More than One", vbInformation, App.title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
             Set objRs = objcom.Get_RS("SELECT purId from purchasesub  WHERE (purId= '" & Trim(CboPurId) & "') and itemId= '" & Trim(Get_Code(cboItem)) & "'", objmyCon)
            
            If objRs.EOF Then
               MsgBox "Invalid Purchase ID...Please Verify.", vbInformation, cmp
               CboPurId.SetFocus
               Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT IssueId from IssueMain  WHERE (IssueId= '" & txtfields(0) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "Invalid Issue Serial..Please Verify.", vbInformation, cmp
               txtfields(0).SelLength = Len(txtfields(0))
               txtfields(0).SetFocus
               Exit Sub
            End If
            
             Set objRs = objcom.Get_RS("SELECT exp_date,sysdate  from PurchaseSub   WHERE (PurId= '" & CboPurId & "' and itemId='" & Get_Code(cboItem) & "' )", objmyCon)
            
             If Not objRs.EOF Then
             If Not IsNull(objRs(0)) Then
              If Format(objRs(0), "dd/mmm/yy") < CDate(Format(objRs(1), "dd/mmm/yy")) Then
                 MsgBox "Date Expired..Please Verify", vbInformation, cmp
                 CboPurId.SetFocus
                 Exit Sub
              End If
            End If
                 
            End If
            
            
           Dim balance As Double
           Dim preqty As Double
  
          Set objRs = objcom.Get_RS("SELECT (PurQty-(UsedQty+ReturnQty)) as balance from PurchaseSub  WHERE (PurId= '" & Trim(CboPurId) & "' and itemId='" & Trim(Get_Code(cboItem)) & "')", objmyCon)
  
            If Not objRs.EOF Then
               balance = objRs(0)
            End If
            
          Set objRs = objcom.Get_RS("SELECT qty as preqty from IssueSub  WHERE TrackId = " & Val(txttrackid) & "", objmyCon)
          If Not objRs.EOF Then
               preqty = objRs(0)
            End If
            
         If (balance + preqty) < Val(txtfields(1).Text) Then
            MsgBox "Insufficient Balance..Please Verify.", vbInformation, cmp
            CboPurId.SetFocus
            Exit Sub
        End If
        
            edit

       MsgBox "Edited successfully.", vbInformation, cmp
       Call ShowFlexData
       cmdnew.SetFocus
  
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
      PopupMenu MnuDelete, 2
   End If
End Sub

Private Sub MSFlexGrid1_SelChange()
  MSFlexGrid1_Click
End Sub

Private Sub MSFlexGrid2_DblClick()
  If Len(Trim(MSFlexGrid2.Text)) <> 0 Then
      cboItem.Text = MSFlexGrid2.Text
      Set objRs = objcom.Get_RS("select item_name from item_info where to_number(item_code)='" & Trim(cboItem.Text) & "'", objmyCon)
      If Not objRs.EOF Then
        txtItemTitle.Text = objRs(0)
         CboPurId.SetFocus
         load_purid
      Else
         txtItemTitle.Text = ""
         cboItem.SetFocus
         Exit Sub
      End If
  End If
 
  MSFlexGrid2.Visible = False
End Sub
Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      MSFlexGrid2_DblClick
  End If
End Sub
Private Sub txtfields_Change(Index As Integer)
    Select Case Index
          Case 0
              If Len(txtfields(0).Text) > 0 Then
                  CmdGenerate.Enabled = False
               Else
                 CmdGenerate.Enabled = True
               End If
               ShowFlexData
              
         Case 4
               If Not IsNumeric(txtfields(4)) Then
                   txtfields(4) = ""
               Else
                  txtfields(3) = Val(txtfields(1)) * Val(txtfields(4))
               End If
                     
        Case 1
               If Not IsNumeric(txtfields(1)) Then
                   txtfields(1) = ""
               Else
                  txtfields(3) = Val(txtfields(1)) * Val(txtfields(4))
               End If
               
        Case 5
               If Not IsNumeric(txtfields(5)) Then
                   txtfields(5) = ""
               End If
               
End Select
End Sub
Private Function select_pur_tpy(str As String) As String
        Select Case str
               Case "T"
                   select_pur_tpy = Trim("Tender~T")
               Case "D"
                   select_pur_tpy = Trim("Donation~D")
               Case "L"
                   select_pur_tpy = Trim("Local~L")
                Case "H"
                   select_pur_tpy = Trim("Other~H")
          End Select
    
End Function
Private Sub txtfields_GotFocus(Index As Integer)
     Select Case Index
                 Case 1
                      txtfields(1).SelStart = 0
                      txtfields(1).SelLength = Len(txtfields(1))
                 Case 4
                      txtfields(4).SelStart = 0
                      txtfields(4).SelLength = Len(txtfields(4))
                 Case 5
                      txtfields(5).SelStart = 0
                      txtfields(5).SelLength = Len(txtfields(5))
                 Case 6
                      txtfields(6).SelStart = 0
                      txtfields(6).SelLength = Len(txtfields(6))
     End Select
     
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
Dim objrs_local As ADODB.Recordset
Dim objrs_name As ADODB.Recordset
Dim value As String
If KeyAscii = 13 Then
    Select Case Index
        Case 0
         If Len(txtfields(0).Text) > 0 Then
                CmdGenerate.Enabled = False
               Else
                CmdGenerate.Enabled = True
               End If
             If Len(txtfields(0)) <> 0 Then
              If Mid(txtfields(0), 1, 1) <> "I" Then
                    txtfields(0).Text = Format(txtfields(0), "000000")
                      txtfields(0).Text = "I-" + CategoryCode + "-" + txtfields(0).Text
                End If
              End If
                   ShowFlexData
               Set objrs_local = objcom.Get_RS("SELECT  a.IssueDate,a.IssueType,b.type_name,a.Issueto,a.Remarks,a.indent_no,a.indent_date,a.reg_no  from IssueMain a,item_issue_type b WHERE (a.issueId= '" & txtfields(0) & "' and  to_number(a.IssueType)=to_number(b.type_code))", objmyCon)
            
              If Not objrs_local.EOF Then
                 MaskEdBox2 = Format(objrs_local(0), "dd/mm/yy")
                 MaskEdBox3 = IIf(IsNull(objrs_local!indent_date) = True, "__/__/__", Format(objrs_local!indent_date, "dd/mm/yy"))
                 txtfields(5).Text = "" & objrs_local!indent_no
                 txtfields(6).Text = "" & objrs_local!reg_no
                 CboPurType.Text = Trim(objrs_local(2)) + "~" + Trim(objrs_local(1)) '''issue type
                 txtfields(7).Text = "" & objrs_local!issueto
                
                If objrs_local!issueto <> "" Then
                   Set objrs_name = objcom.Get_RS("SELECT emp_nm  from payroll.emp_info where upper(emp_id) = upper('" & Trim(txtfields(7).Text) & "')", objmyCon)
                   If Not objrs_name.EOF Then
                      txtfields(8).Text = objrs_name(0)
                   End If
                End If
'                 txtfields(2).Text = "" & objRs(4)
              End If
              MaskEdBox2.SetFocus
              
    Case 1
          Command1(1).SetFocus
    Case 2
         If CmdGenerate.Enabled = True Then
            CmdGenerate.SetFocus
         Else
            cboItem.SetFocus
         End If
   Case 5
       MaskEdBox3.SetFocus
    Case 6
       CboPurType.SetFocus
    Case 7
           If txtfields(7).Text = "" Then
              txtfields(2).SetFocus
              
           Else
             Set objRs = objcom.Get_RS("SELECT emp_nm  from payroll.emp_info where upper(emp_id)=upper('" & txtfields(7).Text & "')", objmyCon)
             If Not objRs.EOF Then
               txtfields(8).Text = "" & objRs(0)
               txtfields(2).SetFocus
             End If
           End If
    End Select
    'txtfields(2).SetFocus
End If
End Sub
Public Function getdata(SQLString As String) As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.Connection
Dim RS As New ADODB.Recordset
con.Open objmyCon
Set cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = SQLString

 Set RS = cmd.Execute
Set getdata = RS
End Function

Private Sub txtFields_LostFocus(Index As Integer)
  Select Case Index
    Case 4
       If txtfields(4).Text = "" Then
          txtfields(4).Text = 0
       End If
  End Select
End Sub
Private Sub ShowFlexData()
On Error GoTo errdes
Dim RS As New ADODB.Recordset
'Set RS = objcom.Get_RS("SELECT a.group_code as Code ,a.group_name as Title,a.cate_code,(select b.cate_name from item_cate_info b where b.cate_code=a.cate_code) as cat_title, a.remarks  From item_group_info a", objmyCon)
Set RS = objcom.Get_RS("SELECT c.group_code, c.group_name,a.ItemCode  as Code ,b.item_name,a.PurchaseId,a.URate,a.Qty,a.TrackId   From issueSub a ,item_info b, item_group_info c where to_number(b.group_code) = to_number(c.group_code) and a.issueId='" & txtfields(0).Text & "' and to_number(b.item_code)=to_number(a.ItemCode) order by a.ItemCode ", objmyCon)
If Not RS.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until RS.EOF
            MSFlexGrid1.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = Trim(RS(0))
                .TextMatrix(i, 1) = Trim(RS(1))
                .TextMatrix(i, 2) = RS(2)
                 MSFlexGrid1.ColAlignment(3) = 1
                .TextMatrix(i, 3) = RS(3)
                .TextMatrix(i, 4) = RS(4)
                .TextMatrix(i, 5) = RS(6)
                .TextMatrix(i, 6) = Val(RS(5))
                .TextMatrix(i, 7) = Val(RS(6)) * Val(RS(5))
                .TextMatrix(i, 8) = RS(7)
                i = i + 1
            RS.MoveNext
        Loop
    End With
Else
    MSFlexGrid1.Rows = 1
 End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title
End Sub
Private Sub MSFlexGrid1_Click()
txtfields(3) = 0
On Error Resume Next
If MSFlexGrid1.Row >= 1 Then
'    Cbocategory = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) + "~" + Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0))
    cboItem = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
    txtItemTitle = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3))
    CboPurId = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
    txtfields(1).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
    txtfields(4).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
    txtfields(3).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)
    '''MaskEdBox1.Text = IIf(Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7), "dd/mm/yy") = "", Trim("__/__/__"), Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7), "dd/mm/yy"))
    txttrackid.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8)
    
End If

Exit Sub

errdes:
'MsgBox err.Description, vbInformation, App.Title


End Sub
Private Function select_acct(str As String) As String
        Dim loacalRs As New ADODB.Recordset
        Set loacalRs = objcom.Get_RS("SELECT acc_name,acc_code from acct.acct  WHERE (acc_code= '" & str & "')", objmyCon)
            
              If Not loacalRs.EOF Then
                select_acct = Trim(loacalRs(0)) + "~" + Trim(loacalRs(1))
              End If
End Function
Private Function select_issueto(str As String) As String
        Dim loacalRs As New ADODB.Recordset
        Set loacalRs = objcom.Get_RS("SELECT type_name,type_code from item_issue_type  WHERE (type_code= '" & str & "')", objmyCon)
            
              If Not loacalRs.EOF Then
                select_issueto = Trim(loacalRs(0)) + "~" + Trim(loacalRs(1))
              End If
End Function
Public Function Check_ValidDate(InitialDate As String) As Boolean
    Dim Day1 As Integer
    Dim Month1 As Integer
    Dim Year1 As Integer
    Dim IDate As Integer
    If IsDate(InitialDate) = False Then
        MsgBox "Invalid Date.", vbInformation, "Daffodil Software"
        Check_ValidDate = False
        Exit Function
    End If
    Day1 = Mid(InitialDate, 1, 2)
    Month1 = Mid(InitialDate, 4, 2)
    Year1 = Mid(InitialDate, 7, 2)
    If Year1 < 50 Then
        IDate = "20" + Format(Year1, "00")
    Else
        IDate = "19" + Format(Year1, "00")
    End If
    Dim Month As Integer
    Dim Day As Integer
    Dim Year As Integer
    Month = Month1
    Day = Day1
    Year = IDate
    If Month = 4 Or Month = 6 Or Month = 9 Or Month = 11 Then
        If Day > 30 Then
            MsgBox "Invalid day format.", vbInformation, "Daffodil Software"
            Check_ValidDate = False
        Else
            Check_ValidDate = True
        End If
    ElseIf Month = 1 Or Month = 3 Or Month = 5 Or Month = 7 Or Month = 8 Or Month = 10 Or Month = 12 Then
        If Day > 31 Then
            MsgBox "Invalid day format.", vbInformation, "Daffodil Software"
            Check_ValidDate = False
        Else
            Check_ValidDate = True
        End If
    ElseIf Month = 2 Then
            If Year Mod 4 = 0 Then
                    If Day > 29 Then
                        MsgBox "Invalid day Format.", vbInformation, "Daffodil Software"
                        Check_ValidDate = False
                    Else
                        Check_ValidDate = True
                    End If
            Else
                If Day > 28 Then
                    MsgBox "Invalid Day Format.", vbInformation, "Daffodil Software"
                    Check_ValidDate = False
                Else
                    Check_ValidDate = True
                End If
            End If
    ElseIf Month > 12 Then
            MsgBox "Invalid Month Format.", vbInformation, "Daffodil Software"
            Check_ValidDate = False
    Else
            Check_ValidDate = True
    End If
End Function
