VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmIssueReturn 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   2430
      Left            =   60
      TabIndex        =   40
      Top             =   3150
      Visible         =   0   'False
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   4286
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
      TabIndex        =   7
      Text            =   "0"
      ToolTipText     =   "Insert Quantity"
      Top             =   2820
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   1695
      Left            =   -30
      TabIndex        =   27
      Top             =   840
      Width           =   11955
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1650
         MaxLength       =   12
         TabIndex        =   0
         ToolTipText     =   "Issue  Serial"
         Top             =   165
         Width           =   1755
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   2
         Left            =   1650
         MaxLength       =   100
         TabIndex        =   4
         ToolTipText     =   "Insert Remarks"
         Top             =   1095
         Width           =   7365
      End
      Begin VB.CommandButton Command1 
         Caption         =   ":::"
         Height          =   315
         Index           =   0
         Left            =   3420
         TabIndex        =   28
         Top             =   150
         Width           =   375
      End
      Begin VB.ComboBox CboPurType 
         Height          =   315
         ItemData        =   "frmIssueReturnInfo.frx":0000
         Left            =   1650
         List            =   "frmIssueReturnInfo.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1755
      End
      Begin VB.ComboBox CboSupplier 
         Height          =   315
         ItemData        =   "frmIssueReturnInfo.frx":0030
         Left            =   6090
         List            =   "frmIssueReturnInfo.frx":0037
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   5385
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   285
         Left            =   6090
         TabIndex        =   1
         Top             =   180
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
      Begin LVbuttons.LaVolpeButton CmdGenerate 
         Height          =   435
         Left            =   9060
         TabIndex        =   5
         ToolTipText     =   "Press to Generate Issue Return Serial"
         Top             =   1080
         Width           =   2475
         _ExtentX        =   4366
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
         MICON           =   "frmIssueReturnInfo.frx":004A
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Return  Serial#"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   33
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   1140
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Date "
         Height          =   195
         Index           =   1
         Left            =   4770
         TabIndex        =   31
         Top             =   210
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Return Type"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   30
         Top             =   660
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Return From"
         Height          =   195
         Index           =   3
         Left            =   4770
         TabIndex        =   29
         Top             =   630
         Width           =   870
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000B&
      Height          =   705
      Left            =   0
      TabIndex        =   21
      Top             =   2460
      Width           =   11925
      Begin VB.TextBox txtItemTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAF2C8&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   360
         Width           =   4305
      End
      Begin VB.TextBox cboItem 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         TabIndex        =   36
         Top             =   360
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0097C8E6&
         Caption         =   ">>"
         Height          =   315
         Index           =   1
         Left            =   11550
         MaskColor       =   &H00C0C000&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   345
      End
      Begin VB.ComboBox CboPurId 
         Height          =   315
         Left            =   5730
         TabIndex        =   6
         Text            =   "CboPurId"
         Top             =   360
         Width           =   1755
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   10650
         TabIndex        =   9
         Top             =   360
         Width           =   885
         _ExtentX        =   1561
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
         Enabled         =   0   'False
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
         MaxLength       =   15
         TabIndex        =   11
         Text            =   "0"
         Top             =   360
         Width           =   1155
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
         TabIndex        =   10
         Text            =   "0"
         Top             =   360
         Width           =   1155
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
         TabIndex        =   39
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
         TabIndex        =   38
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Id"
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
         Left            =   5790
         TabIndex        =   34
         Top             =   150
         Width           =   600
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
         Left            =   10650
         TabIndex        =   25
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
         TabIndex        =   24
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
         Left            =   7470
         TabIndex        =   23
         Top             =   120
         Width           =   330
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
         Left            =   8310
         TabIndex        =   22
         Top             =   120
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   615
      Left            =   0
      TabIndex        =   18
      Top             =   7260
      Width           =   11925
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3990
         TabIndex        =   35
         Text            =   "CboPurId"
         Top             =   210
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox txttrackid 
         Height          =   285
         Left            =   1200
         TabIndex        =   26
         Text            =   "Text1"
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
         Left            =   8790
         TabIndex        =   19
         ToolTipText     =   "Click to Edit Information"
         Top             =   180
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
         Left            =   6840
         TabIndex        =   13
         ToolTipText     =   "Click to insert new information"
         Top             =   180
         Width           =   945
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H8000000C&
         Caption         =   "Save"
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
         Left            =   7815
         TabIndex        =   12
         ToolTipText     =   "Click to Save"
         Top             =   180
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
         Left            =   9765
         TabIndex        =   14
         ToolTipText     =   "Click to Delete"
         Top             =   180
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
         Left            =   10740
         TabIndex        =   15
         ToolTipText     =   "Click to Close"
         Top             =   180
         Width           =   945
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   6810
         Top             =   150
         Width           =   4935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4155
      Left            =   0
      TabIndex        =   17
      Top             =   3150
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   7329
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   13627123
      ForeColor       =   12582912
      BackColorSel    =   -2147483630
      BackColorBkg    =   -2147483637
      GridColor       =   -2147483637
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000001&
      FillColor       =   &H00E0E0E0&
      Height          =   915
      Left            =   -30
      ScaleHeight     =   855
      ScaleWidth      =   11895
      TabIndex        =   16
      Top             =   -30
      Width           =   11955
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Issue Return  Entry"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEF0F7&
         Height          =   405
         Left            =   4350
         TabIndex        =   20
         Top             =   180
         Width           =   3810
      End
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mnuDL 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmIssueReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objcom As New DSLComFram.clsCommon
Dim objRs As New ADODB.Recordset
Dim paramval As Double
Dim issuedqty As Double
Dim prevReturn As Double

Private Sub Cbocategory_Click()
    load_item
End Sub

Private Sub CboItem_Click()
    load_purid
End Sub

Private Sub load_qty()
  Set objRs = objcom.Get_RS("SELECT Qty,URate from IssueSub  WHERE (IssueId= '" & Trim(CboPurId) & "' and ItemCode='" & Trim(Get_Code(cboItem)) & "') and purchaseid ='" & Trim(Combo1) & "'", objmyCon)

       If Not objRs.EOF Then
          txtfields(1) = objRs(0)
          txtfields(4) = objRs(1)
      End If
End Sub

Private Sub cboItem_GotFocus()
  cboItem.SelStart = 0
  cboItem.SelLength = Trim(Len(cboItem))
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
    MSFlexGrid2.SetFocus
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

Private Sub cboItem_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Len(Trim(cboItem.Text)) = 0 Then
       txtItemTitle = ""
     Exit Sub
  End If
  
  
  If IsNumeric(cboItem) Then
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

Private Sub CboPurId_Click()
    'load_qty
    Combo1.Text = Combo1.List(CboPurId.ListIndex)
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


Private Sub CboPurId_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtfields(1).SetFocus
   End If
End Sub

Private Sub CboPurType_Click()
   load_issueto (CboPurType.ListIndex)
End Sub

Private Sub CboPurType_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     CboSupplier.SetFocus
  End If
End Sub

Private Sub CboSupplier_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     txtfields(2).SetFocus
  End If
End Sub

Private Sub cmdDelete_Click()
           If txtfields(1).Text = "" Then
          txtfields(1).Text = 0
       End If
       If Len(txtfields(0)) = 0 Then
                MsgBox "Issue Return Serial Mandatory", vbInformation, App.title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(cboItem) = 0 Then
                MsgBox "Please Select an Item", vbInformation, App.title
                cboItem.SetFocus
                Exit Sub
            End If
            
            If Len(CboPurId) = 0 Then
                MsgBox "Please Select an Issue ID", vbInformation, App.title
                CboPurId.SetFocus
                Exit Sub
            End If
            
            If Len(CboSupplier) = 0 Then
                MsgBox "Return From Required...", vbInformation, App.title
                CboSupplier.SetFocus
                Exit Sub
            End If
            
            
           
            
             If Val(txtfields(3)) = 0 Then
                MsgBox "Amount Must be More than One", vbInformation, App.title
                CboPurId.SetFocus
                Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT IssueId from IssueSub  WHERE (IssueId= '" & Trim(CboPurId) & "')", objmyCon)
            If objRs.EOF Then
               MsgBox "No such Issue Id exists...Please Verify  issue ID", vbInformation, cmp
               CboPurId.SetFocus
               Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT exp_date,sysdate  from PurchaseSub   WHERE (PurId= '" & CboPurId & "' and itemId='" & Get_Code(cboItem) & "' )", objmyCon)
            
            If Not objRs.EOF Then
             If Not IsNull(objRs(0)) Then
              If Format(objRs(0), "dd/mmm/yy") > Format(objRs(1), "dd/mmm/yy") Then
                 MsgBox "Date Expired..Please Verify", vbInformation, cmp
                 CboPurId.SetFocus
                 Exit Sub
              End If
            End If
                 
            End If
            
            
           
           Set objRs = objcom.Get_RS("SELECT ReturnId from IssueReturnMain   WHERE (ReturnId= '" & txtfields(0) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "Invalid Issue Return Serial...Please Verify.", vbInformation, cmp
               txtfields(0).SelLength = Len(txtfields(0))
               txtfields(0).SetFocus
               Exit Sub
            End If
            
         
          If MsgBox("Are you sure to Delete the Whole Opening Serial Information ", vbYesNo + vbInformation, cmp) = vbYes Then
                  delete
          Else
            Exit Sub
          End If

       MsgBox "Deleted successfully.", vbInformation, cmp
       Call ShowFlexData
       cmdnew.SetFocus
End Sub

Private Sub CmdEdit_Click()
Dim balance As Double
Dim preqty As Double


If txtfields(1).Text = "" Then
          txtfields(1).Text = 0
       End If
       If Len(txtfields(0)) = 0 Then
                MsgBox "Issue Return Serial Mandatory", vbInformation, App.title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(cboItem) = 0 Then
                MsgBox "Please Select an Item", vbInformation, App.title
                cboItem.SetFocus
                Exit Sub
            End If
            
            If Len(CboPurId) = 0 Then
                MsgBox "Please Select an Issue ID", vbInformation, App.title
                CboPurId.SetFocus
                Exit Sub
            End If
            
            If Len(CboSupplier) = 0 Then
                MsgBox "Return From Required...", vbInformation, App.title
                CboSupplier.SetFocus
                Exit Sub
            End If
            
            
           
            
             If Val(txtfields(3)) = 0 Then
                MsgBox "Amount Must be More than One", vbInformation, App.title
                CboPurId.SetFocus
                Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT exp_date,sysdate  from PurchaseSub   WHERE (PurId= '" & CboPurId & "' and itemId='" & Get_Code(cboItem) & "' )", objmyCon)
            
            If Not objRs.EOF Then
             If Not IsNull(objRs(0)) Then
              If Format(objRs(0), "dd/mmm/yy") > Format(objRs(1), "dd/mmm/yy") Then
                 MsgBox "Date Expired..Please Verify", vbInformation, cmp
                 CboPurId.SetFocus
                 Exit Sub
              End If
            End If
                 
            End If
            
            
           
            Set objRs = objcom.Get_RS("SELECT ReturnId from IssueReturnMain   WHERE (ReturnId= '" & txtfields(0) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "Invalid Issue Return Serial...Please Verify.", vbInformation, cmp
               txtfields(0).SelLength = Len(txtfields(0))
               txtfields(0).SetFocus
               Exit Sub
            End If
            
             Set objRs = objcom.Get_RS("SELECT trackid from IssueReturnSub   WHERE (ReturnId = '" & txtfields(0) & "') and trackid='" & txttrackid & "' ", objmyCon)
            
            If objRs.EOF Then
               MsgBox "Invalid Issue Return Serial..Please Verify.", vbInformation, cmp
               Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT IssueId from IssueSub  WHERE (IssueId= '" & Trim(CboPurId) & "')", objmyCon)
           If objRs.EOF Then
              MsgBox "No such Issue Id exists...Please Verify  issue ID", vbInformation, cmp
              CboPurId.SetFocus
              Exit Sub
           End If
           
            Set objRs = objcom.Get_RS("SELECT Qty  from IssueSub  WHERE (IssueId= '" & Trim(CboPurId) & "' and itemCode='" & Trim(Get_Code(cboItem)) & "')", objmyCon)
            If Not objRs.EOF Then
               If Val(objRs(0)) > Val(txtfields(1).Text) Then
                  MsgBox "Quantity Exceeds...Please Verify  issue quantity", vbInformation, cmp
                  txtfields(1).SelLength = Len(txtfields(1).Text)
                  CboPurId.SetFocus
                  Exit Sub
               End If
          End If
          
          Set objRs = objcom.Get_RS("SELECT (PurQty-(UsedQty+ReturnQty)) as balance from PurchaseSub  WHERE (PurId= '" & Trim(CboPurId) & "' and itemId='" & Trim(Get_Code(cboItem)) & "')", objmyCon)
  
            If objRs.EOF Then
               balance = objRs(0)
            End If
            
          Set objRs = objcom.Get_RS("SELECT qty as preqty from IssueSub  WHERE (TrackId = '" & Trim(txttrackid) & "'", objmyCon)
          If objRs.EOF Then
               preqty = objRs(0)
            End If
            
         If (balance + preqty) < Val(txtfields(1).Text) Then
            MsgBox "Insufficient Balance..Please Verify.", vbInformation, cmp
            CboPurId.SetFocus
            Exit Sub
        End If
               
        paramval = Val(txtfields(1).Text) - Val(preqty)
         Set objRs = objcom.Get_RS("SELECT Qty from IssueSub  WHERE (IssueId= '" & Trim(CboPurId) & "' and ItemCode='" & Trim(Get_Code(cboItem)) & "')", objmyCon)
           If Not objRs.EOF Then
               issuedqty = objRs(0)
           End If
           
           
          
           Set objRs = objcom.Get_RS("SELECT nvl(sum(Qty),0)  from IssueReturnSub  WHERE (IssueId= '" & Trim(CboPurId) & "' and ItemCode='" & Trim(Get_Code(cboItem)) & "')", objmyCon)
            If Not objRs.EOF Then
               prevReturn = objRs(0)
          End If
                   
          If issuedqty < (prevReturn + Val(txtfields(1))) Then
             MsgBox "Return Quantity can''t be more than issued quantity....Please Verify", vbInformation, cmp
              CboPurId.SetFocus
             Exit Sub
          End If
        
            edit

       MsgBox "Edited successfully.", vbInformation, cmp
       Call ShowFlexData
       cmdnew.SetFocus
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
        MsgBox "Return Date Mandatory", vbInformation, cmp
        MaskEdBox2.SetFocus
        Exit Sub
    End If
     
     If Len(CboPurType) = 0 Then
        MsgBox "Return Type Required..", vbInformation, cmp
        CboPurType.SetFocus
        Exit Sub
    End If
        
    If Len(CboSupplier) = 0 Then
        MsgBox "Return From whom Required..", vbInformation, cmp
        CboSupplier.SetFocus
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
    
     If mode = 1 Then
        Set objRs = objcom.Get_RS("select to_char(nvl(max(to_number(substr(returnId,7,6))),0)+1,'000000')  from IssueReturnMain where to_char(substr(returnId,4,2))='" & CategoryCode & "'", objmyCon)
         If Not objRs.EOF Then
            txtfields(0) = "IR-" + CategoryCode + "-" + Trim(objRs(0))
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
    
     
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 10, Get_Code(Trim(CboSupplier)))
    cmd.Parameters.Append Param5
     
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 150, txtfields(2).Text)
    cmd.Parameters.Append Param6
    
     Set Param7 = cmd.CreateParameter("param7", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param7
   
     
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param8
     
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_issuereturn_info_main(?,?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
     cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False

End Sub

Private Sub cmdnew_Click()
    txtfields(0) = ""
    txtfields(2) = ""
'    txtfields(5) = ""
    MaskEdBox2 = "__/__/__"
   MaskEdBox2.SetFocus
End Sub

Private Sub cmdSAVE_Click()
       If txtfields(1).Text = "" Then
          txtfields(1).Text = 0
       End If
       If Len(txtfields(0)) = 0 Then
                MsgBox "Issue Return Serial Mandatory", vbInformation, App.title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(cboItem) = 0 Then
                MsgBox "Please Select an Item", vbInformation, App.title
                cboItem.SetFocus
                Exit Sub
            End If
            
            If Len(CboPurId) = 0 Then
                MsgBox "Please Select an Issue ID", vbInformation, App.title
                CboPurId.SetFocus
                Exit Sub
            End If
            
            If Len(CboSupplier) = 0 Then
                MsgBox "Return From Required...", vbInformation, App.title
                CboSupplier.SetFocus
                Exit Sub
            End If
            
            
           
            
             If Val(txtfields(3)) = 0 Then
                MsgBox "Amount Must be More than One", vbInformation, App.title
                CboPurId.SetFocus
                Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT exp_date,sysdate  from PurchaseSub   WHERE (PurId= '" & CboPurId & "' and itemId='" & Get_Code(cboItem) & "' )", objmyCon)
            
            If Not objRs.EOF Then
             If Not IsNull(objRs(0)) Then
              If Format(objRs(0), "dd/mmm/yy") > Format(objRs(1), "dd/mmm/yy") Then
                 MsgBox "Date Expired..Please Verify", vbInformation, cmp
                 CboPurId.SetFocus
                 Exit Sub
              End If
            End If
                 
            End If
            
            
           
            Set objRs = objcom.Get_RS("SELECT ReturnId from IssueReturnMain   WHERE (ReturnId= '" & txtfields(0) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "Invalid Issue Return Serial...Please Verify.", vbInformation, cmp
               txtfields(0).SelLength = Len(txtfields(0))
               txtfields(0).SetFocus
               Exit Sub
            End If
            
             Set objRs = objcom.Get_RS("SELECT IssueId from IssueSub  WHERE (IssueId= '" & Trim(CboPurId) & "')", objmyCon)
             If objRs.EOF Then
                MsgBox "No such Issue Id exists...Please Verify  issue ID", vbInformation, cmp
                CboPurId.SetFocus
                Exit Sub
             End If
           
            
            Set objRs = objcom.Get_RS("SELECT Qty  from IssueSub  WHERE (IssueId= '" & Trim(CboPurId) & "' and itemcode='" & Trim(Get_Code(cboItem)) & "')", objmyCon)
            If Not objRs.EOF Then
               If Val(txtfields(1).Text) > Val(objRs(0)) Then
                  MsgBox "Quantity Exceeds...Please Verify  issue quantity", vbInformation, cmp
                  txtfields(1).SelLength = Len(txtfields(1).Text)
                  CboPurId.SetFocus
                  Exit Sub
               End If
          End If
          
           Set objRs = objcom.Get_RS("SELECT Qty from IssueSub  WHERE (IssueId= '" & Trim(CboPurId) & "' and ItemCode='" & Trim(Get_Code(cboItem)) & "')", objmyCon)
           If Not objRs.EOF Then
               issuedqty = objRs(0)
           End If
           
           
          
           Set objRs = objcom.Get_RS("SELECT nvl(sum(Qty),0)  from IssueReturnSub  WHERE (IssueId= '" & Trim(CboPurId) & "' and ItemCode='" & Trim(Get_Code(cboItem)) & "')", objmyCon)
            If Not objRs.EOF Then
               prevReturn = objRs(0)
          End If
                   
          If issuedqty < (prevReturn + Val(txtfields(1))) Then
             MsgBox "Return Quantity can''t be more than issued quantity....Please Verify", vbInformation, cmp
             CboPurId.SetFocus
             Exit Sub
          End If
          
            save

       MsgBox "Saved successfully.", vbInformation, cmp
       Call ShowFlexData
       cmdnew.SetFocus

End Sub
Private Sub save()
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
    Dim Param15 As New Parameter
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 1)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, Trim(txtfields(0)))
    cmd.Parameters.Append Param2
    
     Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, Trim(CboPurId))
    cmd.Parameters.Append Param3
   
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 10, Trim(Combo1))
    cmd.Parameters.Append Param4
   
    
    Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 10, MaskEdBox2.Text)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 2, Get_Code(CboPurType))
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 4, Get_Code(CboSupplier))
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 5, Get_Code(cboItem))
    cmd.Parameters.Append Param8
   
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 18, Val(txtfields(1)))
    cmd.Parameters.Append Param9
   
        
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, Val(txtfields(4)))
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 150, Trim(txtfields(2)))
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param12
    
    Set Param13 = cmd.CreateParameter("param13", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param13
    
    Set Param14 = cmd.CreateParameter("param14", adDouble, adParamInput, 10, Val(txttrackid))
    cmd.Parameters.Append Param14
 
      
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_issuereturn_info_S_P(?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    'Debug.Print cmd.CommandText
    
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
    Dim Param15 As New Parameter
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 2)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, Trim(txtfields(0)))
    cmd.Parameters.Append Param2
    
     Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, Trim(CboPurId))
    cmd.Parameters.Append Param3
   
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 10, Trim(Combo1))
    cmd.Parameters.Append Param4
   
    
    Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 10, MaskEdBox2.Text)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 2, Get_Code(CboPurType))
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 4, Get_Code(CboSupplier))
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 5, Get_Code(cboItem))
    cmd.Parameters.Append Param8
   
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 18, Val(txtfields(1)))
    cmd.Parameters.Append Param9
   
        
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, Val(txtfields(4)))
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 150, Trim(txtfields(2)))
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param12
    
    Set Param13 = cmd.CreateParameter("param13", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param13
    
    Set Param14 = cmd.CreateParameter("param14", adDouble, adParamInput, 10, Val(txttrackid))
    cmd.Parameters.Append Param14
 
      
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_issuereturn_info_S_P(?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    'Debug.Print cmd.CommandText
    
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
    Dim Param15 As New Parameter
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 3)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, Trim(txtfields(0)))
    cmd.Parameters.Append Param2
    
     Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, Trim(CboPurId))
    cmd.Parameters.Append Param3
   
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 10, Trim(Combo1))
    cmd.Parameters.Append Param4
   
    
    Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 10, MaskEdBox2.Text)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 2, Get_Code(CboPurType))
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 4, Get_Code(CboSupplier))
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 5, Get_Code(cboItem))
    cmd.Parameters.Append Param8
   
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 18, Val(txtfields(1)))
    cmd.Parameters.Append Param9
   
        
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, Val(txtfields(4)))
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 150, Trim(txtfields(2)))
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param12
    
    Set Param13 = cmd.CreateParameter("param13", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param13
    
    Set Param14 = cmd.CreateParameter("param14", adDouble, adParamInput, 10, Val(txttrackid))
    cmd.Parameters.Append Param14
 
      
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_issuereturn_info_S_P(?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    'Debug.Print cmd.CommandText
    
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
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    Dim Param10 As New Parameter
    Dim Param11 As New Parameter
    Dim Param12 As New Parameter
    Dim Param13 As New Parameter
    Dim Param14 As New Parameter
    Dim Param15 As New Parameter
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 4)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, Trim(txtfields(0)))
    cmd.Parameters.Append Param2
    
     Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, Trim(CboPurId))
    cmd.Parameters.Append Param3
   
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 10, Trim(Combo1))
    cmd.Parameters.Append Param4
   
    
    Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 10, MaskEdBox2.Text)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 2, Get_Code(CboPurType))
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 4, Get_Code(CboSupplier))
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 5, Get_Code(cboItem))
    cmd.Parameters.Append Param8
   
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 18, Val(txtfields(1)))
    cmd.Parameters.Append Param9
   
        
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, Val(txtfields(4)))
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 150, Trim(txtfields(2)))
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param12
    
    Set Param13 = cmd.CreateParameter("param13", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param13
    
    Set Param14 = cmd.CreateParameter("param14", adDouble, adParamInput, 10, Val(txttrackid))
    cmd.Parameters.Append Param14
 
      
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_issuereturn_info_S_P(?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    'Debug.Print cmd.CommandText
    
     cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
         
End Sub
Private Sub dtpic_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdsave.SetFocus
End If
End Sub

Private Sub Combo1_Change()
 Call load_qty
End Sub

Private Sub Combo1_Click()
  Call load_qty
End Sub

Private Sub Command1_Click(Index As Integer)
  Dim RS As New ADODB.Recordset
Select Case Index
    Case 1
       If txtfields(1).Text = "" Then
          txtfields(1).Text = 0
       End If
       If Len(txtfields(0)) = 0 Then
                MsgBox "Issue Return Serial Mandatory", vbInformation, App.title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(cboItem) = 0 Then
                MsgBox "Please Select an Item", vbInformation, App.title
                cboItem.SetFocus
                Exit Sub
            End If
            
            If Len(CboPurId) = 0 Then
                MsgBox "Please Select an Issue ID", vbInformation, App.title
                CboPurId.SetFocus
                Exit Sub
            End If
            
            If Len(CboSupplier) = 0 Then
                MsgBox "Return From Required...", vbInformation, App.title
                CboSupplier.SetFocus
                Exit Sub
            End If
            
            
           
            
             If Val(txtfields(3)) = 0 Then
                MsgBox "Amount Must be More than One", vbInformation, App.title
                CboPurId.SetFocus
                Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT ReturnId from IssueReturnsub   WHERE (ReturnId= '" & txtfields(0) & "') and IssueId= '" & Trim(CboPurId) & "' and  ItemCode='" & Get_Code(cboItem) & "'", objmyCon)
            
            If Not objRs.EOF Then
               Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT exp_date,sysdate  from PurchaseSub   WHERE (PurId= '" & CboPurId & "' and itemId='" & Get_Code(cboItem) & "' )", objmyCon)
            
            If Not objRs.EOF Then
             If Not IsNull(objRs(0)) Then
              If Format(objRs(0), "dd/mmm/yy") > Format(objRs(1), "dd/mmm/yy") Then
                 MsgBox "Date Expired..Please Verify", vbInformation, cmp
                 CboPurId.SetFocus
                 Exit Sub
              End If
            End If
                 
            End If
            
            
           
            Set objRs = objcom.Get_RS("SELECT ReturnId from IssueReturnMain   WHERE (ReturnId= '" & txtfields(0) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "Invalid Issue Return Serial...Please Verify.", vbInformation, cmp
               txtfields(0).SelLength = Len(txtfields(0))
               txtfields(0).SetFocus
               Exit Sub
            End If
            
            
           Set objRs = objcom.Get_RS("SELECT IssueId from IssueSub  WHERE (IssueId= '" & Trim(CboPurId) & "')", objmyCon)
           If objRs.EOF Then
              MsgBox "No such Issue Id exists...Please Verify  issue ID", vbInformation, cmp
              CboPurId.SetFocus
              Exit Sub
           End If
           
            
            
            
            Set objRs = objcom.Get_RS("SELECT Qty  from IssueSub  WHERE (IssueId= '" & Trim(CboPurId) & "' and ItemCode='" & Trim(Get_Code(cboItem)) & "')", objmyCon)
            If Not objRs.EOF Then
               If Val(txtfields(1).Text) > Val(objRs(0)) Then
                  MsgBox "Quantity Exceeds...Please Verify  issue quantity", vbInformation, cmp
                  txtfields(1).SelLength = Len(txtfields(1).Text)
                  CboPurId.SetFocus
                  Exit Sub
               End If
          End If
          
          
          Set objRs = objcom.Get_RS("SELECT Qty from IssueSub  WHERE (IssueId= '" & Trim(CboPurId) & "' and ItemCode='" & Trim(Get_Code(cboItem)) & "')", objmyCon)
           If Not objRs.EOF Then
               issuedqty = objRs(0)
           End If
           
           
          
           Set objRs = objcom.Get_RS("SELECT nvl(sum(Qty),0)  from IssueReturnSub  WHERE (IssueId= '" & Trim(CboPurId) & "' and ItemCode='" & Trim(Get_Code(cboItem)) & "')", objmyCon)
            If Not objRs.EOF Then
               prevReturn = objRs(0)
          End If
                   
          If issuedqty < (prevReturn + Val(txtfields(1))) Then
             MsgBox "Return Quantity can''t be more than issued quantity....Please Verify", vbInformation, cmp
             CboPurId.SetFocus
             Exit Sub
          End If
        save

      
       Call ShowFlexData
       txtfields(1).Text = 0
       txtfields(4).Text = 0
       cboItem.SetFocus
  End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
'      SendKeys (Chr(9))
   Else
       If KeyAscii = 27 Then
          Unload Me
       End If
       
   
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
    .ColWidth(6) = 1100
    .ColWidth(7) = 2350
    .ColWidth(8) = 0
    

    
End With

Call ShowFlexData
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title

End Sub

Private Sub load_item()
 Set objRs = objcom.Get_RS("SELECT item_code ,item_name from item_info where group_code='" & Get_Code(CategoryCode) & "' order by item_code ", objmyCon)
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
 Set objRs = objcom.Get_RS("SELECT group_code,group_name from item_group_info  order by group_code", objmyCon)
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
 Set objRs = objcom.Get_RS("select issueid, PurchaseId  from issueSub  where ItemCode = '" & Get_Code(cboItem) & "' and Qty > 0 order by issueid", objmyCon)
 CboPurId.Clear
 Combo1.Clear
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       CboPurId.AddItem objRs(0)
       Combo1.AddItem objRs(1)
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
    MaskEdBox2.SelStart = 0
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
               CboPurType.SetFocus
            End If
   Else
       CboPurType.SetFocus
  End If
 End If
End Sub

Private Sub mnuDL_Click()
        If txtfields(1).Text = "" Then
          txtfields(1).Text = 0
       End If
       If Len(txtfields(0)) = 0 Then
                MsgBox "Issue Return Serial Mandatory", vbInformation, App.title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(cboItem) = 0 Then
                MsgBox "Please Select an Item", vbInformation, App.title
                cboItem.SetFocus
                Exit Sub
            End If
            
            If Len(CboPurId) = 0 Then
                MsgBox "Please Select an Issue ID", vbInformation, App.title
                CboPurId.SetFocus
                Exit Sub
            End If
            
            If Len(CboSupplier) = 0 Then
                MsgBox "Return From Required...", vbInformation, App.title
                CboSupplier.SetFocus
                Exit Sub
            End If
            
            
           
            
             If Val(txtfields(3)) = 0 Then
                MsgBox "Amount Must be More than One", vbInformation, App.title
                CboPurId.SetFocus
                Exit Sub
            End If
            Set objRs = objcom.Get_RS("SELECT ReturnId from IssueReturnMain   WHERE (ReturnId= '" & txtfields(0) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "Invalid Issue Return Serial...Please Verify.", vbInformation, cmp
               txtfields(0).SelLength = Len(txtfields(0))
               txtfields(0).SetFocus
               Exit Sub
            End If
           
             
            Set objRs = objcom.Get_RS("SELECT trackid from IssueReturnSub   WHERE (ReturnId = '" & txtfields(0) & "') and trackid='" & txttrackid & "' ", objmyCon)
            
            If objRs.EOF Then
               MsgBox "Invalid Issue Return Serial..Please Verify.", vbInformation, cmp
               Exit Sub
            End If
            
            popup_delete

      
       Call ShowFlexData
       cmdnew.SetFocus
  
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
      PopupMenu MnuDelete, 2
   End If
End Sub

Private Sub MSFlexGrid2_DblClick()
   If Len(Trim(MSFlexGrid2.Text)) <> 0 Then
      cboItem.Text = MSFlexGrid2.Text
      Set objRs = objcom.Get_RS("select item_name from item_info where to_number(item_code)='" & Trim(cboItem.Text) & "'", objmyCon)
      If Not objRs.EOF Then
        txtItemTitle.Text = objRs(0)
      End If
'       Call cboItem_LostFocus
'       txtItemTitle.SetFocus
       'nbrDollar.SetFocus
'    Else
'
'       Exit Sub
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
               Set objRs = objcom.Get_RS("SELECT  ReturnDate,ReturnType,ReturnFrom,Remarks from IssueReturnMain  WHERE (ReturnId= '" & txtfields(0) & "')", objmyCon)
            
              If Not objRs.EOF Then
                 MaskEdBox2 = Format(objRs(0), "dd/mm/yy")
'                 CboPurType = select_pur_tpy(objRs(1))
'                 txtfields(5).Text = "" & objRs(3)
'                 txtfields(2).Text = "" & objRs(4)
              End If
          
              
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
     End Select
     
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
        Case 0
         If Len(txtfields(0).Text) > 0 Then
                CmdGenerate.Enabled = False
               Else
                CmdGenerate.Enabled = True
               End If
             If Len(txtfields(0)) <> 0 Then
              If Mid(txtfields(0), 1, 2) <> "IR" Then
                   txtfields(0).Text = Format(txtfields(0), "000000")
                   txtfields(0).Text = "IR-" + CategoryCode + "-" + txtfields(0).Text
                End If
              End If
                   ShowFlexData
               Set objRs = objcom.Get_RS("SELECT  a.IssueDate,a.IssueType,b.type_name,a.Issueto,a.Remarks from IssueMain a,item_issue_type b WHERE (a.issueId= '" & txtfields(0) & "' and a.IssueType=b.type_code)", objmyCon)
            
              If Not objRs.EOF Then
                 MaskEdBox2 = Format(objRs(0), "dd/mm/yy")
                 CboPurType = Trim(objRs(2)) + "~" + Trim(objRs(1))
                
'                 txtfields(2).Text = "" & objRs(4)
              End If
              MaskEdBox2.SetFocus
              
  Case 1
       Command1(1).SetFocus
  Case 2
        If CmdGenerate.Enabled = False Then
           cboItem.SetFocus
        Else
           CmdGenerate.SetFocus
        End If
  End Select
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
Set RS = objcom.Get_RS("SELECT c.group_code, c.group_name,a.Itemcode  as Code ,b.item_name,a.issueId,a.URate,a.Qty,a.TrackId   From issuereturnSub a ,item_info b, item_group_info c where to_number(b.group_code) =to_number( c.group_code) and a.ReturnId='" & txtfields(0).Text & "' and to_number(b.item_code)=to_number(a.ItemCode) order by a.ItemCode ", objmyCon)
If Not RS.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until RS.EOF
            MSFlexGrid1.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = Trim(RS(0))
                .TextMatrix(i, 1) = Trim(RS(1))
                .TextMatrix(i, 2) = RS(2)
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
If MSFlexGrid1.Row > 1 Then
    Cbocategory = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) + "~" + Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0))
    cboItem = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)) + "~" + Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
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
Private Sub txtItemTitle_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     CboPurId.SetFocus
  End If
End Sub
