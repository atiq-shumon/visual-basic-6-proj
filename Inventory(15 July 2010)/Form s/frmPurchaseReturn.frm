VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmPurchaseReturn 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   2430
      Left            =   60
      TabIndex        =   39
      Top             =   3090
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
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   1680
      Left            =   -30
      TabIndex        =   27
      Top             =   720
      Width           =   11970
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1485
         MaxLength       =   12
         TabIndex        =   1
         ToolTipText     =   "Purchase Serial"
         Top             =   165
         Width           =   1755
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   2
         Left            =   1485
         MaxLength       =   100
         TabIndex        =   4
         ToolTipText     =   "Insert Remarks Category"
         Top             =   1035
         Width           =   7365
      End
      Begin VB.CommandButton Command1 
         Caption         =   ":::"
         Height          =   285
         Index           =   0
         Left            =   3240
         TabIndex        =   28
         Top             =   180
         Width           =   375
      End
      Begin VB.ComboBox CboPurType 
         Height          =   315
         ItemData        =   "frmPurchaseReturn.frx":0000
         Left            =   1485
         List            =   "frmPurchaseReturn.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   577
         Width           =   1755
      End
      Begin VB.ComboBox CboSupplier 
         Height          =   315
         ItemData        =   "frmPurchaseReturn.frx":0030
         Left            =   5430
         List            =   "frmPurchaseReturn.frx":0032
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   577
         Width           =   5385
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   285
         Left            =   5430
         TabIndex        =   0
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
         Left            =   8850
         TabIndex        =   5
         ToolTipText     =   "Press to Generate Purchase Return Serial"
         Top             =   1035
         Width           =   1995
         _ExtentX        =   3519
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
         MICON           =   "frmPurchaseReturn.frx":0034
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
         Caption         =   "Purchase Returnl#"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   33
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   32
         Top             =   1065
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Return Date "
         Height          =   195
         Index           =   1
         Left            =   4110
         TabIndex        =   31
         Top             =   210
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Return Type"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   30
         Top             =   637
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
         Height          =   195
         Index           =   3
         Left            =   4095
         TabIndex        =   29
         Top             =   637
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000B&
      Height          =   705
      Left            =   0
      TabIndex        =   21
      Top             =   2400
      Width           =   11940
      Begin VB.TextBox txtItemTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAF2C8&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   360
         Width           =   4485
      End
      Begin VB.TextBox cboItem 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   360
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0097C8E6&
         Caption         =   ">>"
         Height          =   315
         Index           =   1
         Left            =   11520
         MaskColor       =   &H00C0C000&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
      Begin VB.ComboBox ComboPurId 
         Height          =   315
         Left            =   5910
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   360
         Width           =   1710
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   10590
         TabIndex        =   10
         Top             =   360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14210769
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8D6D1&
         Height          =   315
         Index           =   3
         Left            =   9525
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   9
         Text            =   "0"
         ToolTipText     =   "Insert Marks Category"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   7620
         MaxLength       =   15
         TabIndex        =   7
         Text            =   "0"
         ToolTipText     =   "Insert Quantity"
         Top             =   360
         Width           =   825
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8D6D1&
         Height          =   315
         Index           =   4
         Left            =   8430
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "0"
         ToolTipText     =   "Insert Unit Rate  "
         Top             =   360
         Width           =   1125
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
         TabIndex        =   38
         Top             =   150
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
         TabIndex        =   37
         Top             =   150
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase ID"
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
         Left            =   6060
         TabIndex        =   35
         Top             =   120
         Width           =   945
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
         Left            =   10620
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
         Left            =   9540
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
         Left            =   7680
         TabIndex        =   23
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
         Left            =   8490
         TabIndex        =   22
         Top             =   120
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   645
      Left            =   0
      TabIndex        =   18
      Top             =   7380
      Width           =   11925
      Begin VB.TextBox txttrackid 
         Height          =   285
         Left            =   1200
         TabIndex        =   26
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
         Left            =   8850
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
         Left            =   6915
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
         Left            =   7890
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
         Left            =   9840
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
         Left            =   10815
         TabIndex        =   15
         ToolTipText     =   "Click to Close"
         Top             =   180
         Width           =   945
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   6885
         Top             =   150
         Width           =   4905
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4305
      Left            =   0
      TabIndex        =   17
      Top             =   3120
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   7594
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   13627123
      ForeColor       =   12582912
      BackColorSel    =   12640511
      ForeColorSel    =   8421631
      BackColorBkg    =   -2147483637
      GridColor       =   -2147483637
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000001&
      FillColor       =   &H00E0E0E0&
      Height          =   795
      Left            =   -30
      ScaleHeight     =   735
      ScaleWidth      =   11895
      TabIndex        =   16
      Top             =   -30
      Width           =   11955
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Return Entry"
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
         Left            =   3975
         TabIndex        =   20
         Top             =   120
         Width           =   3495
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
Attribute VB_Name = "frmPurchaseReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objcom As New DSLComFram.clsCommon
Dim objRs As New ADODB.Recordset
Dim NewTotalReturn As Double
Dim mode As Integer

Private Sub generate(mode As Integer)
'On Error GoTo errdes
Dim cmd As New ADODB.Command

Dim Param1 As New Parameter
Dim Param2 As New Parameter
Dim Param3 As New Parameter
Dim Param4 As New Parameter
Dim Param5 As New Parameter
Dim Param6 As New Parameter
Dim Param7 As New Parameter

If mode = 1 Then
   Set objRs = objcom.Get_RS("select to_char(nvl(max(to_number(substr(PurReturnId,7,6))),0)+1,'000000')  from PurchaseReturnMain where to_char(substr(PurReturnId,4,2))='" & CategoryCode & "'", objmyCon)
   If Not objRs.EOF Then
        txtfields(0) = "PR-" + CategoryCode + "-" + Trim(objRs(0))
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
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 100, txtfields(2))
    cmd.Parameters.Append Param6
     
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param7
   
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_PurchaseReturn_Main(?,?,?,?,?,?,?)}"
    
    'Debug.Print cmd.CommandText
    cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title
End Sub

Private Sub Cbocategory_Click()
   Call load_item
End Sub

Private Sub Cbocategory_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cboItem.SetFocus
End If
End Sub

Private Sub Cbocategory_LostFocus()
Call Cbocategory_Click
End Sub

Private Sub CboItem_Click()
'On Error GoTo errdes
 
Set objRs = objcom.Get_RS("SELECT a.PurId from PurchaseMain a, PurchaseSub b Where a.purid=b.purid and a.Supplierid = '" & Get_Code(CboSupplier) & "' AND b.itemId = '" & Get_Code(cboItem) & "' and (b.PurQty-(b.UsedQty+b.ReturnQty)) > 0 order by b.PurId", objmyCon)
ComboPurId.Clear
If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       ComboPurId.AddItem Trim(objRs(0))
       objRs.MoveNext
    Loop
End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title
End Sub

Private Sub cboItem_GotFocus()
   cboItem.SelStart = 0
  cboItem.SelLength = Trim(Len(cboItem))
End Sub

Private Sub CboItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    ComboPurId.SetFocus
End If
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
        ComboPurId.SetFocus
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

Private Sub CboPurType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CboSupplier.SetFocus
End If
End Sub

Private Sub CboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtfields(2).SetFocus
End If
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Sure to Delete", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
If Len(txtfields(0)) = 0 Then
    MsgBox "Return Serial Mandatory", vbInformation, App.title
    txtfields(0).SetFocus
    Exit Sub
End If

mode = 3
Call popup_delete

Call ShowFlexData
txtfields(1).Text = 0
txtfields(4).Text = 0
txttrackid = ""
txtfields(0) = ""
cboItem.Clear
Call ShowFlexData
cmdnew.SetFocus
End Sub

Private Sub CmdEdit_Click()
'On Error GoTo errdes
Dim PreValance As Double
Dim PreReturnQty As Double

If Len(txtfields(0)) = 0 Then
    MsgBox "Return Serial Mandatory", vbInformation, App.title
    txtfields(0).SetFocus
    Exit Sub
End If
            
            
If Len(cboItem) = 0 Then
    MsgBox "Please Select an Item", vbInformation, App.title
    cboItem.SetFocus
    Exit Sub
End If
            
If MaskEdBox2.Text = "__/__/__" Then
    MsgBox "Return Date Required", vbInformation, App.title
    MaskEdBox2.SetFocus
    Exit Sub
End If
                       
Set objRs = objcom.Get_RS("SELECT * from PurchaseReturnMain  WHERE (PURRETURNID= '" & txtfields(0) & "')", objmyCon)

If objRs.EOF Then
   MsgBox "Invalid Return Serial..Please Verify.", vbInformation, cmp
   CmdGenerate.SetFocus
   Exit Sub
End If

'Set objRs = objcom.Get_RS("SELECT * from PurchaseReturnSub  WHERE (PURRETURNID= '" & txtfields(0) & "' and TrackId ='" & txttrackid & "')", objmyCon)
'If objRs.EOF Then
'   MsgBox "No Such Item available..Please Verify.", vbInformation, cmp
'   MSFlexGrid1.SetFocus
'   Exit Sub
'End If

'Set objRs = objcom.Get_RS("SELECT PurQty - (UsedQty + ReturnQty) FROM PurchaseSub Where Purid ='" & ComboPurId & "' and itemId = '" & Get_Code(cboItem) & "'", objmyCon)
'If objRs(0) <> 0 Then
'    PreValance = objRs(0)
'Else
'    PreValance = 0
'End If
'
'Set objRs = objcom.Get_RS("SELECT RETURNQTY FROM PurchaseReturnSub Where PURRETURNID ='" & txtfields(0) & "' and ITEMCODE = '" & Get_Code(cboItem) & "'", objmyCon)
'If objRs(0) <> 0 Then
'    PreReturnQty = objRs(0)
'Else
'    PreReturnQty = 0
'End If
'
'NewTotalReturn = Val(txtfields(1)) - PreReturnQty
'If (PreValance + PreReturnQty) < Val(txtfields(1)) Then
'    MsgBox "Insufficient Balance", vbInformation, App.title
'    txtfields(1).SetFocus
'    Exit Sub
'End If
'
'mode = 2
'Call popup_delete
'
'Call ShowFlexData
'txtfields(1).Text = 0
'txtfields(4).Text = 0
'txttrackid = ""
'cboItem.Clear
'Call ShowFlexData
 generate (2)
MsgBox "Updated Successfully", vbInformation, cmp
cmdnew.SetFocus
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title
End Sub

Private Sub cmdExit_Click()
 Unload Me
End Sub

Private Sub CmdGenerate_Click()
'On Error GoTo errdes
    'Set objRs = objcom.Get_RS("SELECT PURRETURNID from PurchaseReturnMain WHERE PURRETURNID = '" & txtfields(0) & "'", objmyCon)
    
    If Len(txtfields(0)) <> 0 Then
       MsgBox "Serial No already exists...Please Blank the field to Generate", vbInformation, cmp
       txtfields(0).SetFocus
       Exit Sub
     End If
     
     If MaskEdBox2.Text = "__/__/__" Then
        MsgBox "Return Date Required", vbInformation, App.title
        MaskEdBox2.SetFocus
        Exit Sub
    End If
     
     If Len(CboPurType) = 0 Then
        MsgBox "Purchase Type Required..", vbInformation, cmp
        CboPurType.SetFocus
        Exit Sub
    End If
        
    If Len(CboSupplier) = 0 Then
        MsgBox "Supplier Required..", vbInformation, cmp
        CboSupplier.SetFocus
        Exit Sub
    End If
    Call generate(1)
    cboItem.SetFocus
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title
End Sub

Private Sub cmdnew_Click()
txtfields(0) = ""
MaskEdBox2.SetFocus
End Sub

Private Sub ComboPurId_Click()
'On Error GoTo errdes
Set objRs = objcom.Get_RS("SELECT PurQty - (UsedQty + ReturnQty), URate, exp_date from PurchaseSub Where Purid ='" & ComboPurId & "' and itemId = '" & Get_Code(cboItem) & "'", objmyCon)
If objRs.EOF = False Then
    txtfields(1) = objRs(0)
    txtQty = objRs(0)
    txtfields(4) = objRs(1)
    If IsNull(objRs(2)) = False Then
        MaskEdBox1.Text = Format(objRs(2), "dd/mm/yy")
    Else
        MaskEdBox1.Text = "__/__/__"
    End If
Else
    txtfields(1) = ""
    txtfields(4) = ""
    MaskEdBox1.Text = "__/__/__"
    txtQty = ""
End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title
End Sub

Private Sub ComboPurId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtfields(1).SetFocus
End If
End Sub

Private Sub Command1_Click(Index As Integer)
   Select Case Index
          Case 1
             'On Error GoTo errdes
  If Len(txtfields(0)) = 0 Then
        MsgBox "Return Serial Mandatory", vbInformation, App.title
        txtfields(0).SetFocus
        Exit Sub
    End If
            
            
     If Len(cboItem) = 0 Then
         MsgBox "Please Select an Item", vbInformation, App.title
         cboItem.SetFocus
         Exit Sub
     End If
     
     If Len(ComboPurId) = 0 Then
        MsgBox "Purchase ID Required", vbInformation, App.title
        ComboPurId.SetFocus
        Exit Sub
     End If
     
     
     
    If Val(txtfields(3)) = 0 Then
       MsgBox "Amount Must be More than Zero(0)", vbInformation, App.title
       txtfields(1).SetFocus
       Exit Sub
     End If
     
    
     Set objRs = objcom.Get_RS("SELECT * from PurchaseReturnMain  WHERE (PURRETURNID= '" & txtfields(0) & "')", objmyCon)
     
     If objRs.EOF Then
        MsgBox "Invalid Return Serial..Please Verify.", vbInformation, cmp
        txtfields(0).SelLength = Len(txtfields(0))
        txtfields(0).SetFocus
        Exit Sub
     End If
    Set objRs = objcom.Get_RS("SELECT PurQty - (UsedQty + ReturnQty) from PurchaseSub Where Purid ='" & ComboPurId & "' and itemId = '" & Get_Code(cboItem) & "'", objmyCon)
    
    If Val(txtfields(1)) > objRs(0) Then
        MsgBox "Quantity exceeds...please Verify.", vbInformation, cmp
        txtfields(1).SetFocus
        Exit Sub
    End If
    
    Call save

    Call ShowFlexData
    txtfields(1).Text = 0
    txtfields(4).Text = 0
    txttrackid = ""
'    cboItem.Clear
    Call Cbocategory_Click
    CboItem_Click
    cboItem.SetFocus
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title
   End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
'      SendKeys (Chr(9))
  End If
   ''Else
       If KeyAscii = 27 Then
          Unload Me
       End If
  'End If
End Sub

Private Sub Form_Load()

Call FillGrid
load_supplier
MaskEdBox2.Text = Format(Date, "dd/mm/yy")
End Sub
Private Sub load_supplier()
 Set objRs = objcom.Get_RS("SELECT acc_code,ACC_NAME  from acct_10_11.acct where  acc_code not in (select acc_head from acct_10_11.acct) order by acc_code", objmyCon)
 CboSupplier.Clear
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       CboSupplier.AddItem Trim(objRs(1)) + "~" + Trim(objRs(0))
       objRs.MoveNext
    Loop
 End If
End Sub

'Private Sub MaskEdBox1_Change()
'MsgBox "Invalid Operation", vbInformation, App.Title
'Exit Sub
'End Sub

Private Sub MaskEdBox1_GotFocus()
   MaskEdBox1.SelStart = 0
   MaskEdBox1.SelLength = Len(MaskEdBox1)
End Sub

Private Sub MaskEdBox1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Command1(1).SetFocus
End If
End Sub

Private Sub MaskEdBox2_GotFocus()
MaskEdBox2.SelLength = Len(MaskEdBox2)
End Sub

Private Sub MaskEdBox2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CboPurType.SetFocus
End If
End Sub

Private Sub mnuDL_Click()
'On Error GoTo errdes
If Len(txtfields(0)) = 0 Then
    MsgBox "Return Serial Mandatory", vbInformation, App.title
    txtfields(0).SetFocus
    Exit Sub
End If
            
            
If Len(cboItem) = 0 Then
    MsgBox "Please Select an Item", vbInformation, App.title
    cboItem.SetFocus
    Exit Sub
End If
            
                       
Set objRs = objcom.Get_RS("SELECT * from PurchaseReturnMain  WHERE (PURRETURNID= '" & txtfields(0) & "')", objmyCon)

If objRs.EOF Then
   MsgBox "Invalid Return Serial..Please Verify.", vbInformation, cmp
   CmdGenerate.SetFocus
   Exit Sub
End If

Set objRs = objcom.Get_RS("SELECT * from PurchaseReturnSub  WHERE (PURRETURNID= '" & txtfields(0) & "' and TrackId ='" & txttrackid & "')", objmyCon)
If objRs.EOF Then
   MsgBox "No Such Item available..Please Verify.", vbInformation, cmp
   MSFlexGrid1.SetFocus
   Exit Sub
End If

If Len(txttrackid) = 0 Then
    MsgBox "Please click on the grid to select a PR item", vbInformation, App.title
    cboItem.SetFocus
    Exit Sub
End If
          


 
mode = 1
Call popup_delete


txtfields(1).Text = 0
txtfields(4).Text = 0
txttrackid = ""
CboItem_Click
Call ShowFlexData
'cmdnew.SetFocus
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title
End Sub

Private Sub MSFlexGrid1_Click()
'On Error GoTo errdes
If MSFlexGrid1.Rows > 1 Then
'     Cbocategory = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) + "~" + Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0))
     cboItem = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
     txtItemTitle = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3))
     'ComboPurId = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
     txtfields(1).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
     txtfields(4).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
    MaskEdBox1.Text = IIf(Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8), "dd/mm/yy") = "", Trim("__/__/__"), Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8), "dd/mm/yy"))
    txttrackid.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9)
     txtfields(3) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)
End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title
End Sub

Private Sub MSFlexGrid2_DblClick()
   If Len(Trim(MSFlexGrid2.Text)) <> 0 Then
      cboItem.Text = MSFlexGrid2.Text
      Set objRs = objcom.Get_RS("select item_name from item_info where to_number(item_code)='" & Trim(cboItem.Text) & "'", objmyCon)
      If Not objRs.EOF Then
        txtItemTitle.Text = objRs(0)
      End If
       ComboPurId.SetFocus
    End If
    MSFlexGrid2.Visible = False
    CboItem_Click
   
End Sub

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      MSFlexGrid2_DblClick
   End If
End Sub

Private Sub txtfields_Change(Index As Integer)
'On Error GoTo errdes
Select Case Index
    Case 0
           If Len(txtfields(0).Text) > 0 Then
                  CmdGenerate.Enabled = False
               Else
                 CmdGenerate.Enabled = True
               End If
     Case 4
        If Not IsNumeric(txtfields(4)) Then
            txtfields(4) = ""
        Else
           txtfields(3) = txtfields(1) * txtfields(4)
        End If
    Case 1
        If Not IsNumeric(txtfields(1)) Then
            txtfields(1) = ""
        Else
           txtfields(3) = Val(txtfields(1)) * Val(txtfields(4))
        End If
End Select
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title
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


Private Sub load_item()
'On Error GoTo errdes
 Set objRs = objcom.Get_RS("SELECT item_code ,item_name from item_info where group_code='" & Get_Code(CategoryCode) & "'", objmyCon)
' cboItem.Clear
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       cboItem.AddItem Trim(objRs(1)) + "~" + Trim(objRs(0))
       objRs.MoveNext
    Loop
 End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title
End Sub

Private Sub save()
'On Error GoTo errdes
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

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 1)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 14, Trim(txtfields(0)))
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 14, ComboPurId.Text)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 4, Get_Code(cboItem))
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 10, Val(txtfields(1)))
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 10, Val(txtfields(4)))
    cmd.Parameters.Append Param6
   
    Set Param7 = cmd.CreateParameter("param7", adDate, adParamInput, 15, IIf(Format(MaskEdBox1.Text, "dd-mon-yyyy'") = "__/__/__", Null, Format(MaskEdBox1.Text, "dd/mm/yy")))
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adDouble, adParamInput, 10, Val(txttrackid))
    cmd.Parameters.Append Param8
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_PurchaseReturn_Sub(?,?,?,?,?,?,?,?)}"
    
    'Debug.Print cmd.CommandText
    
    cmd.Execute
    cmd.Properties("PLSQLRSet") = False
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title
End Sub

Private Sub FillGrid()
With MSFlexGrid1
     .Rows = 1
    .Cols = 10
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
    .ColWidth(3) = 4500
    .ColWidth(4) = 1680
    .ColWidth(5) = 850
    .ColWidth(6) = 1100
    .ColWidth(7) = 2200
    .ColWidth(8) = 0
    .ColWidth(9) = 0
  End With
End Sub

Private Sub ShowFlexData()
'On Error GoTo errdes
Dim RS As New ADODB.Recordset
Set RS = objcom.Get_RS("SELECT c.cate_code, c.cate_name,a.ITEMCODE,b.item_name,a.PURID,a.RETURNQTY,a.RETURRATE,a.exp_date,a.TrackId   From PurchaseReturnSub a ,item_info b, item_cate_info c where b.cate_code = c.cate_code and a.PURRETURNID='" & txtfields(0).Text & "' and to_number(b.item_code)=to_number(a.ITEMCODE) order by a.ITEMCODE ", objmyCon)
If Not RS.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until RS.EOF
            MSFlexGrid1.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = Trim(RS(0))
                .TextMatrix(i, 1) = Trim(RS(1))
                .TextMatrix(i, 2) = Trim(RS(2))
                .TextMatrix(i, 3) = Trim(RS(3))
                .TextMatrix(i, 4) = Trim(RS(4))
                .TextMatrix(i, 5) = Trim(RS(5))
                .TextMatrix(i, 6) = Trim(RS(6))
                .TextMatrix(i, 7) = RS(5) * RS(6)
                .TextMatrix(i, 8) = "" & Trim(RS(7))
                .TextMatrix(i, 9) = Trim(RS(8))
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
Private Sub txtfields_GotFocus(Index As Integer)
            txtfields(Index).SelStart = 0
            txtfields(Index).SelLength = Len(txtfields(Index))
 
End Sub

Private Sub txtfields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Select Case Index
    Case 0
        MaskEdBox2.SetFocus
    
    Case 1
        Command1(1).SetFocus
End Select
End If
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
         If Len(txtfields(0).Text) > 0 Then
             If Mid(txtfields(0), 1, 2) <> "PR" Then
                 txtfields(0).Text = Format(txtfields(0), "000000")
                 txtfields(0).Text = "PR-" + CategoryCode + "-" + txtfields(0).Text
             End If
          End If
          ShowFlexData
          Set objRs = objcom.Get_RS("SELECT  PurReturnDate,PurReturnType,SupplierId,Remarks from PurchaseReturnMain  WHERE (PurReturnId = '" & txtfields(0) & "')", objmyCon)
                               
          If Not objRs.EOF Then
                 MaskEdBox2 = Format(objRs(0), "dd/mm/yy")
                 CboPurType = select_pur_tpy(objRs(1))
                 CboSupplier = Trim(select_acct(objRs(2)))
                 txtfields(2).Text = "" & objRs(3)
                
              Else
               
               txtfields(2) = ""
              
               MaskEdBox2 = "__/__/__"
              
             End If
             MaskEdBox2.SetFocus
             
    Case 2
          If CmdGenerate.Enabled = False Then
             cboItem.SetFocus
          Else
             CmdGenerate.SetFocus
          End If
    End Select
    
  End If
End Sub
Private Function select_acct(str As String) As String
        Dim loacalRs As New ADODB.Recordset
        Set loacalRs = objcom.Get_RS("SELECT acc_name,acc_code from acct.acct  WHERE (acc_code= '" & str & "')", objmyCon)
            
              If Not loacalRs.EOF Then
                select_acct = Trim(loacalRs(0)) + "~" + Trim(loacalRs(1))
              End If
End Function
Private Sub txtFields_LostFocus(Index As Integer)
'Select Case Index
'    Case 0
'        If Len(Trim(txtfields(0))) = 0 Then Exit Sub
'        If Left(txtfields(0), 2) <> "PR" Then txtfields(0) = "PR-" & Format(Val(txtfields(0)), "0000")
'End Select
End Sub

Private Sub popup_delete()
'On Error GoTo errdes
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

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, mode)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 14, Trim(txtfields(0)))
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 14, MaskEdBox2.Text)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 1, Get_Code(CboPurType.Text))
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 10, Get_Code(Trim(CboSupplier)))
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 100, txtfields(2))
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 12, ComboPurId.Text)
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 4, Get_Code(cboItem))
    cmd.Parameters.Append Param8
    
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 10, Val(txtfields(1)))
    cmd.Parameters.Append Param9
    
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, Val(txtfields(4)))
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("param11", adDate, adParamInput, 15, IIf(Format(MaskEdBox1.Text, "dd-mon-yyyy'") = "__/__/__", Null, Format(MaskEdBox1.Text, "dd/mm/yy")))
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adDouble, adParamInput, 10, Val(txttrackid))
    cmd.Parameters.Append Param12
    
    Set Param13 = cmd.CreateParameter("param13", adDouble, adParamInput, 10, NewTotalReturn)
    cmd.Parameters.Append Param13
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_PurchaseReturn_Edit(?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
      PopupMenu MnuDelete, 2
   End If
End Sub
