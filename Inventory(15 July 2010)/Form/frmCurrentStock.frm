VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCurrentStock 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Current Stock Information"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5715
      Index           =   4
      Left            =   -30
      TabIndex        =   9
      Top             =   1830
      Width           =   9975
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5595
         Left            =   30
         TabIndex        =   10
         Top             =   150
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   9869
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColor       =   13627123
         ForeColor       =   12582912
         BackColorSel    =   10087415
         ForeColorSel    =   192
         BackColorBkg    =   -2147483637
         GridColor       =   -2147483637
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   690
      Width           =   9975
      Begin VB.TextBox txtItem 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   720
         Width           =   9795
      End
      Begin VB.ComboBox Cbo 
         Height          =   315
         Index           =   2
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   5145
      End
      Begin VB.ComboBox Cbo 
         Height          =   315
         Index           =   1
         Left            =   5190
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item  Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   1500
         TabIndex        =   8
         Top             =   120
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   6330
         TabIndex        =   7
         Top             =   150
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   765
      Index           =   2
      Left            =   -30
      TabIndex        =   5
      Top             =   -30
      Width           =   10035
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Stock Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEF0F7&
         Height          =   315
         Index           =   0
         Left            =   3450
         TabIndex        =   11
         Top             =   240
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   705
      Index           =   1
      Left            =   -60
      TabIndex        =   4
      Top             =   7440
      Width           =   10005
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   270
         Top             =   240
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton Cmdbtn 
         Caption         =   "&Close"
         Height          =   345
         Index           =   4
         Left            =   8790
         TabIndex        =   2
         ToolTipText     =   "Click to Close"
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton Cmdbtn 
         Caption         =   "Print"
         Height          =   345
         Index           =   3
         Left            =   7680
         TabIndex        =   3
         ToolTipText     =   "Click to View Report"
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu subMnuDelete 
         Caption         =   "Refresh"
      End
      Begin VB.Menu jhfgjn 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmCurrentStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objcom As New DSLComFram.clsCommon
Dim objRs As New ADODB.Recordset
Dim newBalance As Double
Private Sub Cbo_Change(Index As Integer)
  load_group
'    load_Brand
'
End Sub

Private Sub Cbo_Click(Index As Integer)
  Select Case Index
          Case 1
               showList (2)
               
          Case 2
                load_group
                showList (1)
'               load_Brand
               
  End Select
End Sub

Private Sub Cmdbtn_Click(Index As Integer)
    Select Case Index
           Case 3
                frmRptStockStatements.Show 1
           Case 4
                Unload Me
                      
           
    End Select
End Sub
Private Sub showList(mode As Integer)
 Dim local_rs As New ADODB.Recordset
 Dim title_rs As New ADODB.Recordset
 Dim total As Double
 
    If mode = 1 Then
       Set local_rs = objcom.Get_RS("SELECT  item_name,stock from min_stock_view where to_number(type_code)=to_number('" & Get_Code(Cbo(2).Text) & "')  order by item_name", objmyCon)
    ElseIf mode = 2 Then
       Set local_rs = objcom.Get_RS("SELECT  item_name,stock from min_stock_view where to_number(type_code)=to_number('" & Get_Code(Cbo(2).Text) & "') and to_number(group_code)=to_number('" & Get_Code(Cbo(1).Text) & "') order by item_name", objmyCon)
     ElseIf mode = 3 Then
       Set local_rs = objcom.Get_RS("SELECT  item_name,stock from min_stock_view where to_number(type_code)='" & Get_Code(Cbo(2).Text) & "' and  to_number(group_code)=to_number('" & Get_Code(Cbo(1).Text) & "') and  upper(item_name) like upper('" & Trim(txtItem) & "%') order by item_name", objmyCon)
    End If
  MSFlexGrid1.Clear
  format_grid
   ''''and ItemBrand='" & objrs(2) & "' and ItemCode='" & objrs(4) & "' and
   If Not local_rs Then
    i = 1
    With MSFlexGrid1
      Do Until local_rs.EOF
               .TextMatrix(i, 0) = local_rs(0)
               .ColAlignment(1) = 0
               .Row = .Row
               .Col = 1
               .CellFontBold = True
               .TextMatrix(i, 1) = "" & Trim(local_rs(1))
              
       local_rs.MoveNext
       i = i + 1
       Loop
      
    End With

    MSFlexGrid1.Rows = 300 + i
Else
    MSFlexGrid1.Rows = 300 + i
 End If
 Set local_rs = Nothing
 Set title_rs = Nothing

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error GoTo ErrorDes
If KeyAscii = 27 Then
   Unload Me
End If

If KeyAscii = 13 Then
   SendKeys Chr(9)
End If
If KeyAscii = 14 Then
   Cmdbtn_Click (1)
End If
If KeyAscii = 19 Then
   Cmdbtn_Click (0)
End If

Exit Sub
ErrorDes:
    MsgBox Err.Description, vbInformation, strmsgtitle

End Sub
Private Sub Form_Load()
  load_type
  load_group
 End Sub
Private Sub format_grid()
  With MSFlexGrid1
    .Rows = 1
    .Cols = 2
    .Col = 0: .Text = "Item"
    .Col = 1: .Text = "Stock"
   
    .ColWidth(0) = 8000
    .ColWidth(1) = 1500
          
    .Rows = 1000
End With
End Sub
Private Sub load_group()
  Set objRs = objcom.Get_RS("SELECT distinct a.group_name ,a.group_code  from item_group_info a  where  a.type_code='" & Get_Code(Cbo(2).Text) & "' order by a.group_code", objmyCon)
      If Not objRs.EOF Then
        objRs.MoveFirst
        Cbo(1).Clear
        Do Until objRs.EOF
          Cbo(1).AddItem objRs(0) & "~" & objRs(1)
          objRs.MoveNext
        Loop
     End If
End Sub
Private Sub load_type()
   Set objRs = objcom.Get_RS("SELECT distinct a.TYPE_name,a.type_code  from item_TYPE_info a  where  a.cate_code='" & CategoryCode & "'", objmyCon)
     If Not objRs.EOF Then
        objRs.MoveFirst
        Cbo(2).Clear
        Do Until objRs.EOF
          Cbo(2).AddItem objRs(0) & "~" & objRs(1)
          objRs.MoveNext
        Loop
     End If
End Sub
Private Sub Frame1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
     PopupMenu mnuDelete, 2
  End If
End Sub
Private Sub Frame2_Click(Index As Integer)
 Select Case Index
         Case 0
             form2.Show 1
         Case 2
             Form1.Show 1
         Case 1
            form3.Show 1
         Case 3
            form11.Show 1
  End Select
         
End Sub
Private Sub Label3_Click(Index As Integer)
   Select Case Index
         Case 0
             form2.Show 1
         Case 2
             Form1.Show 1
         Case 1
            form3.Show 1
         Case 3
            form11.Show 1
  End Select
End Sub

Private Sub mnuClose_Click()
  Unload Me
End Sub

Private Sub subMnuDelete_Click()
  load_group
  
'  load_Brand
  
End Sub

Private Sub txtItem_Change()
   showList (3)
End Sub
