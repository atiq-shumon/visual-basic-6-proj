VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Bed_status 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                                                                  SEARCH       :    ADMITTED PATIENT"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   10590
      TabIndex        =   10
      Top             =   7140
      Width           =   1305
   End
   Begin MSComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   7305
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   7275
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TXTNAME 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   210
      Width           =   2985
   End
   Begin VB.Frame Frame1 
      Height          =   7185
      Left            =   0
      TabIndex        =   3
      Top             =   660
      Width           =   11985
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6285
         Left            =   30
         TabIndex        =   4
         Top             =   180
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   11086
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         RowHeightMin    =   3
         BackColor       =   -2147483624
         BackColorBkg    =   15133394
         FocusRect       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   345
      Left            =   7140
      TabIndex        =   1
      Top             =   180
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   12582912
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mm-YYYY"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   345
      Left            =   9210
      TabIndex        =   2
      Top             =   180
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   12582912
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mm-YYYY"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   1
      Left            =   8790
      TabIndex        =   9
      Top             =   270
      Width           =   270
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ADMISSION DATE :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Index           =   0
      Left            =   5310
      TabIndex        =   6
      Top             =   270
      Width           =   1725
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PATIENT NAME :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   270
      Width           =   1635
   End
   Begin VB.Menu MNUFILE 
      Caption         =   "FILE"
      Begin VB.Menu MNUCLOSE 
         Caption         =   "CLOSE"
      End
   End
End
Attribute VB_Name = "Bed_status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UTILITY As New clsUtility
Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
     Unload Me
  End If
End Sub

Private Sub Form_Load()
  POPULATE_TREE
  MaskEdBox1.Text = Format(Date, "DD/MM/YY")
  MaskEdBox2.Text = Format(Date, "DD/MM/YY")
  FORMAT_GRID (1)
  Call Load_Patient_Info(1, " ")
End Sub
Private Sub Load_Patient_Info(MODE As Integer, NAME As String)
  Dim Conn As New ADODB.Connection
  Dim RS As New ADODB.Recordset
  Dim cmd As New ADODB.Command
  Dim CAB As Integer
  Dim PAY As Integer
  Dim FREE As Integer
  
  Dim i As Integer
  
  If Conn.State = 0 Then
     Conn.ConnectionString = strcn.Connection_String
     Conn.Open
  End If
  cmd.ActiveConnection = Conn
  cmd.CommandType = adCmdText
  If MODE = 1 Then  ''ALL
        cmd.CommandText = "select  A.IN_REG_NO,A.PAT_NAME AS PATIENT_NAME  " & _
        " ,A.BED_TYPE||'(' ||A.CAB_WARD_NO||')-' ||A.BED_NO AS BED,A.ADDR,A.ADMISSION_DATE " & _
           "FROM PAT_SEARCH_OLTP  A  ORDER BY A.ADMISSION_DATE DESC "

        
ElseIf MODE = 2 Then ''PATIENT NAME WISE
    cmd.CommandText = "select  A.IN_REG_NO,A.PAT_NAME AS PATIENT_NAME  " & _
        " ,A.BED_TYPE||'(' ||A.CAB_WARD_NO||')-' ||A.BED_NO AS BED,A.ADDR,A.ADMISSION_DATE " & _
           "FROM PAT_SEARCH_OLTP  A  WHERE  upper(A.PAT_NAME) LIKE upper('%" & NAME & "%') ORDER BY A.ADMISSION_DATE DESC "
    
'ElseIf MODE = 3 Then ''DATE WISE
'    cmd.CommandText = "select  A.IN_REG_NO,A.PAT_NAME AS PATIENT_NAME  " & _
'        " ,A.BED_TYPE||'(' ||A.CAB_WARD_NO||')-' ||A.BED_NO AS BED,A.ADDR,A.ADMISSION_DATE " & _
'           "FROM PAT_SEARCH_OLTP  A  WHERE  TO_DATE(To_CHAR(A.ADMISSION_DATE,'DD-MON-YYYY'),'DD-MON-YYYY') BETWEEN  TO_DATE('" & MaskEdBox1.Text & "','DD-MON-YYYY') AND TO_DATE('" & MaskEdBox2.Text & "','DD-MON-YYYY') ORDER BY A.ADMISSION_DATE DESC "
    
    
    
   
End If

  cmd.Properties("iRowsetChange") = True
  cmd.Properties("updatability") = 7
  RS.CursorLocation = adUseClient

  RS.Open cmd.CommandText, Conn, adOpenForwardOnly, adLockReadOnly
  cmd.Properties("iRowsetChange") = False

  If Not RS.EOF Then
       i = 1
    With MSFlexGrid1
          Do Until RS.EOF
                .Rows = i + 1
               .ColAlignment(0) = 0
               .TextMatrix(i, 0) = RS!in_reg_no
               .Col = 1
               .Row = i
               .CellForeColor = vbBlue
               .TextMatrix(i, 1) = RS!PATIENT_NAME
               .TextMatrix(i, 2) = RS!BED
               .ColAlignment(3) = 0
               .Col = 3
               .Row = i
               .CellForeColor = &H8080FF
               .TextMatrix(i, 3) = RS!addr
               .ColAlignment(4) = 0
               .TextMatrix(i, 4) = Format(RS!admission_date, "DD/MM/YY")
               If UCase(Mid(RS!BED, 1, 3)) = "CAB" Then
                  CAB = CAB + 1
               ElseIf UCase(Mid(RS!BED, 1, 3)) = "PAY" Then
                  PAY = PAY + 1
               Else
                  FREE = FREE + 1
               End If
              
               i = i + 1
            RS.MoveNext
        Loop
        With StatusBar2
            
            .Panels(1).Width = 3000
            .Panels(1).Text = "TOTAL SHOWN PATIENT :" & RS.RecordCount
            .Panels(2).Text = " ( CABIN : " & CAB & ")"
            .Panels(3).Width = 2000
            
            .Panels(3).Text = " ( PAYING : " & PAY & ")"
            .Panels(4).Width = 2000
            .Panels(4).Text = " ( FREE-BED : " & FREE & ")"
            
            
             End With

    End With
Else
    MSFlexGrid1.Rows = 2
 End If
 
End Sub

Private Sub POPULATE_TREE()
  Dim ND As Node
    With TreeView1
'     Set ND = .Nodes.Add("", Root, Root, 1)
    End With
End Sub

Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then
'      If UTILITY.START_END_VALIDATION(MaskEdBox1, MaskEdBox2) = False Then
'          MsgBox "Start Date can't be greater(>) than End date..Verify"
'          MaskEdBox1.SetFocus
'          Exit Sub
'     End If
'    Call Load_Patient_Info(3, "")
'  End If
End Sub

Private Sub MNUCLOSE_Click()
   Unload Me
End Sub
Private Sub FORMAT_GRID(MODE As Integer)
  If MODE = 1 Then
     With MSFlexGrid1
         .Rows = 1
         .Cols = 5
         .Col = 0: .Text = "REG NO"
         .Col = 1: .Text = " NAME"
         .Col = 2: .Text = " BED  "
         .Col = 3: .Text = " ADDRESS "
         .Col = 4: .Text = " ADM. DATE "
         
         .ColWidth(0) = 1000
         .ColWidth(1) = 2500
         .ColWidth(2) = 2000
         .ColWidth(3) = 4550
         .ColWidth(4) = 1500
         
     End With
  End If
  
End Sub

Private Sub MSFlexGrid1_Click()
      TxtName.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
End Sub

Private Sub MSFlexGrid1_SelChange()
  MSFlexGrid1_Click
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Len(TxtName) <> 0 Then
         Call Load_Patient_Info(2, Trim(TxtName))
     Else
         Call Load_Patient_Info(1, "")
     End If
  End If
End Sub
