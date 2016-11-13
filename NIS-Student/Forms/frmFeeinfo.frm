VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFeeInfo 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   0
      TabIndex        =   10
      Top             =   5190
      Width           =   8595
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
         Left            =   5580
         TabIndex        =   11
         ToolTipText     =   "Click to insert new information"
         Top             =   210
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
         Left            =   3630
         TabIndex        =   3
         ToolTipText     =   "Click to insert new information"
         Top             =   210
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
         Left            =   4605
         TabIndex        =   2
         ToolTipText     =   "Click to Save"
         Top             =   210
         Width           =   945
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H8000000C&
         Caption         =   "Delete"
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
         Left            =   6555
         TabIndex        =   4
         ToolTipText     =   "Click to Delete"
         Top             =   210
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
         Left            =   7530
         TabIndex        =   5
         ToolTipText     =   "Click to Exit"
         Top             =   210
         Width           =   945
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   3600
         Top             =   180
         Width           =   4935
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   30
      ScaleHeight     =   795
      ScaleWidth      =   8535
      TabIndex        =   7
      Top             =   30
      Width           =   8595
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fee Category Information "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEF0F7&
         Height          =   270
         Left            =   2880
         TabIndex        =   12
         Top             =   210
         Width           =   2940
      End
      Begin VB.Image Image1 
         Height          =   1050
         Left            =   -480
         Picture         =   "frmFeeinfo.frx":0000
         Stretch         =   -1  'True
         Top             =   -90
         Width           =   9015
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   1185
      Left            =   0
      TabIndex        =   0
      Top             =   870
      Width           =   8625
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   1
         Left            =   1380
         MaxLength       =   80
         TabIndex        =   1
         ToolTipText     =   "Insert Marks Category"
         Top             =   630
         Width           =   5475
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1380
         TabIndex        =   6
         Top             =   210
         Width           =   1515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fee Title"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   630
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fee Code"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   240
         Width           =   690
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3135
      Left            =   -30
      TabIndex        =   13
      Top             =   2070
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   5530
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   15005934
      BackColorSel    =   -2147483624
      ForeColorSel    =   16711680
      BackColorBkg    =   15724265
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmFeeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDelete_Click()

 Dim rs1 As New ADODB.Recordset
     If Len(txtfields(0)) = 0 Then
                MsgBox "Fee Code Mandatory", vbInformation, App.Title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(txtfields(1)) = 0 Then
                MsgBox "Fee Title Mandatory", vbInformation, App.Title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            Set rs1 = getdata("SELECT fee_code from fee_info WHERE (fee_code= '" & txtfields(0) & "')")
            
            If rs1.EOF Then
               MsgBox "No such Fee Code Exists..Please Verify.", vbInformation, cmp
               cmdnew.SetFocus
               Exit Sub
            End If
            
                     
        If Len(Trim(txtfields(1))) <> 0 Then
           If (MsgBox("Are You sure to delete ?", vbYesNo + vbCritical) = vbYes) Then
                Dim rs As New ADODB.Recordset
                Dim cmd As New ADODB.Command
                Dim con As New ADODB.connection
                con.Open GConnString
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "Fee_info_Save"
                cmd(1) = "d"
                cmd(2) = Format(Trim(txtfields(0)), "00")
                cmd(3) = Trim(txtfields(1))
                cmd(4) = soft_user
                cmd(5) = Format(Date, "DD MMM YYYY")
                cmd.Execute
                MsgBox "Deleted successfully.", vbInformation, cmp
                Call ShowFlexData
                cmdnew.SetFocus
         End If
                

End If

End Sub

Private Sub CmdEdit_Click()
   Dim rs1 As New ADODB.Recordset
     If Len(txtfields(0)) = 0 Then
                MsgBox "Fee Code Mandatory", vbInformation, App.Title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(txtfields(1)) = 0 Then
                MsgBox "Fee Title Mandatory", vbInformation, App.Title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            Set rs1 = getdata("SELECT fee_code from fee_info WHERE (fee_code= '" & txtfields(0) & "')")
            
            If rs1.EOF Then
               MsgBox "No such Fee Code Exists..Please Verify.", vbInformation, cmp
               cmdnew.SetFocus
               Exit Sub
            End If
            
            
                Dim rs As New ADODB.Recordset
                Dim cmd As New ADODB.Command
                Dim con As New ADODB.connection
                con.Open GConnString
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "Fee_info_Save"
                cmd(1) = "u"
                cmd(2) = Format(Trim(txtfields(0)), "00")
                cmd(3) = Trim(txtfields(1))
                cmd(4) = soft_user
                cmd(5) = Format(Date, "DD MMM YYYY")
                cmd.Execute
                MsgBox "Edited successfully.", vbInformation, cmp
                Call ShowFlexData
                cmdnew.SetFocus

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
Set rs = getdata("select max(Fee_code+1)from fee_info")
If Not rs.EOF Then
    txtfields(0) = IIf(IsNull(rs(0)) = True, "01", Format(rs(0), "00"))
Else
    txtfields(0) = "01"
End If


    txtfields(1) = ""

txtfields(1).SetFocus
End Sub


Private Sub cmdSAVE_Click()

            If Len(txtfields(0)) = 0 Then
                MsgBox "Fee Code Mandatory", vbInformation, App.Title
                cmdnew.SetFocus
                Exit Sub
            End If
            
            
            If Len(txtfields(1)) = 0 Then
                MsgBox "Fee Title Mandatory", vbInformation, App.Title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            
            Dim rs1 As New ADODB.Recordset
             Set rs1 = getdata("SELECT fee_code from fee_info WHERE (fee_code= '" & txtfields(0) & "')")
            
            If Not rs1.EOF Then
               MsgBox "Same Fee Code already Exists..Please Verify.", vbInformation, cmp
               Exit Sub
            End If
            
                Dim rs As New ADODB.Recordset
                Dim cmd As New ADODB.Command
                Dim con As New ADODB.connection
                con.Open GConnString
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "Fee_info_Save"
                cmd(1) = "s"
                cmd(2) = Format(Trim(txtfields(0)), "00")
                cmd(3) = Trim(txtfields(1))
                cmd(4) = soft_user
                cmd(5) = Format(Date, "DD MMM YYYY")
                cmd.Execute
                MsgBox "Saved successfully.", vbInformation, "Student Management System"
                Call ShowFlexData
                cmdnew.SetFocus

End Sub

Private Sub dtpic_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdsave.SetFocus
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys (Chr(9))
   Else
       If KeyAscii = 27 Then
          Unload Me
       End If
       
   
  End If
   
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Set rs = getdata("select max (Fee_code+1)from fee_info")
If Not rs.EOF Then
    txtfields(0) = IIf(IsNull(rs(0)) = True, "01", Format(rs(0), "00"))
Else
    txtfields(0) = "01"
End If
With MSFlexGrid1
    .Rows = 1
    .Cols = 2
    .Col = 0: .Text = " Code"
    .Col = 1: .Text = " Title"
    
    .ColWidth(0) = 800
    .ColWidth(1) = 7700
    
End With
Call ShowFlexData
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub MSFlexGrid1_SelChange()
   MSFlexGrid1_Click
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
    Case Index
        If Index <> 3 Then
'            txtfields(Index + 1).SetFocus
        Else
            dtpic.SetFocus
        End If
    End Select
End If
End Sub
Public Function getdata(SQLString As String) As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = SQLString

 Set rs = cmd.Execute
Set getdata = rs
End Function

Private Sub txtfields_LostFocus(Index As Integer)
Dim rs As New ADODB.Recordset

Select Case Index
    Case 0
        If Len(Trim(txtfields(0))) = 0 Then Exit Sub
      
            txtfields(0) = Format(txtfields(0), "00000")
          
            Set rs = getdata("SELECT mcategoryDsc,Note,EntryBy,Entrydate from Markscategory WHERE (McategoryID= '" & txtfields(0) & "')")
                 If Not rs.EOF Then
                        txtfields(1) = rs!mcategoryDsc
                        txtfields(2) = rs!Note
                        txtfields(3) = rs!EntryBy
'                        dtpic = rs!Format(Entrydate, "dd/mmm/yyyy")
                End If
        
End Select
End Sub
Private Sub ShowFlexData()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT Fee_code as Code ,Fee_title as Title From fee_info")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = rs!Code
                .TextMatrix(i, 1) = rs!Title
                i = i + 1
            rs.MoveNext
        Loop
    End With
Else
    MSFlexGrid1.Rows = 1
 End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub MSFlexGrid1_Click()

On Error GoTo errdes
txtfields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
txtfields(2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
Exit Sub
errdes:
'MsgBox err.Description, vbInformation, App.Title


End Sub

