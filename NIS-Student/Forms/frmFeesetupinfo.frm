VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFeesetupInfo 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   18
      Top             =   6330
      Width           =   9795
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
         Left            =   6660
         TabIndex        =   10
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
         Left            =   4710
         TabIndex        =   9
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
         Left            =   5685
         TabIndex        =   8
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
         Left            =   7635
         TabIndex        =   11
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
         Left            =   8610
         TabIndex        =   12
         ToolTipText     =   "Click to Exit"
         Top             =   210
         Width           =   945
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   4680
         Top             =   180
         Width           =   4905
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   645
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   9765
      TabIndex        =   15
      Top             =   0
      Width           =   9825
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fee Setup  Information "
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
         Left            =   3600
         TabIndex        =   25
         Top             =   120
         Width           =   2640
      End
      Begin VB.Image Image1 
         Height          =   690
         Left            =   -30
         Picture         =   "frmFeesetupinfo.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   9765
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   1935
      Left            =   -30
      TabIndex        =   13
      Top             =   660
      Width           =   9855
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   5
         Left            =   4740
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1530
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000B&
         Caption         =   "Have Alternative"
         Height          =   225
         Left            =   960
         TabIndex        =   6
         Top             =   1530
         Width           =   1665
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000B&
         Caption         =   "Optional"
         Height          =   255
         Left            =   8430
         TabIndex        =   5
         Top             =   1140
         Width           =   885
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000B&
         Caption         =   "Compulsory"
         Height          =   255
         Left            =   6600
         TabIndex        =   4
         Top             =   1140
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   4
         Left            =   4740
         MaxLength       =   2
         TabIndex        =   3
         ToolTipText     =   "Insert No of Times"
         Top             =   1080
         Width           =   1035
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   8700
         TabIndex        =   24
         Text            =   "Combo3"
         Top             =   1650
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   8130
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   23
         ToolTipText     =   "Insert Marks Category"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   2
         Left            =   960
         MaxLength       =   8
         TabIndex        =   2
         ToolTipText     =   "Insert Marks Category"
         Top             =   1080
         Width           =   1665
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   630
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   210
         Width           =   1695
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   4740
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   19
         ToolTipText     =   "Insert Marks Category"
         Top             =   210
         Width           =   4785
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   4740
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   14
         ToolTipText     =   "Insert Marks Category"
         Top             =   630
         Width           =   4785
      End
      Begin VB.Shape Shape2 
         Height          =   315
         Left            =   6450
         Top             =   1110
         Width           =   3075
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alternative Fee Code"
         Height          =   195
         Index           =   4
         Left            =   3000
         TabIndex        =   28
         Top             =   1530
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No of Times(Yearly)"
         Height          =   195
         Index           =   3
         Left            =   3030
         TabIndex        =   27
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   22
         Top             =   1110
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fee Title"
         Height          =   195
         Index           =   1
         Left            =   3060
         TabIndex        =   21
         Top             =   645
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Name"
         Height          =   195
         Index           =   1
         Left            =   3060
         TabIndex        =   20
         Top             =   195
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fee Code"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   17
         Top             =   690
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Id"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   16
         Top             =   270
         Width           =   555
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3765
      Left            =   0
      TabIndex        =   26
      Top             =   2580
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   6641
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
Attribute VB_Name = "frmFeesetupInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
  If Check1.Value = 1 Then
     Label3(4).Visible = True
     txtfields(5).Visible = True
  Else
     Label3(4).Visible = False
     txtfields(5).Visible = False
  End If
End Sub

Private Sub cmdDelete_Click()
If Len(Combo1.Text) = 0 Then
                MsgBox "Please Select a Class Id", vbInformation, App.Title
                Combo1.SetFocus
                Exit Sub
            End If
            
            
            If Len(Combo2.Text) = 0 Then
                MsgBox "Please Select a Fee Code", vbInformation, App.Title
                Combo2.SetFocus
                Exit Sub
            End If
            
            
            Dim rs2 As New ADODB.Recordset
             Set rs2 = getdata("SELECT Srl_No from fee_setup WHERE Class_id= '" & Trim(Combo1.Text) & "' and Fee_Code='" & Trim(Combo2.Text) & "'")
             
            If rs2.EOF Then
               MsgBox "No Such Class & Fee Code Exists..Please Verify.", vbInformation, cmp
               Exit Sub
           End If
            
            
            
            
            Dim rs1 As New ADODB.Recordset
             Set rs1 = getdata("SELECT Srl_No from fee_setup WHERE (Srl_No= '" & txtfields(3) & "')")

            If rs1.EOF Then
               MsgBox "No Such Class & Fee Code Exists..Please Verify.", vbInformation, cmp
               Exit Sub
            End If
            
            If MsgBox("Are You Sure to Delete ?", vbQuestion + vbYesNo, cmp) = vbYes Then

                Dim rs As New ADODB.Recordset
                Dim cmd As New ADODB.Command
                Dim con As New ADODB.connection
                con.Open GConnString
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "Fee_setup_Save"
                cmd(1) = "d"
                cmd(2) = IIf(IsNull(Val(Trim(txtfields(3)))), 0, Val(Trim(txtfields(3))))
                cmd(3) = Trim(Combo1.Text)
                cmd(4) = Trim(Combo2.Text)
                cmd(5) = Trim(Combo3.Text)
                cmd(6) = Trim(txtfields(2).Text)
                cmd(7) = soft_user
                cmd(8) = Format(Date, "DD MMM YYYY")
                cmd(9) = txtfields(4)
                If Option1.Value = True Then
                 cmd(10) = 1
                Else
                 cmd(10) = 0
                End If
                If Check1.Value = 1 Then
                  cmd(11) = 1
                  cmd(12) = txtfields(5)
                Else
                  cmd(11) = 0
                  cmd(12) = Null
                End If
                cmd.Execute
                
                MsgBox "Deleted successfully.", vbInformation, "Student Management System"
                Call ShowFlexData
                cmdnew.SetFocus
             End If

End Sub

Private Sub cmdEdit_Click()
  If Len(Combo1.Text) = 0 Then
                MsgBox "Please Select a Class Id", vbInformation, App.Title
                Combo1.SetFocus
                Exit Sub
            End If
            
            
            If Len(Combo2.Text) = 0 Then
                MsgBox "Please Select a Fee Code", vbInformation, App.Title
                Combo2.SetFocus
                Exit Sub
            End If
            
            If Len(txtfields(2)) = 0 Then
                MsgBox "Please enter an amount", vbInformation, App.Title
                txtfields(2).SetFocus
                Exit Sub
            End If
            
            
            
            Dim rs2 As New ADODB.Recordset
             Set rs2 = getdata("SELECT Srl_No from fee_setup WHERE Class_id= '" & Trim(Combo1.Text) & "' and Fee_Code='" & Trim(Combo2.Text) & "'")
             
            If rs2.EOF Then
               MsgBox "No Such Class & Fee Code Exists..Please Verify.", vbInformation, cmp
               Exit Sub
            End If
            
            
            
            
            Dim rs1 As New ADODB.Recordset
             Set rs1 = getdata("SELECT Srl_No from fee_setup WHERE (Srl_No= '" & txtfields(3) & "')")

            If rs1.EOF Then
               MsgBox "No Such Class & Fee Code Exists..Please Verify.", vbInformation, cmp
               Exit Sub
            End If

                Dim rs As New ADODB.Recordset
                Dim cmd As New ADODB.Command
                Dim con As New ADODB.connection
                con.Open GConnString
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "Fee_setup_Save"
                cmd(1) = "u"
                cmd(2) = IIf(IsNull(Val(Trim(txtfields(3)))), 0, Val(Trim(txtfields(3))))
                cmd(3) = Trim(Combo1.Text)
                cmd(4) = Trim(Combo2.Text)
                cmd(5) = Trim(Combo3.Text)
                cmd(6) = Trim(txtfields(2).Text)
                cmd(7) = soft_user
                cmd(8) = Format(Date, "DD MMM YYYY")
                cmd(9) = txtfields(4)
                If Option1.Value = True Then
                 cmd(10) = 1
                Else
                 cmd(10) = 0
                End If
                If Check1.Value = 1 Then
                  cmd(11) = 1
                  cmd(12) = txtfields(5)
                Else
                  cmd(11) = 0
                  cmd(12) = Null
                End If
                cmd.Execute
                MsgBox "Edited successfully.", vbInformation, "Student Management System"
                Call ShowFlexData
                cmdnew.SetFocus

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
     txtfields(2).Text = ""
     txtfields(3).Text = ""
     txtfields(0).Text = ""
     txtfields(1).Text = ""
     txtfields(4).Text = ""
     Combo1.Clear
     Combo2.Clear
     Combo1.SetFocus
     load_class
     load_Fee
End Sub


Private Sub cmdSAVE_Click()

            If Len(Combo1.Text) = 0 Then
                MsgBox "Please Select a Class Id", vbInformation, App.Title
                Combo1.SetFocus
                Exit Sub
            End If
            
            
            If Len(Combo2.Text) = 0 Then
                MsgBox "Please Select a Fee Code", vbInformation, App.Title
                Combo2.SetFocus
                Exit Sub
            End If
            
            If Len(txtfields(2)) = 0 Then
                MsgBox "Please enter an amount", vbInformation, App.Title
                txtfields(2).SetFocus
                Exit Sub
            End If
            
            
            Dim rs2 As New ADODB.Recordset
             Set rs2 = getdata("SELECT Srl_No from fee_setup WHERE Class_id= '" & Trim(Combo1.Text) & "' and Fee_Code='" & Trim(Combo2.Text) & "'")
             
            If Not rs2.EOF Then
               MsgBox "Same Class & Fee Code already Exists..Please Verify.", vbInformation, cmp
               Combo1.SetFocus
               Exit Sub
            End If
            
            
            
            
            Dim rs1 As New ADODB.Recordset
             Set rs1 = getdata("SELECT Srl_No from fee_setup WHERE (Srl_No= '" & txtfields(3) & "')")

            If Not rs1.EOF Then
               MsgBox "Same Class & Fee Code already Exists..Please Verify.", vbInformation, cmp
               Exit Sub
            End If

                Dim rs As New ADODB.Recordset
                Dim cmd As New ADODB.Command
                Dim con As New ADODB.connection
                con.Open GConnString
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "Fee_setup_Save"
                cmd(1) = "s"
                cmd(2) = IIf(IsNull(Val(Trim(txtfields(3)))), 0, Val(Trim(txtfields(3))))
                cmd(3) = Trim(Combo1.Text)
                cmd(4) = Trim(Combo2.Text)
                cmd(5) = Trim(Combo3.Text)
                cmd(6) = Trim(txtfields(2).Text)
                cmd(7) = soft_user
                cmd(8) = Format(Date, "DD MMM YYYY")
                cmd(9) = txtfields(4)
                If Option1.Value = True Then
                 cmd(10) = 1
                Else
                 cmd(10) = 0
                End If
                If Check1.Value = 1 Then
                  cmd(11) = 1
                  cmd(12) = txtfields(5)
                Else
                  cmd(11) = 0
                  cmd(12) = Null
                End If
                
                cmd.Execute
                MsgBox "Saved successfully.", vbInformation, "Student Management System"
                Call ShowFlexData
                Combo2.SetFocus
                txtfields(2).Text = ""
                txtfields(3).Text = ""
                txtfields(1).Text = ""
                txtfields(4).Text = ""
 

End Sub

Private Sub dtpic_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdsave.SetFocus
End If
End Sub

Private Sub Combo1_Click()
     load_title
     varify_srl
     ShowFlexData
End Sub

Private Sub Combo2_Click()
  load_Fee_title
  varify_srl
End Sub
Private Sub varify_srl()
  Dim rs2 As New ADODB.Recordset
             Set rs2 = getdata("SELECT Srl_No from fee_setup WHERE Class_id= '" & Trim(Combo1.Text) & "' and Fee_Code='" & Trim(Combo2.Text) & "'")
             
            If Not rs2.EOF Then
               txtfields(3).Text = rs2!srl_no
            Else
               txtfields(3).Text = ""
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
Private Sub load_class()
Dim rs As New ADODB.Recordset
Combo1.Clear
Set rs = getdata("Select ClassId from ClassInfo")
If Not rs.EOF Then
    Do Until rs.EOF
        Combo1.AddItem rs(0)
        rs.MoveNext
    Loop
    
End If
End Sub
Private Sub load_Fee()
Dim rs As New ADODB.Recordset
Combo2.Clear
Set rs = getdata("Select fee_code from Fee_Info")
If Not rs.EOF Then
    Do Until rs.EOF
        Combo2.AddItem rs(0)
        rs.MoveNext
    Loop
    
End If
End Sub
Private Sub load_title()
Dim rs As New ADODB.Recordset
Set rs = getdata("Select Classname  from ClassInfo where classid='" & Trim(Combo1.Text) & "'")
If Not rs.EOF Then
  txtfields(0).Text = rs(0)
End If
End Sub
Private Sub load_Fee_title()
Dim rs As New ADODB.Recordset
Set rs = getdata("Select Fee_title  from Fee_Info where Fee_code='" & Trim(Combo2.Text) & "'")
If Not rs.EOF Then
  txtfields(1).Text = rs(0)
End If
End Sub

Private Sub Form_Load()
'Dim rs As New adodb.Recordset
'Set rs = GetData("select max (Fee_code+1)from fee_info")
'If Not rs.EOF Then
'    txtfields(0) = IIf(IsNull(rs(0)) = True, "01", Format(rs(0), "00"))
'Else
'    txtfields(0) = "01"
'End If
With MSFlexGrid1
    .Rows = 1
    .Cols = 11
    .Col = 0: .Text = " Class ID"
    .Col = 1: .Text = " Class Name"
    .Col = 2: .Text = " Fee Code"
    .Col = 3: .Text = " Fee Title"
    .Col = 4: .Text = " Amount"
    .Col = 5: .Text = " Serial"
    .Col = 6: .Text = "No of Times"
    .Col = 7: .Text = "Total"

    .ColWidth(0) = 0
    .ColWidth(1) = 2000
    .ColAlignment(2) = 0
    .ColWidth(2) = 1000
    .ColWidth(3) = 3800
    .ColAlignment(4) = 0
    .ColWidth(4) = 1000
    .ColWidth(5) = 0
    .ColAlignment(6) = 0
    .ColWidth(6) = 900
    .ColAlignment(7) = 0
    .ColWidth(7) = 1100
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(10) = 0

End With
Call ShowFlexData
load_class
load_Fee
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub MSFlexGrid1_SelChange()
  MSFlexGrid1_Click
End Sub

Private Sub txtfields_Change(Index As Integer)
            Select Case Index
                   Case 2, 4
                        If Not IsNumeric(txtfields(Index)) Then
                           txtfields(Index) = ""
                        End If
            End Select
End Sub

Private Sub txtfields_GotFocus(Index As Integer)
             Select Case Index
                    Case Index
                         txtfields(Index).SelStart = 0
                         txtfields(Index).SelLength = Len(Trim(txtfields(Index)))
             End Select
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
'Dim rs As New adodb.Recordset
'
'Select Case Index
'    Case 0
'        If Len(Trim(txtfields(0))) = 0 Then Exit Sub
'
'            txtfields(0) = Format(txtfields(0), "00000")
'
'            Set rs = GetData("SELECT mcategoryDsc,Note,EntryBy,Entrydate from Markscategory WHERE (McategoryID= '" & txtfields(0) & "')")
'                 If Not rs.EOF Then
'                        txtfields(1) = rs!mcategoryDsc
'                        txtfields(2) = rs!Note
'                        txtfields(3) = rs!EntryBy
''                        dtpic = rs!Format(Entrydate, "dd/mmm/yyyy")
'                End If
'
'End Select
End Sub
Private Sub ShowFlexData()
'On Error GoTo errdes
Dim rs As New ADODB.Recordset
Dim amt As Double
Set rs = getdata("SELECT a.Class_id as Class,(Select Classname  from ClassInfo where classid=a.class_id) as class_name,a.Fee_code as Code,(Select Fee_title  from Fee_Info where Fee_code=a.fee_code) as fee_title,a.Fee_amt as Amount,a.srl_no as serial,a.NoOfTimes,a.FeesStatus,a.AlternativeFlag,a.AlternativeCode From fee_setup a where a.Class_id='" & Trim(Combo1) & "'")
If Not rs.EOF Then
    i = 1
    amt = 0
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = rs!Class
                .TextMatrix(i, 1) = rs!Class_name
                .TextMatrix(i, 2) = rs!Code
                .TextMatrix(i, 3) = rs!fee_title
                .TextMatrix(i, 4) = rs!amount
                .TextMatrix(i, 5) = rs!serial
                .TextMatrix(i, 6) = "" & rs!NoOfTimes
                If Not IsNull(rs!NoOfTimes) Then
                   .TextMatrix(i, 7) = rs!NoOfTimes * rs!amount
                   amt = amt + .TextMatrix(i, 7)
                End If
                .TextMatrix(i, 8) = "" & rs!FeesStatus
                .TextMatrix(i, 9) = "" & rs!AlternativeFlag
               .TextMatrix(i, 10) = "" & rs!AlternativeCode
                i = i + 1
            rs.MoveNext
        Loop
        .Rows = i + 1
        .Row = i
        .Col = 6
        .CellFontBold = True
        .CellForeColor = vbRed
        .TextMatrix(i, 6) = "Net Total"
        .Row = i
        .Col = 7
        .CellFontBold = True
        .CellForeColor = vbRed
        
        .TextMatrix(i, 7) = amt
       
        
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
'Combo1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
Combo2 = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
txtfields(2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
txtfields(3) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
txtfields(4) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8) = 1 Then
   Option1.Value = True
   Option2.Value = False
Else
   Option1.Value = False
   Option2.Value = True
End If
If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9) = 1 Then
   Check1.Value = 1
   Label3(4).Visible = True
   txtfields(5).Visible = True
   txtfields(5).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)
Else
   Check1.Value = 0
   Label3(4).Visible = False
   txtfields(5).Visible = False
   txtfields(5).Text = ""
End If
Exit Sub
errdes:
'MsgBox err.Description, vbInformation, App.Title


End Sub

