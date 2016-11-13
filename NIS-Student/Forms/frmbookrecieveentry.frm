VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmbookrecieveentry 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   435
      Left            =   6900
      TabIndex        =   9
      Top             =   4710
      Width           =   945
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H8000000C&
      Caption         =   "Delete"
      Height          =   435
      Left            =   5910
      TabIndex        =   8
      Top             =   4710
      Width           =   945
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   435
      Left            =   4920
      TabIndex        =   7
      Top             =   4710
      Width           =   945
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H8000000C&
      Caption         =   "New"
      Height          =   435
      Left            =   3930
      TabIndex        =   0
      Top             =   4710
      Width           =   945
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1605
      Left            =   0
      TabIndex        =   26
      Top             =   3060
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   2831
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   0
      TabIndex        =   11
      Top             =   810
      Width           =   7845
      Begin VB.CommandButton Command2 
         Height          =   300
         Left            =   2460
         Picture         =   "frmbookrecieveentry.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1200
         Width           =   420
      End
      Begin VB.CommandButton Command1 
         Height          =   300
         Left            =   2460
         Picture         =   "frmbookrecieveentry.frx":02E2
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   870
         Width           =   420
      End
      Begin VB.CommandButton cmdsearch 
         Height          =   300
         Left            =   2460
         Picture         =   "frmbookrecieveentry.frx":05C4
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   540
         Width           =   420
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   6
         Top             =   1860
         Width           =   6615
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   5
         Top             =   1530
         Width           =   1305
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   4
         Top             =   1200
         Width           =   1305
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   3
         Top             =   870
         Width           =   1305
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   285
         Left            =   4110
         TabIndex        =   1
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   540
         Width           =   1305
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   12
         Top             =   210
         Width           =   1305
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   2940
         TabIndex        =   22
         Top             =   1200
         Width           =   4755
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   2940
         TabIndex        =   21
         Top             =   870
         Width           =   4755
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   2940
         TabIndex        =   20
         Top             =   540
         Width           =   4755
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   1860
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   255
         Left            =   90
         TabIndex        =   18
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject ID"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   1260
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class ID"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   900
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recieve Date"
         Height          =   255
         Left            =   2940
         TabIndex        =   15
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier ID"
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   570
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recieve No"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000006&
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   7785
      TabIndex        =   10
      Top             =   0
      Width           =   7845
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Recieved Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEF0F7&
         Height          =   285
         Left            =   1980
         TabIndex        =   27
         Top             =   150
         Width           =   3105
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   -90
         Picture         =   "frmbookrecieveentry.frx":08A6
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   7935
      End
   End
End
Attribute VB_Name = "frmbookrecieveentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
If MsgBox("Are You sure to Delete?", vbYesNo + vbCritical) = vbYes Then
    cmd.CommandType = adCmdText
    cmd.CommandText = "Delete from BookRecievedInfo  where suppId='" & Trim(txtfields(1)) & "'and ClassId='" & Trim(txtfields(2)) & "'and SubjectId='" & Trim(txtfields(3)) & "'and recieveno='" & Trim(txtfields(0)) & "' "
    cmd.Execute
    MsgBox "Delete Successfully Book Recieval Information.", vbInformation, App.Title
    txtfields(3) = ""
    txtfields(4) = ""
    txtfields(5) = ""
    Label12(1).Caption = ""
    Label12(2).Caption = ""
    Label12(3).Caption = ""
    MaskEdBox1 = Format(MaskEdBox1, "__/__/__")
    Call ShowFlexData
Else
    Exit Sub
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
Select Case Index
Case 0
   
    For i = 0 To 5
       txtfields(i) = ""
    Next
    Label12(1).Caption = ""
    Label12(2).Caption = ""
    Label12(3).Caption = ""
    MaskEdBox1 = Format(MaskEdBox1, "__/__/__")
    MaskEdBox1.SetFocus
    End Select
End Sub

Private Sub cmdSAVE_Click()
On Error GoTo errdes
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
    
    If Len(txtfields(1)) = 0 And Len(txtfields(2)) = 0 And Len(txtfields(3)) = 0 And Len(txtfields(4)) = 0 Then Exit Sub
    
    If Len(txtfields(1)) = 0 Then
        MsgBox "Please Enter Supplier ID.", vbInformation, App.Title
        txtfields(1).SetFocus
        Exit Sub
    End If
    
    If Len(txtfields(2)) = 0 Then
       MsgBox "Please Enter Class Id.", vbInformation, App.Title
       txtfields(2).SetFocus
       Exit Sub
    End If
    
     If Len(txtfields(3)) = 0 Then
       MsgBox "Please Enter Subject Id.", vbInformation, App.Title
       txtfields(3).SetFocus
       Exit Sub
    End If
     If Len(txtfields(4)) = 0 Then
       MsgBox "Please Enter Quantity.", vbInformation, App.Title
       txtfields(4).SetFocus
       Exit Sub
    End If
    If MaskEdBox1 = "__/__/__" Then
    MsgBox "Please Enter Date.", vbInformation, App.Title
       MaskEdBox1.SetFocus
       Exit Sub
    End If
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BMEntryInformation"
    cmd(1) = txtfields(0)
    cmd(2) = Trim(txtfields(1))
    cmd(3) = IIf(MaskEdBox1 = "__/__/__", Format(Date, "DD MMM YYYY"), Format(MaskEdBox1, "dd mmm yyyy"))
    
    cmd(4) = Trim(txtfields(2))
    
    cmd(5) = Trim(txtfields(3))
    cmd(6) = Trim(txtfields(4))
    cmd(7) = Trim(txtfields(5))
    cmd(8) = Date
    cmd(9) = "DSL"

    cmd.Execute
    MsgBox "Save Successfully.", vbInformation, "Student Management System"
    cmdnew.SetFocus
    get_Maximum
    Call ShowFlexData
errdes:

End Sub

Private Sub cmdsearch_Click()
Dim f As New frmFind
Set f.OwnerForm = Me
    f.intInputsel = 1
    f.SQLString = "Select SuppId, Suppname from SupplierInfo"
    f.Show 1
    txtfields(1).SetFocus
End Sub

Private Sub Command1_Click()
Dim f As New frmFind
Set f.OwnerForm = Me
    f.intInputsel = 2
    f.SQLString = "Select ClassId, Classname from ClassInfo"
    f.Show 1
    txtfields(2).SetFocus
End Sub

Private Sub Command2_Click()
Dim f As New frmFind
Set f.OwnerForm = Me
    f.intInputsel = 3
    f.SQLString = "Select Sub_code, Sub_title from Subject_info_sub where Class_code='" & Trim(txtfields(2)) & "'"
    f.Show 1
    txtfields(3).SetFocus
End Sub


Private Sub Form_Load()

MSFlexGrid1.ColWidth(0) = 0

With MSFlexGrid1
    .Rows = 1
    .Cols = 7
    .Col = 1: .Text = "  Recieve No       "
    .Col = 2: .Text = "SubjectID#  "
    .Col = 3: .Text = " Subject Name"
    .Col = 4: .Text = " Quantity"
     .Col = 5: .Text = " Remarks"
    .Col = 6: .Text = " Recieve Date"
    
    
    .ColWidth(1) = 2000
    .ColWidth(2) = 2000
    .ColWidth(3) = 5500
    .ColWidth(4) = 2000
    .ColWidth(5) = 4000
End With
Call ShowFlexData

Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If MaskEdBox1 <> "__/__/__" Then
            If Check_ValidDate(MaskEdBox1) = False Then
                MaskEdBox1.SetFocus
                Exit Sub
            End If
    Else
        txtfields(1).SetFocus
    End If
End If
End Sub

Private Sub MSFlexGrid1_Click()

On Error GoTo errdes
If MSFlexGrid1.Rows = 1 Then Exit Sub
txtfields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
txtfields(3) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
Label12(3) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
txtfields(4) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
txtfields(5) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
MaskEdBox1 = Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6), "DD/mm/yy")
Exit Sub
errdes:
'MsgBox Err.Description, vbInformation, App.Title

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

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
Dim rs As New ADODB.Recordset
If KeyAscii = 13 Then
    Select Case Index
        Case 1
             If Len(Trim(txtfields(1))) = 0 Then
                cmdsearch_Click
                
             Else
                Set rs = getdata("select SuppId from SupplierInfo where SuppId='" & Trim(txtfields(1)) & "'")
                If Not rs.EOF Then
                    Set rs = getdata("select SuppName from SupplierInfo where SuppId='" & Trim(txtfields(1)) & "'")
                    If Not rs.EOF Then
                       Label12(1).Caption = "" & rs!SuppName
                    End If
                    txtfields(2).SetFocus
                Else
                    cmdsearch_Click
                End If
            End If
           
        Case 2
             If Len(Trim(txtfields(2))) = 0 Then
                Command1_Click
                If Len(txtfields(3)) <> 0 Then
                    txtfields(3) = ""
                     Label12(3).Caption = ""
                End If
             Else
                If Len(txtfields(3)) <> 0 Then
                    txtfields(3) = ""
                     Label12(3).Caption = ""
                End If
                Set rs = getdata("select ClassId from ClassInfo where ClassId='" & Trim(txtfields(2)) & "'")
                
                If Not rs.EOF Then
                    Set rs = getdata("select ClassName from ClassInfo where ClassId='" & Trim(txtfields(2)) & "'")
                    If Not rs.EOF Then
                       Label12(2).Caption = "" & rs!ClassName
                    End If
                        txtfields(3).SetFocus
                Else
                    Command1_Click
                End If
            End If
            
        Case 3
            If Len(Trim(txtfields(3))) = 0 Then
               Command2_Click
             Else
                Set rs = getdata("select SubjectId from subjectInfo where SubjectID ='" & Trim(txtfields(3)) & "'")
                If Not rs.EOF Then
                    Set rs = getdata("select Subjectdsc from subjectInfo where SubjectID ='" & Trim(txtfields(3)) & "'")
                    If Not rs.EOF Then
                       Label12(3).Caption = "" & rs!SubjectDsc
                    End If
                    txtfields(4).SetFocus
                Else
                   Command2_Click
                End If
            End If
            
        Case 4
            txtfields(5).SetFocus
        Case 5
            cmdSave.SetFocus
    End Select
End If
End Sub


Private Sub ShowFlexData()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Set rs = getdata("SELECT RecieveNo,SubjectID,Qty,Notes,Recdate From BookRecievedInfo where suppId='" & Trim(txtfields(1)) & "'and ClassId='" & Trim(txtfields(2)) & "' ")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .TextMatrix(i, 1) = "" & rs!RecieveNo
                .TextMatrix(i, 2) = "" & rs!SubjectID
                Set rs1 = getdata("SELECT SubjectDsc From SubjectInfo where SubjectID='" & rs!SubjectID & "'")
                If Not rs1.EOF Then
                 .TextMatrix(i, 3) = "" & rs1!SubjectDsc
                End If
                .TextMatrix(i, 4) = "" & rs!Qty
                .TextMatrix(i, 5) = "" & rs!Notes
                .TextMatrix(i, 6) = Format("" & rs!Recdate, "DD/mm/yy")
                
                
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
Private Sub get_Maximum()
Dim rs As New ADODB.Recordset
Set rs = getdata("select max(recieveno) from BookRecievedInfo")
If Not rs.EOF Then
        txtfields(0) = rs.Fields(0)
Else
    txtfields(0) = "R-00000001"
End If
End Sub

Private Sub txtfields_LostFocus(Index As Integer)
Select Case Index
Case 2
    If Len(txtfields(3)) <> 0 Then
        txtfields(3) = ""
        Label12(3).Caption = ""
    End If
    Call ShowFlexData
End Select

End Sub
