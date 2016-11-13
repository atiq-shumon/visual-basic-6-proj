VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMarksCat 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   0
      TabIndex        =   12
      Top             =   5130
      Width           =   6495
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
         Left            =   2520
         TabIndex        =   4
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
         Left            =   3495
         TabIndex        =   3
         ToolTipText     =   "Click to Save"
         Top             =   180
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
         Left            =   4455
         TabIndex        =   5
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
         Left            =   5430
         TabIndex        =   6
         ToolTipText     =   "Click to Exit"
         Top             =   180
         Width           =   945
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   2490
         Top             =   150
         Width           =   3915
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   915
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6405
      TabIndex        =   8
      Top             =   0
      Width           =   6465
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mark's Category Description"
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
         Left            =   1260
         TabIndex        =   13
         Top             =   270
         Width           =   3225
      End
      Begin VB.Image Image1 
         Height          =   1050
         Left            =   -30
         Picture         =   "frmMarksCat.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   6495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   900
      Width           =   6405
      Begin VB.TextBox txtfields 
         Height          =   465
         Index           =   2
         Left            =   1380
         MaxLength       =   80
         TabIndex        =   2
         ToolTipText     =   "Insert Short Note"
         Top             =   1020
         Width           =   4515
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   1
         Left            =   1380
         MaxLength       =   80
         TabIndex        =   1
         ToolTipText     =   "Insert Marks Category"
         Top             =   630
         Width           =   4485
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1380
         TabIndex        =   7
         Top             =   210
         Width           =   1515
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   1020
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category Name"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   165
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2415
      Left            =   -30
      TabIndex        =   14
      Top             =   2730
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   4260
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
Attribute VB_Name = "frmMarksCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDelete_Click()
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
Set cmd.ActiveConnection = con
Set rs = getdata("select CategoryID from SubjectMarksDistribution where term_code='" & Trim(txtfields(0)) & "'")
 If Not rs.EOF Then
    MsgBox "Already Used...You can't delete", vbInformation, cmp
    Exit Sub
  End If
 If Len(Trim(txtfields(1))) <> 0 Then
        If (MsgBox("Are You sure to delete ?", vbYesNo + vbCritical) = vbYes) Then
             Set rs = getdata("select * from subjectmarksdistribution where categoryID= '" & Mid(Trim(txtfields(0)), 1, 5) & "'")
                If rs.EOF Then
                    cmd.CommandType = adCmdText
                    cmd.CommandText = "Delete from Markscategory  where (McategoryID= '" & Mid(Trim(txtfields(0)), 1, 5) & "') "
                    cmd.Execute
                    MsgBox "Delete successfully .", vbInformation, App.Title
                    txtfields(1) = ""
                    txtfields(2) = ""
                    Call ShowFlexData
                    
                Else
                    MsgBox "Data is existed for this category.", vbInformation, App.Title
                    Exit Sub
                End If
        Else
            Exit Sub
        End If
Else
        MsgBox "Data Doesn't exist for deletion.", vbCritical, "School Management System"
        Exit Sub
End If


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
Set rs = getdata("select max (McategoryID+1)from Markscategory")
If Not rs.EOF Then
    txtfields(0) = IIf(IsNull(rs(0)) = True, "00001", Format(rs(0), "00000"))
Else
    txtfields(0) = "00001"
End If

For i = 1 To 2
    txtfields(i) = ""
Next
dtpic = Format(dtpic, "##/##/##")
txtfields(1).SetFocus
End Sub


Private Sub cmdSAVE_Click()
If Len(txtfields(1)) = 0 Then
    MsgBox "Please Enter Category Name.", vbInformation, App.Title
    txtfields(1).SetFocus
    Exit Sub
End If
 
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "Markscategorydes"
cmd(1) = Format(Trim(txtfields(0)), "00000")
cmd(2) = Trim(txtfields(1))
cmd(3) = Trim(txtfields(2))
cmd(4) = "DSL"
'cmd(5) = IIf(dtpic = "__/__/__", Format(Date, "DD MMM YYYY"), Format(dtpic, "dd mmm yyyy"))
cmd(5) = Format(Date, "dd mmm yyyy")
cmd.Execute
MsgBox "Save successfully.", vbInformation, "Student Management System"
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
Set rs = getdata("select max (McategoryID+1)from Markscategory")
If Not rs.EOF Then
    txtfields(0) = IIf(IsNull(rs(0)) = True, "00001", Format(rs(0), "00000"))
Else
    txtfields(0) = "00001"
End If
With MSFlexGrid1
    .Rows = 1
    .Cols = 3
    .Col = 0: .Text = "         ID #"
    .Col = 1: .Text = "Marks Category Name"
    .Col = 2: .Text = " Remarks "
    
    .ColWidth(0) = 800
    .ColWidth(1) = 2800
    .ColWidth(2) = 2600
    
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
Set rs = getdata("SELECT McategoryID,mcategoryDsc,Note From Markscategory")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = rs!McategoryID
                .TextMatrix(i, 1) = rs!mcategoryDsc
                .TextMatrix(i, 2) = rs!Note
                
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

