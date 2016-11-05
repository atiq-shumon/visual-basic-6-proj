VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmItemUnitInfo 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   645
      Left            =   0
      TabIndex        =   12
      Top             =   5280
      Width           =   8175
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
         Left            =   5160
         TabIndex        =   13
         ToolTipText     =   "Click to insert new information"
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
         Left            =   3210
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
         Left            =   4185
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
         Left            =   6135
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
         Left            =   7110
         TabIndex        =   6
         ToolTipText     =   "Click to Exit"
         Top             =   180
         Width           =   945
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   3180
         Top             =   150
         Width           =   4935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2715
      Left            =   -30
      TabIndex        =   11
      Top             =   2610
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4789
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   13627123
      ForeColor       =   12582912
      BackColorSel    =   12640511
      ForeColorSel    =   8388608
      BackColorBkg    =   -2147483637
      GridColor       =   -2147483637
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000001&
      FillColor       =   &H00E0E0E0&
      Height          =   915
      Left            =   -30
      ScaleHeight     =   855
      ScaleWidth      =   8145
      TabIndex        =   8
      Top             =   -30
      Width           =   8205
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Unit Information Setup"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEF0F7&
         Height          =   315
         Left            =   2760
         TabIndex        =   14
         Top             =   210
         Width           =   3540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   1725
      Left            =   0
      TabIndex        =   0
      Top             =   870
      Width           =   8145
      Begin VB.TextBox txtfields 
         Height          =   525
         Index           =   2
         Left            =   1770
         MaxLength       =   100
         TabIndex        =   2
         ToolTipText     =   "Insert Marks Category"
         Top             =   1005
         Width           =   5445
      End
      Begin VB.TextBox txtfields 
         Height          =   315
         Index           =   1
         Left            =   1770
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Insert Marks Category"
         Top             =   585
         Width           =   5415
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1770
         MaxLength       =   2
         TabIndex        =   7
         Top             =   195
         Width           =   1515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   1110
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   630
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Code"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmItemUnitInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objcom As New DSLComFram.clsCommon
Dim objRs As New ADODB.Recordset


Private Sub cmdDelete_Click()
If Len(txtfields(0)) = 0 Then
                MsgBox "Item Unit Code Mandatory", vbInformation, App.Title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(txtfields(1)) = 0 Then
                MsgBox "Item Unit Name Mandatory", vbInformation, App.Title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT unit_code from item_unit_info  WHERE (unit_code= '" & txtfields(0) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "No such Code Exists..Please Verify.", vbInformation, cmp
               cmdnew.SetFocus
               Exit Sub
            End If
            
           If MsgBox("Are you sure to delete?", vbYesNo + vbInformation, cmp) = vbYes Then
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
                MsgBox "Item Unit Code Mandatory", vbInformation, App.Title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(txtfields(1)) = 0 Then
                MsgBox "Item Unit Name Mandatory", vbInformation, App.Title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT unit_code from item_unit_info  WHERE (unit_code= '" & txtfields(0) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "No such Code Exists..Please Verify.", vbInformation, cmp
               cmdnew.SetFocus
               Exit Sub
            End If
            
            edit

       MsgBox "Updated successfully.", vbInformation, cmp
       Call ShowFlexData
       cmdnew.SetFocus

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
Dim RS As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.Connection

cmd.ActiveConnection = objmyCon
Set RS = getdata("select max(unit_code+1)from item_unit_info")
If Not RS.EOF Then
    txtfields(0) = IIf(IsNull(RS(0)) = True, "01", Format(RS(0), "00"))
Else
    txtfields(0) = "01"
End If


    txtfields(1) = ""

txtfields(1).SetFocus
End Sub


Private Sub cmdSAVE_Click()
          If Len(txtfields(0)) = 0 Then
                MsgBox "Item Unit Code Mandatory", vbInformation, App.Title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(txtfields(1)) = 0 Then
                MsgBox "Item Unit Name Mandatory", vbInformation, App.Title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT unit_code from item_unit_info  WHERE (unit_code= '" & txtfields(0) & "')", objmyCon)
            
            If Not objRs.EOF Then
               MsgBox "Same already Code Exists..Please Verify.", vbInformation, cmp
               cmdnew.SetFocus
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
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 1)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 2, Format(Trim(txtfields(0)), "00"))
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 80, Trim(txtfields(1)))
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 100, Trim(txtfields(2)))
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param5

    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param6
    
    


    
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL s_U_d_item_unit_info(?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
     cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
   
End Sub
Private Sub edit()
    Dim cmd As New ADODB.Command
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 2)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 2, Format(Trim(txtfields(0)), "00"))
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 80, Trim(txtfields(1)))
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 100, Trim(txtfields(2)))
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param5

    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param6
    
    


    
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL s_U_d_item_unit_info(?,?,?,?,?,?)}"
    
   
    
    cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
   
End Sub
Private Sub delete()
  Dim RS As New ADODB.Recordset
  Dim cmd As New ADODB.Command
  Set RS = Nothing

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 3)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 2, Format(Trim(txtfields(0)), "00"))
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 80, Trim(txtfields(1)))
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 100, Trim(txtfields(2)))
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param5

    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param6
    
    


    
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL s_U_d_item_unit_info(?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
   
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
Dim RS As New ADODB.Recordset
Set RS = getdata("select max (unit_code+1)from item_unit_info")
If Not RS.EOF Then
    txtfields(0) = IIf(IsNull(RS(0)) = True, "01", Format(RS(0), "00"))
Else
    txtfields(0) = "01"
End If
With MSFlexGrid1
    .Rows = 1
    .Cols = 3
    .Col = 0: .Text = " Code"
    .Col = 1: .Text = " Title"
    .Col = 2: .Text = " Remarks"
    
    .ColWidth(0) = 800
    .ColWidth(1) = 5500
    .ColWidth(2) = 5500
    
End With
Call ShowFlexData
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title

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
Dim RS As New ADODB.Recordset

Select Case Index
    Case 0
        If Len(Trim(txtfields(0))) = 0 Then Exit Sub
      
            txtfields(0) = Format(txtfields(0), "00000")
          
            Set RS = getdata("SELECT mcategoryDsc,Note,EntryBy,Entrydate from Markscategory WHERE (McategoryID= '" & txtfields(0) & "')")
                 If Not RS.EOF Then
                        txtfields(1) = RS!mcategoryDsc
                        txtfields(2) = RS!Note
                        txtfields(3) = RS!EntryBy
'                        dtpic = rs!Format(Entrydate, "dd/mmm/yyyy")
                End If
        
End Select
End Sub
Private Sub ShowFlexData()
On Error GoTo errdes
Dim RS As New ADODB.Recordset
Set RS = objcom.Get_RS("SELECT unit_code as Code ,unit_name as Title,remarks  From item_unit_info order by unit_code", objmyCon)
If Not RS.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until RS.EOF
            MSFlexGrid1.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = RS!Code
                .TextMatrix(i, 1) = RS!Title
                .TextMatrix(i, 2) = "" & RS!remarks
                i = i + 1
            RS.MoveNext
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

