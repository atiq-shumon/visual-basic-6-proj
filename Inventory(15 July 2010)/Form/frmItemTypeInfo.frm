VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmItemTypeInfo 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   705
      Left            =   0
      TabIndex        =   13
      Top             =   7140
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
         TabIndex        =   14
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   7
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
         TabIndex        =   8
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
      Height          =   4605
      Left            =   -30
      TabIndex        =   12
      Top             =   2610
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8123
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
      TabIndex        =   9
      Top             =   -30
      Width           =   8205
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Store Item Type Information Setup"
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
         TabIndex        =   15
         Top             =   210
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   1725
      Left            =   0
      TabIndex        =   4
      Top             =   870
      Width           =   8145
      Begin VB.ComboBox Cbocategory 
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   3075
      End
      Begin VB.TextBox txtfields 
         Height          =   525
         Index           =   2
         Left            =   1770
         MaxLength       =   100
         TabIndex        =   3
         ToolTipText     =   "Insert Marks Category"
         Top             =   1005
         Width           =   5445
      End
      Begin VB.TextBox txtfields 
         Height          =   315
         Index           =   1
         Left            =   1770
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "Insert Marks Category"
         Top             =   585
         Width           =   5415
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   5970
         MaxLength       =   2
         TabIndex        =   1
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category Code"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   210
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1110
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type  Name"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   630
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type Code"
         Height          =   195
         Left            =   4980
         TabIndex        =   10
         Top             =   210
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmItemTypeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objcom As New DSLComFram.clsCommon
Dim objRs As New ADODB.Recordset


Private Sub Cbocategory_Click()
    ShowFlexData
End Sub

Private Sub cmdDelete_Click()
            If Len(txtfields(0)) = 0 Then
                MsgBox "Group Code Mandatory", vbInformation, App.title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(txtfields(1)) = 0 Then
                MsgBox "Group Name Mandatory", vbInformation, App.title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            If Len(Cbocategory) = 0 Then
                MsgBox "Item Category Mandatory", vbInformation, App.title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT group_code from item_group_info  WHERE (group_code= '" & txtfields(0) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "No such Code Exists..Please Verify.", vbInformation, cmp
               cmdnew.SetFocus
               Exit Sub
            End If
            
             Set objRs = objcom.Get_RS("SELECT group_code from item_info  WHERE (group_code= '" & txtfields(0) & "')", objmyCon)
            
            If Not objRs.EOF Then
               MsgBox "Already Used You can't delete..Please Verify.", vbInformation, cmp
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
                MsgBox "Group Code Mandatory", vbInformation, App.title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(txtfields(1)) = 0 Then
                MsgBox "Group Name Mandatory", vbInformation, App.title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            If Len(Cbocategory) = 0 Then
                MsgBox "Item Category Mandatory", vbInformation, App.title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT group_code from item_group_info  WHERE (group_code= '" & txtfields(0) & "')", objmyCon)
            
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
Set RS = getdata("select max(type_code+1)from item_type_info")
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
                MsgBox "Type Code Mandatory", vbInformation, App.title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(txtfields(1)) = 0 Then
                MsgBox "Type Name Mandatory", vbInformation, App.title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            If Len(Cbocategory) = 0 Then
                MsgBox "Item Group Mandatory", vbInformation, App.title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT type_code from item_type_info  WHERE (type_code= '" & txtfields(0) & "')", objmyCon)
            
            If Not objRs.EOF Then
               MsgBox "Same Type Code  already Exists..Please Verify.", vbInformation, cmp
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
    Dim Param7 As New Parameter
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 1)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 2, Format(Trim(txtfields(0)), "00"))
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 80, Trim(txtfields(1)))
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 2, Trim(Get_Code(Cbocategory)))
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 100, Trim(txtfields(2)))
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param6

    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param7
    
    


    
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  s_U_d_item_Type_info(?,?,?,?,?,?,?)}"
    
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
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 2)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 2, Format(Trim(txtfields(0)), "00"))
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 80, Trim(txtfields(1)))
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 2, Trim(Get_Code(Cbocategory)))
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 100, Trim(txtfields(2)))
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param6

    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param7
    
    


    
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  s_U_d_item_type_info(?,?,?,?,?,?,?)}"
    
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
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 3)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 2, Format(Trim(txtfields(0)), "00"))
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 80, Trim(txtfields(1)))
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 2, Trim(Get_Code(Cbocategory)))
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 100, Trim(txtfields(2)))
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param6

    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param7
    
    


    
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  s_U_d_item_Type_info(?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
     cmd.Execute
    

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
load_category
Dim RS As New ADODB.Recordset
Set RS = getdata("select max (type_code+1)from item_type_info")
If Not RS.EOF Then
    txtfields(0) = IIf(IsNull(RS(0)) = True, "01", Format(RS(0), "00"))
Else
    txtfields(0) = "01"
End If
With MSFlexGrid1
    .Rows = 1
    .Cols = 5
    .Col = 0: .Text = " Code"
    .Col = 1: .Text = " Title"
    .Col = 2: .Text = " Type Code"
    .Col = 3: .Text = "Category Title"
    .Col = 4: .Text = " Remarks"
    
    .ColWidth(0) = 800
    .ColWidth(1) = 5500
    .ColWidth(2) = 0
    .ColWidth(3) = 5500
    .ColWidth(4) = 5500
    
    
End With
Call ShowFlexData
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title

End Sub

Private Sub load_category()
 Set objRs = objcom.Get_RS("SELECT cate_name,cate_code from item_cate_info where cate_code='" & CategoryCode & "'", objmyCon)
 Cbocategory.Clear
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       Cbocategory.AddItem objRs(0) + "~" + objRs(1)
       objRs.MoveNext
    Loop
 End If
   
  
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
'On Error GoTo errdes
Dim RS As New ADODB.Recordset
'Set RS = objcom.Get_RS("SELECT a.group_code as Code ,a.group_name as Title,a.cate_code,(select b.cate_name from item_cate_info b where b.cate_code=a.cate_code) as cat_title, a.remarks  From item_group_info a", objmyCon)
Set RS = objcom.Get_RS("SELECT a.type_code as Code ,a.type_name as Title,a.cate_code, b.cate_name, a.remarks  From item_type_info a,item_cate_info b where a.cate_code=b.cate_code and a.cate_code='" & Get_Code(CategoryCode) & "'order by a.type_code asc", objmyCon)
If Not RS.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until RS.EOF
            MSFlexGrid1.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = RS(0)
                .TextMatrix(i, 1) = RS(1)
                .TextMatrix(i, 2) = "" & RS(2)
               .TextMatrix(i, 3) = "" & RS(3)
                .TextMatrix(i, 4) = "" & RS(4)
                
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

On Error GoTo errdes
txtfields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
Cbocategory.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)) + "-" + Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3))
Exit Sub
errdes:
'MsgBox err.Description, vbInformation, App.Title


End Sub

