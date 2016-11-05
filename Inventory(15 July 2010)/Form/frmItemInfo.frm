VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmItemInfo 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   645
      Left            =   0
      TabIndex        =   17
      Top             =   7230
      Width           =   10245
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
         Left            =   7200
         TabIndex        =   10
         ToolTipText     =   "Click to Edit information"
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
         Left            =   5250
         TabIndex        =   8
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
         Left            =   6225
         TabIndex        =   7
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
         Left            =   8175
         TabIndex        =   11
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
         Left            =   9150
         TabIndex        =   9
         ToolTipText     =   "Click to Close"
         Top             =   180
         Width           =   945
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   5220
         Top             =   150
         Width           =   4935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3615
      Left            =   -30
      TabIndex        =   16
      Top             =   3630
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   6376
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
      ScaleWidth      =   10245
      TabIndex        =   13
      Top             =   -30
      Width           =   10305
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Information Setup"
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
         Left            =   4140
         TabIndex        =   18
         Top             =   210
         Width           =   2925
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   2745
      Left            =   0
      TabIndex        =   12
      Top             =   870
      Width           =   10215
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   3
         Left            =   8790
         MaxLength       =   5
         TabIndex        =   5
         Top             =   1575
         Width           =   945
      End
      Begin VB.ComboBox CboGroup 
         Height          =   315
         Left            =   6390
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   630
         Width           =   3345
      End
      Begin VB.ComboBox CboUnit 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1590
         Width           =   3315
      End
      Begin VB.CommandButton Command1 
         Caption         =   ":::"
         Height          =   285
         Left            =   3330
         TabIndex        =   22
         Top             =   210
         Width           =   375
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   6390
         TabIndex        =   20
         Top             =   180
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   503
         _Version        =   393216
         Format          =   22609921
         CurrentDate     =   38865
      End
      Begin VB.ComboBox CboType 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   630
         Width           =   3285
      End
      Begin VB.TextBox txtfields 
         Height          =   495
         Index           =   2
         Left            =   1500
         MaxLength       =   100
         TabIndex        =   6
         ToolTipText     =   "Insert  Remarks here"
         Top             =   2115
         Width           =   8235
      End
      Begin VB.TextBox txtfields 
         Height          =   315
         Index           =   1
         Left            =   1500
         MaxLength       =   150
         TabIndex        =   3
         ToolTipText     =   "Insert Product Name"
         Top             =   1125
         Width           =   8235
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1500
         MaxLength       =   5
         TabIndex        =   0
         Top             =   195
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min Level"
         Height          =   195
         Index           =   2
         Left            =   7980
         TabIndex        =   26
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type Code"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   25
         Top             =   660
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Code"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   24
         Top             =   1620
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Code"
         Height          =   195
         Index           =   0
         Left            =   5430
         TabIndex        =   23
         Top             =   690
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entry Date"
         Height          =   195
         Index           =   1
         Left            =   5430
         TabIndex        =   21
         Top             =   210
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Specification"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   19
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   15
         Top             =   1170
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   14
         Top             =   210
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmItemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objcom As New DSLComFram.clsCommon
Dim objRs As New ADODB.Recordset
Dim clkMode As Integer


Private Sub Cbocategory_Click()

  load_group


End Sub

Private Sub CboGroup_Click()
    ShowFlexData
End Sub

Private Sub CboType_Click()
    load_group
End Sub

Private Sub cmdDelete_Click()
      If Len(txtfields(0)) = 0 Then
                MsgBox "Item Code Mandatory", vbInformation, App.title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(txtfields(1)) = 0 Then
                MsgBox "Item Name Mandatory", vbInformation, App.title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            If Len(cboType) = 0 Then
                MsgBox "Item Type Mandatory", vbInformation, App.title
                cboType.SetFocus
                Exit Sub
            End If
            
            If Len(CboUnit) = 0 Then
                MsgBox "Item Unit Mandatory", vbInformation, App.title
                CboUnit.SetFocus
                Exit Sub
            End If
            
            If Len(CboGroup) = 0 Then
                MsgBox "Item Group Mandatory", vbInformation, App.title
                CboGroup.SetFocus
                Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT item_code from item_info  WHERE (item_code= '" & txtfields(0) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "No such Code Exists..Please Verify.", vbInformation, cmp
               cmdnew.SetFocus
               Exit Sub
            End If
            
            If Len(txtfields(3)) = 0 Then
               txtfields(3) = 0
            End If
            
            
            Set objRs = objcom.Get_RS("SELECT itemId from PurchaseSub  WHERE (to_number(itemId)= to_number('" & Trim(txtfields(0)) & "'))", objmyCon)
            
               
            If Not objRs.EOF Then
               MsgBox "Already Used..You can't delete...Please Verify.", vbInformation, cmp
               cmdnew.SetFocus
               Exit Sub
            End If
             If Len(txtfields(3)) = 0 Then
               txtfields(3) = 0
            End If
             
            
           If MsgBox("Are you sure to delete?", vbYesNo + vbInformation, cmp) = vbYes Then
             delete
           Else
             Exit Sub
           End If

       MsgBox "Deleted successfully.", vbInformation, cmp
       cmdnew_Click
       Call ShowFlexData
       cmdnew.SetFocus
End Sub

Private Sub CmdEdit_Click()
     If Len(txtfields(0)) = 0 Then
                MsgBox "Item Code Mandatory", vbInformation, App.title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(txtfields(1)) = 0 Then
                MsgBox "Item Name Mandatory", vbInformation, App.title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            If Len(cboType) = 0 Then
                MsgBox "Item Type Mandatory", vbInformation, App.title
                cboType.SetFocus
                Exit Sub
            End If
            
            If Len(CboUnit) = 0 Then
                MsgBox "Item Unit Mandatory", vbInformation, App.title
                CboUnit.SetFocus
                Exit Sub
            End If
            
            If Len(CboGroup) = 0 Then
                MsgBox "Item Group Mandatory", vbInformation, App.title
                CboGroup.SetFocus
                Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT item_code from item_info  WHERE (item_code= '" & txtfields(0) & "')", objmyCon)
            
            If objRs.EOF Then
               MsgBox "No such Code Exists..Please Verify.", vbInformation, cmp
               cmdnew.SetFocus
               Exit Sub
            End If
            
            If Len(txtfields(3)) = 0 Then
               txtfields(3) = 0
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
Set RS = objcom.Get_RS("select to_char(nvl(max(to_number(item_code)),0)+1,'0000') as max_number from item_info", objmyCon)
If Not RS.EOF Then
    txtfields(0) = RS(0)
End If


    txtfields(1) = ""

txtfields(1).SetFocus
End Sub


Private Sub cmdSAVE_Click()
          If Len(txtfields(0)) = 0 Then
                MsgBox "Item Code Mandatory", vbInformation, App.title
                txtfields(0).SetFocus
                Exit Sub
            End If
            
            
            If Len(txtfields(1)) = 0 Then
                MsgBox "Item Name Mandatory", vbInformation, App.title
                txtfields(1).SetFocus
                Exit Sub
            End If
            
            If Len(cboType) = 0 Then
                MsgBox "Item Type Mandatory", vbInformation, App.title
                cboType.SetFocus
                Exit Sub
            End If
            
            If Len(CboUnit) = 0 Then
                MsgBox "Item Unit Mandatory", vbInformation, App.title
                CboUnit.SetFocus
                Exit Sub
            End If
            
            If Len(CboGroup) = 0 Then
                MsgBox "Item Group Mandatory", vbInformation, App.title
                CboGroup.SetFocus
                Exit Sub
            End If
            
            Set objRs = objcom.Get_RS("SELECT item_code from item_info  WHERE (item_code= '" & txtfields(0) & "')", objmyCon)
            
            If Not objRs.EOF Then
               MsgBox "Same  Code  already Exists..Please Verify.", vbInformation, cmp
               cmdnew.SetFocus
               Exit Sub
            End If
            
            If Len(txtfields(3)) = 0 Then
               txtfields(3) = 0
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
    
   
    Set RS = objcom.Get_RS("select to_char(nvl(max(to_number(item_code)),0)+1,'0000') as max_number from item_info", objmyCon)
    If Not RS.EOF Then
      txtfields(0) = RS(0)
    End If

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 1)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 4, Format(Trim(txtfields(0)), "0000"))
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 150, Trim(txtfields(1)))
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 2, Trim(CategoryCode))
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 3, Trim(Get_Code(CboGroup)))
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 2, Trim(Get_Code(CboUnit)))
    cmd.Parameters.Append Param6
   
    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 18, Val(txtfields(3)))
    cmd.Parameters.Append Param7
   
        
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 100, Trim(txtfields(2)))
    cmd.Parameters.Append Param8
    
    Set Param9 = cmd.CreateParameter("param9", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param9

    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 10, Get_Code(cboType))
    cmd.Parameters.Append Param11
    
    


    
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_item_info(?,?,?,?,?,?,?,?,?,?,?)}"
    
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
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    Dim Param10 As New Parameter
    Dim Param11 As New Parameter
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 2)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 5, Trim(txtfields(0)))
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 150, Trim(txtfields(1)))
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 2, Trim(CategoryCode))
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 3, Trim(Get_Code(CboGroup)))
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 2, Trim(Get_Code(CboUnit)))
    cmd.Parameters.Append Param6
   
    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 18, Trim(txtfields(3)))
    cmd.Parameters.Append Param7
   
        
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 100, Trim(txtfields(2)))
    cmd.Parameters.Append Param8
    
    Set Param9 = cmd.CreateParameter("param9", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param9

    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 10, Trim(Get_Code(cboType)))
    cmd.Parameters.Append Param11
       
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_item_info(?,?,?,?,?,?,?,?,?,?,?)}"
    
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
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    Dim Param10 As New Parameter
    Dim Param11 As New Parameter
   

    Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 1, 3)
    cmd.Parameters.Append Param1


    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 5, Trim(txtfields(0)))
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 150, Trim(txtfields(1)))
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 2, Trim(CategoryCode))
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 2, Trim(Get_Code(CboGroup)))
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 2, Trim(Get_Code(CboUnit)))
    cmd.Parameters.Append Param6
   
    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 18, Trim(txtfields(3)))
    cmd.Parameters.Append Param7
   
        
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 100, Trim(txtfields(2)))
    cmd.Parameters.Append Param8
    
    Set Param9 = cmd.CreateParameter("param9", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param9

    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 10, strAppUser)
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 10, Trim(Get_Code(cboType)))
    cmd.Parameters.Append Param11
       
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL  S_U_d_item_info(?,?,?,?,?,?,?,?,?,?,?)}"
    
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
clkMode = 1
load_type
load_unit
DTPicker1.value = Date
Dim RS As New ADODB.Recordset
Set RS = objcom.Get_RS("select to_char(nvl(max(to_number(item_code)),0)+1,'0000') as max_number from item_info", objmyCon)
If Not RS.EOF Then
    txtfields(0) = RS(0)
End If
With MSFlexGrid1
    .Rows = 1
    .Cols = 10
    .Col = 0: .Text = " Code"
    .Col = 1: .Text = " Title"
    .Col = 2: .Text = " Unit Code"
    .Col = 3: .Text = "Unit"
    .Col = 4: .Text = "Group Code"
    .Col = 5: .Text = "Group "
    .Col = 6: .Text = "cat Code"
    .Col = 7: .Text = "Category"
    .Col = 8: .Text = "Re-Order"
    .Col = 9: .Text = " Remarks"
    
    .ColWidth(0) = 800
    .ColWidth(1) = 4000
    .ColWidth(2) = 0
    .ColWidth(3) = 1500
    .ColWidth(4) = 0
    .ColWidth(5) = 1500
    .ColWidth(6) = 0
    .ColWidth(7) = 1500
    .ColWidth(8) = 900
    .ColWidth(9) = 5000
    
End With
Call ShowFlexData
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title

End Sub

Private Sub load_type()

 Set objRs = objcom.Get_RS("SELECT type_code,type_name from item_type_info where cate_code='" & Get_Code(CategoryCode) & "'", objmyCon)
 cboType.Clear
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       cboType.AddItem objRs(1) + "~" + objRs(0)
       objRs.MoveNext
    Loop
 End If

  
End Sub
Private Sub load_unit()

 Set objRs = objcom.Get_RS("SELECT unit_code,unit_name from item_unit_info", objmyCon)
 CboUnit.Clear
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       CboUnit.AddItem objRs(1) + "~" + objRs(0)
       objRs.MoveNext
    Loop
  End If
   

End Sub
Private Sub load_group()

 Set objRs = objcom.Get_RS("SELECT group_code,group_name from item_group_info where type_code='" & Trim(Get_Code(cboType)) & "'", objmyCon)
 CboGroup.Clear
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       CboGroup.AddItem objRs(1) + "~" + objRs(0)
       objRs.MoveNext
    Loop
 End If
End Sub

Private Sub MSFlexGrid1_SelChange()
   MSFlexGrid1_Click
End Sub
Private Sub txtfields_Change(Index As Integer)
   Select Case Index
          Case 3
               If Not IsNumeric(txtfields(3)) Then
                   txtfields(3) = ""
               End If
  End Select
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
'If KeyAscii = 13 Then
'    Select Case Index
'    Case Index
'        If Index <> 3 Then
''            txtfields(Index + 1).SetFocus
'        Else
'            dtpic.SetFocus
'        End If
'    End Select
'End If
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
With MSFlexGrid1
    .Rows = 1
    .Cols = 10
    .Col = 0: .Text = " Code"
    .Col = 1: .Text = " Title"
    .Col = 2: .Text = " Unit Code"
    .Col = 3: .Text = "Unit"
    .Col = 4: .Text = "Group Code"
    .Col = 5: .Text = "Group "
    .Col = 6: .Text = "cat Code"
    .Col = 7: .Text = "Type"
    .Col = 8: .Text = "Re-Order"
    .Col = 9: .Text = " Remarks"
    
    .ColWidth(0) = 800
    .ColWidth(1) = 4000
    .ColWidth(2) = 0
    .ColWidth(3) = 1500
    .ColWidth(4) = 0
    .ColWidth(5) = 1500
    .ColWidth(6) = 0
    .ColWidth(7) = 1500
    .ColWidth(8) = 900
    .ColWidth(9) = 5000
    
End With
Dim RS As New ADODB.Recordset
'Set RS = objcom.Get_RS("SELECT a.group_code as Code ,a.group_name as Title,a.cate_code,(select b.cate_name from item_cate_info b where b.cate_code=a.cate_code) as cat_title, a.remarks  From item_group_info a", objmyCon)
Set RS = objcom.Get_RS("SELECT a.item_code  as Code ,a.item_name   as Title,a.unit_code,c.unit_name,a.group_code,d.group_name,a.type_code, b.type_name, a. re_ord_lbl,a.remarks  From item_info a,item_type_info b,item_unit_info c,item_group_info d  where a.type_code=b.type_code  and a.unit_code=c.unit_code and a.group_code=d.group_code and  a.cate_code='" & CategoryCode & "' and a.type_code='" & Get_Code(cboType) & "'  and a.group_code='" & Get_Code(CboGroup) & "' order by a.item_code desc", objmyCon)
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
                .TextMatrix(i, 5) = "" & RS(5)
                .TextMatrix(i, 6) = "" & RS(6)
                .TextMatrix(i, 7) = "" & RS(7)
                .ColAlignment(8) = 0
                .TextMatrix(i, 8) = "" & RS(8)
                
                
                i = i + 1
            RS.MoveNext
        Loop
         MSFlexGrid1.Rows = 1000
    End With
Else
    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 1000
 End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.title
End Sub
Private Sub MSFlexGrid1_Click()

On Error GoTo errdes
txtfields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
CboUnit.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)) + "~" + Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
'CboGroup.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)) + "~" + Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4))
'Cbocategory.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)) + "~" + Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))
txtfields(3) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8))
Exit Sub
errdes:
'MsgBox err.Description, vbInformation, App.Title
End Sub
