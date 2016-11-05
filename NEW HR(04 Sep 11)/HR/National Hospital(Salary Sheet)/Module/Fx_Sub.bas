Attribute VB_Name = "Fx_Sub"
    Private con As New ADODB.Connection
    Private cmd As New ADODB.Command
    Private RS As New ADODB.Recordset
    
    Public Comp_Name As String
    Public Comp_Add As String
    Public Comp_Tel As String
    Public Comp_Email As String
    Public Comp_Fax As String
    
    Public rptmode As String
     
Public Sub Company_Nm_Add()
   On Error Resume Next
    con.ConnectionString = strCN.Connection_String
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "Select Co_Nm, Co_Type, Address, Phone='Telephone: ' + Phone, Fax='Fax: ' +Fax , E_mail='E-mail: ' + E_mail from Company_Info"
       
    Set RS = cmd.Execute
       
    With RS
        Comp_Name = !Co_Nm
        Comp_Add = !Address
        Comp_Tel = !Phone
        Comp_Fax = !Fax
        Comp_Email = !E_mail
        .Close
    End With
    
    con.Close
    
End Sub

'Public Sub Report_Header(RptName As Report)
'
'       On ERROR Resume Next
'
'        With RptName
'                .txtCompNm.SetText Comp_Name
'                .txtCompAddress.SetText Comp_Add
'                .txtCompTel.SetText Comp_Tel
'                .txtCompFax.SetText Comp_Fax
'                .txtCompEmail.SetText Comp_Email
'        End With
'End Sub

Public Sub Load_FiscalYr(frm As Form)

'add items to the cboFiscalYr combo of the  calling form
'frm object variable refers to the calling form

    With frm.cboFiscalYr
        .Clear
        .AddItem "2002-2003"
        .AddItem "2003-2004"
        .AddItem "2004-2005"
        .AddItem "2005-2006"
        .AddItem "2006-2007"
        .AddItem "2007-2008"
        .AddItem "2008-2009"
        .AddItem "2009-2010"
        .AddItem "2010-2011"
        .AddItem "2011-2012"
        .AddItem "2012-2013"
        .ListIndex = 0
    End With
    
End Sub
Public Sub Load_Departments(frm As Form)

'add items to the cboFiscalYr combo of the  calling form
'frm object variable refers to the calling form

    With frm.cboDept
        .Clear
        .AddItem "IT Division"
        .AddItem "ACCOUNTS Section"
        .AddItem "ADMINISTRATION"
        .AddItem "Anaesthesia"
        .AddItem "CARDIOLOGY"
        .AddItem "DENTAL"
        .AddItem "EMERGENCY"
        .AddItem "ENGINEERING Section"
        .AddItem "ENT"
        .AddItem "GYNAE & OBST."
        .AddItem "Liebrary"
        .AddItem "MEDICINE"
        .AddItem "NURSING Section"
        .AddItem "OPHTHALMOLOGY"
        .AddItem "Orthopedics"
        .AddItem "PAEDIATRICS"
        .AddItem "PATHOLOGY"
        .AddItem "PERSONNEL Section"
        .AddItem "Physiotherapy"
        .AddItem "RADIOLOGY & IMAGING"
        .AddItem "SKIN VD"
        .AddItem "STORE"
        .AddItem "SURGERY"
        .AddItem "WARD MASTER"
        .ListIndex = 0
    End With
    
End Sub



Public Sub Load_Religion(frm As Form)
  
'add items to the cboReligion combo of the  calling form
'frm object variable refers to the calling form

    With frm.cboReligion
        .Clear
        .AddItem "Islam"
        .AddItem "Sanatan"
        .AddItem "Christian"
        .AddItem "Buddist"
        .AddItem "Others"
        .ListIndex = 0
    End With
    
End Sub
Public Sub LOAD_PAYMENT_RECE_TYPE(frm As Form)
  
'add items to the cboReligion combo of the  calling form
'frm object variable refers to the calling form

    With frm.Combo1(1)
        .Clear
        .AddItem "Cash"
        .AddItem "Check"
        .ListIndex = 0
    End With
    
End Sub
Public Sub LOAD_PAYMENT_TYPE(frm As Form)
  
'add items to the cboReligion combo of the  calling form
'frm object variable refers to the calling form

    With frm.Combo1(3)
        .Clear
        .AddItem "Cash"
        .AddItem "Check"
        .ListIndex = 0
    End With
    
End Sub

Public Sub LOAD_PF_RECEIVE_TYPE(frm As Form)
  
'add items to the cboReligion combo of the  calling form
'frm object variable refers to the calling form

    With frm.cmbReceiveType
        .Clear
        .AddItem "Cash"
        .AddItem "Check"
        .ListIndex = 0
    End With
    
End Sub
Public Sub LOAD_PF_PAYMENT_TYPE(frm As Form)
  
'add items to the cboReligion combo of the  calling form
'frm object variable refers to the calling form

    With frm.cmbPaymentType
        .Clear
        .AddItem "Cash"
        .AddItem "Check"
        .ListIndex = 0
    End With
    
End Sub



Public Sub Load_MonthNm(frm As Form)

'add items to the cboMonth combo of the  calling form
'frm object variable refers to the calling form
    
    With frm.cboMonth
        .Clear
        .AddItem "January"
        .AddItem "February"
        .AddItem "March"
        .AddItem "April"
        .AddItem "May"
        .AddItem "June"
        .AddItem "July"
        .AddItem "August"
        .AddItem "September"
        .AddItem "October"
        .AddItem "November"
        .AddItem "December"
        .Text = MonthName(Month(Now))
    End With
End Sub
Public Sub Load_Yr(frm As Form)

'add items to the cboYear combo of the  calling form
'frm object variable refers to the calling form
    
    With frm.cboYear
        .Clear
        .AddItem "2002"
        .AddItem "2003"
        .AddItem "2004"
        .AddItem "2005"
        .AddItem "2006"
        .AddItem "2007"
        .AddItem "2008"
        .AddItem "2009"
        .AddItem "2010"
        .AddItem "2011"
        .AddItem "2012"
        .AddItem "2013"
        .AddItem "2014"
        .AddItem "2015"
        .Text = YEAR(Now)
    End With

End Sub

Public Sub Load_Department(frm As Form)

'This Sub routine brings Unit name from St_Unit table and
'add items to the cboUnit combo of the  calling form
'frm object variable refers to the calling form

On Error Resume Next

    
    If Not frm.Name = "frmSearch" Then
        frm.cboDept.Clear
    Else
        frm.cboSearch.Clear
    End If

    con.ConnectionString = strCN.Connection_String
    con.Open
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "Select Dept_Nm from St_Dept"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, con, adOpenStatic, adLockOptimistic
   
    If RS.EOF = False Then
    
        Do Until RS.EOF = True
          
            If Not frm.Name = "frmSearch" Then
                  frm.cboDept.AddItem RS.Fields(0)
            Else
                frm.cboSearch.AddItem RS.Fields(0)
            End If
            
            RS.MoveNext
        Loop
        
        frm.cboDept.AddItem ""
        
    End If
    
    RS.Close
    con.Close
    
End Sub
Public Sub Load_BelowType(frm As Form)
    
'This Sub routine brings Accomodation type name from St_HRent3 table and
'add items to the cboBelow_Type combo of the  calling form
'frm object variable refers to the calling form
On Error Resume Next
    frm.cboBelow_Type.Clear

    con.ConnectionString = strCN.Connection_String
    
    con.Open
   
    RS.CursorType = adOpenKeyset
    RS.Open "exec Get_Names 'Below'", con
    If RS.EOF = False Then
    
        Do Until RS.EOF = True
    
            frm.cboBelow_Type.AddItem RS.Fields(0)
            RS.MoveNext
        Loop
        
        frm.cboBelow_Type.ListIndex = 0
        
    End If
    
    RS.Close
    con.Close
    
End Sub
Public Sub Load_Desig(frm As Form)
    
'This Sub routine brings Designation name from St_Desig table and
'add items to the cboDesig combo of the  calling form
'frm object variable refers to the calling form
 On Error Resume Next
    If Not frm.Name = "frmSearch" Then
        frm.cboDesig.Clear
    Else
        frm.cboSearch.Clear
    End If
'---------------------------------------------------------------
    
    con.ConnectionString = strCN.Connection_String
    con.Open
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "Select Designation from St_Desig"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, con, adOpenStatic, adLockOptimistic
    
'---------------------------------------------------------------
    
    If RS.EOF = False Then
    
        RS.MoveFirst
        
        Do Until RS.EOF = True
            
            If Not frm.Name = "frmSearch" Then
                frm.cboDesig.AddItem RS.Fields(0)
            Else
                frm.cboSearch.AddItem RS.Fields(0)
            End If
                        
            RS.MoveNext
        Loop
        
        frm.cboDesig.AddItem ""
        
    End If
    
    RS.Close
    con.Close
    
End Sub
Public Sub Load_PScale(frm As Form)

'This Sub routine brings Pay scale Code from St_PScale table and
'add items to the cboScale combo of the  calling form
'frm object variable refers to the calling form
On Error Resume Next
       
    If Not frm.Name = "frmSearch" Then
        frm.cboScale.Clear
    Else
        frm.cboSearch.Clear
    End If

    '-------------------------------------------------------------------
        con.ConnectionString = strCN.Connection_String
        con.Open
        
        cmd.CommandType = adCmdText
        cmd.CommandText = "Select Scale_Code from St_PayScale"
        RS.CursorLocation = adUseClient
        RS.Open cmd.CommandText, con, adOpenStatic, adLockOptimistic
    '-------------------------------------------------------------------
    
    If RS.EOF = False Then
    
        Do Until RS.EOF = True

            If Not frm.Name = "frmSearch" Then
                frm.cboScale.AddItem RS.Fields(0)
            Else
                frm.cboSearch.AddItem RS.Fields(0)
            End If
            
            RS.MoveNext
        Loop
        
        frm.cboScale.AddItem ""
        
    End If
    
    RS.Close
    con.Close
    
End Sub
Public Sub Load_JbType(frm As Form)
    
'This Sub routine brings Job type name from St_JType table and
'add items to the cboType combo of the  calling form
'frm object variable refers to the calling form
On Error Resume Next
    
    
    If Not frm.Name = "frmSearch" Then
        frm.cboType.Clear
    Else
        frm.cboSearch.Clear
    End If

   
   con.ConnectionString = strCN.Connection_String
    con.Open
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "Select JType_Nm from St_JbType"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, con, adOpenStatic, adLockOptimistic
     
   
    If RS.EOF = False Then
    
        Do Until RS.EOF = True
                        
            If Not frm.Name = "frmSearch" Then
                frm.cboType.AddItem RS.Fields(0)
            Else
                frm.cboSearch.AddItem RS.Fields(0)
            End If
            
            RS.MoveNext
        Loop
        
        frm.cboType.AddItem ""
        
    End If
    
    RS.Close
    con.Close
    
End Sub

Public Sub Load_BankNm(frm As Form)
On Error Resume Next
    frm.cboBank.Clear

    con.ConnectionString = strCN.Connection_String
    con.Open
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "Select Distinct BankName from Emp_Job_Info"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, con, adOpenStatic, adLockOptimistic
    
    
    If RS.EOF = False Then
    
        Do Until RS.EOF = True
            Form2.cboBank.AddItem "" & RS!BankName
            RS.MoveNext
        Loop
        
        frm.cboBank.AddItem ""
        
    End If
    
    RS.Close
    con.Close
    
End Sub

Public Sub Load_Bank_Branch_Nm(frm As Form)
On Error Resume Next
    frm.cboBankBranch.Clear

    con.ConnectionString = strCN.Connection_String
    con.Open
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "Select Distinct Branch_name from Emp_Job_Info"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, con, adOpenStatic, adLockOptimistic
    
    
    If RS.EOF = False Then
    
        Do Until RS.EOF = True
            frm.cboBankBranch.AddItem "" & RS!Branch_name
            RS.MoveNext
        Loop
        
        frm.cboBankBranch.AddItem ""
        
    End If
    
    RS.Close
    con.Close
    
End Sub


Public Sub Load_Subs(frm As Form)
 On Error Resume Next
    frm.cboSub.Clear

    con.ConnectionString = strCN.Connection_String
    
    con.Open
   
    RS.CursorType = adOpenKeyset
    RS.Open "exec Get_Names 'Subs'", con
    If RS.EOF = False Then
    
        Do Until RS.EOF = True
            frm.cboSub.AddItem RS.Fields(0)
            RS.MoveNext
        Loop
        
    End If
    
    RS.Close
    con.Close
    
End Sub

Public Function ChkForQuote(str As String)

' When single coat exits in a string (ie. Father's name, Coxe's Bazar )
' the string is trunketed and saved partially in the SQL Server Database
' like Father's name-->  Father ,Coxe's Bazar---> Coxe.
' It happens due to the difference of strig delimiter in VB and in SQL Server
' To overcome this problem a string needs to be processed before sending it to
' the database.

On Error Resume Next

Dim slposl As Integer, Ln As Integer
Dim str2 As String

str2 = ""
For Ln = 1 To Len(str)
    If Mid$(str, Ln, 1) = "'" Then
        str2 = str2 + "'" + Mid(str, Ln, 1)
    Else
        str2 = str2 + Mid$(str, Ln, 1)
    End If
Next
ChkForQuote = str2
End Function
Public Function Valid_Dt(Dt As String) As String
'This sub routine returns a date  type data in yyyy-mm-dd format

   On Error Resume Next
    
    Dim VD As String
    
    VD = Format(Dt, "yyyy-mm-dd")
    Valid_Dt = VD

End Function

Public Sub Get_Employee(Emp_ID As String, frm As Form, Optional ControlArray As Boolean = False, Optional Index As Integer)

On Error Resume Next
    If Emp_ID = "" Then Exit Sub
    
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter

    con.Open strCN.Connection_String
    Set cmd.ActiveConnection = con
    cmd.CommandType = adCmdText

    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, Emp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamOutput)
    cmd.Parameters.Append Param2

    '----------------------------------------------------------------------------------
    ' Enable PLSQLRSet property

    cmd.Properties("PLSQLRSet") = True

    cmd.CommandText = "{CALL PKG_Misc.Get_Employee_Info(?,?)}"

    Set RS = cmd.Execute

     cmd.Properties("PLSQLRSet") = False
    
    If RS.EOF = False Then
    
            If ControlArray = False Then
                    frm.lblDesig = Space(2) + RS!designation
                    frm.lblName = Space(2) + RS!Emp_Nm
                    frm.lblDept = Space(2) + RS!DEPT_NM
            Else
                    frm.lblDesig(Index) = Space(2) + RS!designation
                    frm.lblName(Index) = Space(2) + RS!Emp_Nm
                    frm.lblDept(Index) = Space(2) + RS!DEPT_NM
                    '---------------------------------------------
               '     frm.lblPF = Space(2) + Rs!Pf_mem_no
            End If
        
    End If
    
    RS.Close
    con.Close


End Sub

Public Sub Clear_Screen()

'This Sub routine clears the text box and combo box exists on the calling form
'frm object variable refers to the calling form

On Error Resume Next

    Dim MyObj As Object
    
    For Each MyObj In Screen.ActiveForm.Controls
    
        If TypeOf MyObj Is TextBox Or _
            TypeOf MyObj Is MSForms.TextBox Or _
                   TypeOf MyObj Is MSForms.ComboBox Or _
                        TypeOf MyObj Is MSForms.Label Then
            MyObj = ""
        End If
        
       ' If TypeOf MyObj Is MaskEdBox Then MyObj = "__/__/____"
        
    Next

End Sub

Public Sub Load_Ln_Type(frm As Form, Ln_Type As String, Optional ControlArray As Boolean = False, Optional Index As Integer)

On Error Resume Next
    con.ConnectionString = strCN.Connection_String
    con.Open
   
    RS.CursorType = adOpenKeyset
    RS.Open "exec Get_Names '" + Ln_Type + "'", con

    If RS.EOF = False Then
        frm.cboLn_Nature(Index).Clear
            If ControlArray = False Then
            
                Do Until RS.EOF = True
                    frm.cboLn_Nature.AddItem RS.Fields(0)
                    RS.MoveNext
                Loop
                
            Else
                
                Do Until RS.EOF = True
                    frm.cboLn_Nature(Index).AddItem RS.Fields(0)
                    RS.MoveNext
                Loop
                
            End If
    End If
        
    RS.Close
    con.Close
    
End Sub
Public Sub Delete_Record(Mode As String, Track_Id As Long, Optional Param_1 As String, Optional Param_2 As String, Optional Param_3 As String)

On Error Resume Next

    con.Open strCN.Connection_String
    
    Set cmd.ActiveConnection = con

    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Delete_Record"
    cmd(1) = Mode
    cmd(2) = Track_Id
    Set RS = cmd.Execute
    
    MsgBox RS!Message, vbOKOnly + vbExclamation
    
    con.Close
    
End Sub


Public Sub Screen_Position(frm As Form)

' set the calling form in a specified position

On Error Resume Next
    
    Dim Frm_ht As Long
    
    Frm_ht = frm.Height

    If Frm_ht > 6300 Then      'form2,form3
            frm.Move (Screen.Width - frm.Width) \ 2, (Screen.Height - frm.Height) / 1.4
    Else
        frm.Move (Screen.Width - frm.Width) \ 2, (Screen.Height - frm.Height) \ 2
    End If

End Sub

Public Sub Destroy(frm As Form)
   On Error Resume Next

    Set frm = Nothing
    
End Sub

Public Sub Show_Tips(ByVal Btn As Integer, ByVal X As Long, ByVal Y As Long, Optional lWd As Long = 3150, Optional lHt As Long = 2500)
    
   On Error Resume Next
    'X= mouse x position,Y= mouse y position
    
    Dim frm As Object, nX As Long
    
    Set frm = Screen.ActiveForm
    
     '   frm.Controls.Add "VB.ListBox", "lstTips", frm
    
    nX = frm.Width - lWd
    nY = frm.Height - lHt

    If X > nX Then X = nX - 500
    If Y > nY Then Y = nY - 500
    
    If Btn = 2 Then         'if left mouse button is clicked
    
        If Not frm!lstTips.Visible = True Then
                 With frm!lstTips
                      .Move X, Y, lWd, lHt
                      .Appearance = 0
                      .BackColor = &HE4FEFD
                      .Visible = True
                      .Clear
                      ''.AddItem frm.Name
                      
                  End With
              Else
                 frm!lstTips.Visible = False
             End If
         Else
             frm!lstTips.Visible = False
         End If

End Sub
Public Function IsNum(ByVal KeyAscii As Integer)
 'it validates number during input
 On Error Resume Next
     Select Case KeyAscii
        Case 8
            IsNum = 8               ' if back space then back space
        Case 13
            IsNum = 13              ' if enter then enter
        Case 46
            IsNum = 46              ' if decimal point then decimal point
        Case Is < 48, Is > 57       ' if non numeral then nothing
            IsNum = 0
        Case Is >= 48, Is <= 57     ' if numeral then numeral
            IsNum = KeyAscii
        
        
    End Select

End Function

Public Sub SetFocus_To(ctrl As Control)
   On Error Resume Next

    'set focus to the control specified
    'with text selection

    With ctrl
        .SelStart = 0
        .SelLength = Len(ctrl)
        .SetFocus
    End With
    
End Sub

Public Function ChunkStr(txt As String, delimiter As String, Optional Chunk_Portion As Boolean = False)
    'It chunks a string and returns according to the specified
    'delimiter and portion
On Error Resume Next
    Dim p As Integer
    
    txt = Trim(txt)
    delimiter = Trim(delimiter)
    
    p = InStr(1, txt, delimiter)       ' position of " . "
    
    
    If Chunk_Portion = 0 Then    'Left
        ChunkStr = Mid(txt, 1, p - 1)
    Else
        ChunkStr = Mid(txt, p + 1, Len(txt))    'Right
    End If
    
    ChunkStr = Trim(ChunkStr)
    
End Function

Public Sub Open_Screen(frm As Form, U_Id As String)
    
    'it uses the Check_permission method of the Security class
    'to validate user's access to the specifed form or scope

    Dim SecureX As New Security

    With SecureX
           .Connstring = strCN.Connection_String
        If .Check_permission(frm.Name, U_Id) = True Then
               frm.Show vbModal
        End If
    End With
    Set SecureX = Nothing
End Sub

Public Function Get_Month_No(Month_Name As String) As Integer
    
    Select Case Month_Name
    
        Case "January": Get_Month_No = 1
        Case "February": Get_Month_No = 2
        Case "March": Get_Month_No = 3
        Case "April": Get_Month_No = 4
        Case "May": Get_Month_No = 5
        Case "June": Get_Month_No = 6
        Case "July": Get_Month_No = 7
        Case "August": Get_Month_No = 8
        Case "September": Get_Month_No = 9
        Case "October": Get_Month_No = 10
        Case "November": Get_Month_No = 11
        Case "December": Get_Month_No = 12
    
    End Select
    

End Function
Public Function Get_Code(SString As String) As String
    Get_Code = Trim(Mid(Trim(SString), InStr(Trim(SString), "~") + 1))
End Function
Public Function Get_Description(SString As String) As String
    Get_Description = Trim(Mid(Trim(SString), 1, InStr(Trim(SString), "~") - 1))
End Function




