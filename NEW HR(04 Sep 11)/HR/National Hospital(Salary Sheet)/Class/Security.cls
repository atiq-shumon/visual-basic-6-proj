VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Security"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private dPW As New EnDecode.clsEndecoder
Private mvarConnstring As String
Private mvarUser_ID As String 'local copy
Private mvarUser_Name As String 'local copy
Private mvarPassword As String 'local copy
Private mvarConfirm_Password As String 'local copy
Private mvarUser_Group As String
Private mvarAccess As AccessPerm
Private mvarCREATED_BY As String
Private mvarAccess_Area As String
Private mvarUserRole As String

Public Enum AccessPerm
    Yes = 1
    No = 0
End Enum

Public Enum Priviles
    NotGiven = 0
    Given = 1
End Enum

Public Enum Permission
    pSingle = 0
    pAll = 1
End Enum
Private mvarOld_Password As String 'local copy
Private mvarPSL As String 'local copy
Private mvarUDT As Date 'local copy
Private mvarU_TYPE As String 'local copy
Public Property Let UserRole(ByVal vData As String)
    mvarUserRole = vData
End Property
Public Property Get UserRole() As String
    UserRole = mvarU_TYPE
End Property
Public Property Let U_TYPE(ByVal vData As String)
    mvarU_TYPE = mvarUserRole
End Property
Public Property Get U_TYPE() As String
    U_TYPE = mvarU_TYPE
End Property
Public Property Let UDT(ByVal vData As Date)
    mvarUDT = vData
End Property
Public Property Get UDT() As Date
    UDT = mvarUDT
End Property
Public Property Let PSL(ByVal vData As String)
    mvarPSL = vData
End Property
Public Property Get PSL() As String
    PSL = mvarPSL
End Property
Public Property Let Connstring(ByVal vData As String)
   mvarConnstring = vData
End Property
Public Property Let Old_Password(ByVal vData As String)
    mvarOld_Password = vData
End Property
Public Property Get Old_Password() As String
   Old_Password = mvarOld_Password
End Property
Public Property Let Confirm_Password(ByVal vData As String)
    mvarConfirm_Password = vData
End Property
Public Sub CreateUser()
End Sub
Public Property Let Password(ByVal vData As String)
    mvarPassword = vData
End Property
Public Property Let User_Name(ByVal vData As String)
    mvarUser_Name = vData
End Property
Public Property Get User_Name() As String
    User_Name = mvarUser_Name
End Property
Public Property Let User_id(ByVal vData As String)
    mvarUser_ID = vData
End Property

Public Property Let User_Group(ByVal vData As String)
    mvarUser_Group = vData
End Property
Public Property Get User_Group() As String
    User_Group = mvarUser_Group
End Property
Public Property Let Access(ByVal vData As AccessPerm)
    mvarAccess = vData
End Property

Public Property Get Access() As AccessPerm
    Access = mvarAccess
End Property
Public Property Let Access_Area(ByVal vData As String)
    mvarAccess_Area = vData
End Property
Public Property Let CREATED_BY(ByVal vData As String)
    mvarCREATED_BY = vData
End Property
Public Function ValidateUser(User_id As String, Pw As String, UserRole As String) As Boolean
On Error Resume Next
'validate user id and password during login
    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New ADODB.Recordset
    
    If Pw = Empty Then
        MsgBox "Enter password please!", vbCritical + vbOKOnly, "Sequrity Manager"
        ValidateUser = False
        Exit Function
    End If
    
'     With dPW                   'Encode password to match stored User password
'         .InputString = Pw
'         .Encode
'         Pw = .OutputString    'Encoded password
'     End With
'    ''------------------------------------------------
      
    'Conn.Open mvarConnString
    
    conn.Open mvarConnstring
    
    
    Set cmd.ActiveConnection = conn

    cmd.CommandType = adCmdText
    cmd.CommandText = "select User_Pass from soft_pass where U_Id='" & User_id & "' and U_type='" & UserRole & "'"
    cmd.Properties("IRowsetChange") = True
    cmd.Properties("Updatability") = 7
    
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
    ''------------------------------------------------
    
    If Not RS.Fields(0) = Pw Then
        MsgBox "Invalid User or Password!", vbCritical + vbOKOnly, "Sequrity Manager"
        ValidateUser = False
    Else
        ValidateUser = True
    End If
End Function
Public Function GetAll_User() As Recordset
On Error Resume Next
    
    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New ADODB.Recordset
    Dim tempRs As New ADODB.Recordset
     
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    '-------------------------------------------
'    cmd.CommandType = adCmdText
'    cmd.CommandText = "Exec POP_User"
'    Rs.CursorLocation = adUseClient
'    Rs.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "select u_id ,u_name as User Name,U_Type ,(select case a.Access" & _
                            "when 0 then 'No'" & _
                            "WHEN 1 then 'Yes' END " & _
                            "from soft_pass a where u_name ='DBA' and a.u_id=soft_pass.u_id) as Access" & _
                            ",user_pass  from soft_pass  where u_name='DBA'"
    
    cmd.CommandText = "Select * from Soft_Pass where u_name!='DBA'"
    cmd.Properties("IRowsetChange") = True
    cmd.Properties("Updatability") = 7
    RS.CursorLocation = adUseClient
    
   
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
        
    
    
    '-------------------------------------------
        
        
   Set GetAll_User = RS
    

End Function
Public Function GetAccess_Area() As Recordset
On Error Resume Next
    
    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New ADODB.Recordset
    
     
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "Select Distinct Access_Area from Soft_Bag"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
        
    Set GetAccess_Area = RS
       
End Function
Public Function Delete_User(User_id As String)

    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset
    
        conn.Open mvarConnstring
        Set cmd.ActiveConnection = conn
        cmd.CommandType = adCmdText
        cmd.CommandText = "Delete from Soft_Pass where U_id='" + User_id + "'"
        
        RS.CursorLocation = adUseClient
        RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic

        'MsgBox "User Id: " + User_Id + " no more exists!", vbOKOnly + vbExclamation

End Function

Public Function GetUserType() As Recordset
'On ERROR Resume Next
    
'    Dim Conn As New Connection
'    Dim cmd As New Command
'    Dim tmpRs As New ADODB.Recordset
        
        
'    Set GetAccess_Area = Rs
       
End Function
Public Function Get_Previleges(pType As Priviles, Access_Area As String, User_id As String) As Recordset
   On Error Resume Next
 
    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New ADODB.Recordset
    
     
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    
    cmd.CommandType = adCmdText
    
    Access_Area = Trim(Access_Area)
    User_id = Trim(User_id)
    
    If pType = Given Then
            cmd.CommandText = "exec Software_Previleges 'In','" + User_id + "','" + Access_Area + "'"
    Else
           cmd.CommandText = "exec Software_Previleges 'Not_In','" + User_id + "','" + Access_Area + "'"
    End If
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
        
    Set Get_Previleges = RS
   
End Function
Public Function Give_Permission(pType As String, Access_Area As String, Permit_On As String, User_id As String)
   On Error Resume Next
    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New ADODB.Recordset
    
    Access_Area = Trim(Access_Area)
    Permit_On = Trim(Permit_On)
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    
    cmd.CommandText = "exec give_pmt '" + pType + "','" + Access_Area + "','" + Permit_On + "','" + User_id + "'"
    
    
    
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
End Function

'Public Function Revoke_Permission(pType As String, Access_Area As String, Revoke As String, User_id As String)
'   On ERROR Resume Next
'    Dim Conn As New Connection
'    Dim Cmd As New Command
'    Dim Rs As New ADODB.Recordset
'
'
'    Access_Area = Trim(Access_Area)
'    Revoke = Trim(Revoke)
'    User_id = Trim(User_id)
'
'
'    Conn.Open mvarConnString
'
'    Set Cmd.ActiveConnection = Conn
'
'    Cmd.CommandType = adCmdText
'    Cmd.CommandText = "exec Revoke_pmt '" + pType + "','" + Access_Area + "','" + Revoke + "','" + User_id + "'"
'    Rs.CursorLocation = adUseClient
'    Rs.Open Cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
'
'
'    ''----------------------------
'
'    Dim Param1 As New Parameter
'    Dim Param2 As New Parameter
'    Dim Param3 As New Parameter
'    Dim Param4 As New Parameter
'
'    ''Here Param1 is declared as Parameter type variable for taking input parameter
'    '' and Param2 is for retrieving records fetched by the command text.
'
'    Conn.Open mvarConnString
'
'    Set Cmd.ActiveConnection = Conn
'    Cmd.CommandType = adCmdText
'
'
'    Set Param1 = Cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarDesig_Code)
'    Cmd.Parameters.Append Param1
'
'    Set Param2 = Cmd.CreateParameter("param2", adSmallInt, adParamOutput)
'    Cmd.Parameters.Append Param2
'
'    Set Param3 = Cmd.CreateParameter("param3", adSmallInt, adParamOutput)
'    Cmd.Parameters.Append Param3
'
'    Set Param4 = Cmd.CreateParameter("param4", adSmallInt, adParamOutput)
'    Cmd.Parameters.Append Param4
'    '----------------------------------------------------------------------------------
'    ' Enable PLSQLRSet property
'
'    Cmd.Properties("PLSQLRSet") = True
'    Cmd.CommandText = "{CALL PKG_Security.Revoke_pmt(?,?,?,?)}"
'
'Revoke_pmt '" + pType + "','" + Access_Area + "','" + Revoke + "','" + User_id + "'"
'
'    Set Rs = Cmd.Execute
'
'    ' Disable PLSQLRSet property
'    Cmd.Properties("PLSQLRSet") = False
'
'    Set GetX = Rs
'
'
'
'    ''--------------------------------
'
'
'End Function

Public Function Check_permission(Screen_Name As String, User_id As String) As Boolean
    
   On Error Resume Next
    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New ADODB.Recordset
    
'     Conn.Open mvarConnString
'    Set Cmd.ActiveConnection = Conn
'    Cmd.CommandType = adCmdText
'    Cmd.CommandText = "select dbo.Chk_Perm ('" + Screen_Name + "','" + User_id + "')"
'
'    Set rs = Cmd.Execute

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn

    cmd.CommandType = adCmdText
    cmd.CommandText = "select count(code)from permit where code in (select code from SOFT_BAG where scr_no='" & Screen_Name & "') and u_id='" & User_id & "'"
    cmd.Properties("IRowsetChange") = True
    cmd.Properties("Updatability") = 7
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
    If RS.Fields(0) = 0 Then
        MsgBox "Access Restricted   ", vbCritical + vbOKOnly, "Sequrity Manager"
        Check_permission = False
    Else
        Check_permission = True
    End If
    
    'Conn.Close
    
End Function

'Public Function Open_Screen(ByVal frm As Form, U_Id As String)
'    If Check_permission(frm, U_Id) = True Then
'        frm.Show vbModal
'    End If
'
'End Function
'Public Function Save1() As Boolean
'    Dim conn As New Connection
'    Dim cmd As New Command
'    Dim Rs As New Recordset
'
'    If Not mvarConfirm_Password = Empty Then
'
'            If StrComp(mvarConfirm_Password, mvarPassword, vbBinaryCompare) = 0 Then
'                With dPW
'                    .InputString = mvarConfirm_Password
'                    .Encode
'                    mvarConfirm_Password = .OutputString    'Encoded password
'           ''         MsgBox mvarConfirm_Password
'                End With
'
'            ''-------------Send to Database-----------------
'
'                        conn.Open mvarConnString
'                        Set cmd.ActiveConnection = conn
'                        cmd.CommandType = adCmdStoredProc
'                        cmd.CommandText = "pro_soft_pass"
'                        cmd(1) = mvarUser_ID
'                        cmd(2) = mvarUser_Name
'                        cmd(3) = mvarUser_Group
'                        cmd(4) = mvarConfirm_Password
'                        cmd(5) = mvarCREATED_BY
'                        cmd(6) = mvarAccess
'                        cmd(7) = "S"            'Operation Mode
'
'                        Set Rs = cmd.Execute
'                        MsgBox Rs!Message, vbOKOnly + vbExclamation
'
'                Save = True
'
'            Else
'
'                MsgBox "Password does not match", vbCritical + vbOKOnly, "Uesr"
'                Save = False
'            End If
'   End If
'End Function
Public Function Save() As Boolean
On Error GoTo Errdes
    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim Param3 As New ADODB.Parameter
    Dim Param4 As New ADODB.Parameter
    Dim Param5 As New ADODB.Parameter
    Dim Param6 As New ADODB.Parameter
    Dim Param7 As New ADODB.Parameter
    Dim Param8 As New ADODB.Parameter
    Dim Param9 As New ADODB.Parameter
    
If Not mvarConfirm_Password = Empty Then

    If StrComp(mvarConfirm_Password, mvarPassword, vbBinaryCompare) = 0 Then
                With dPW
                    .InputString = mvarConfirm_Password
                    .Encode
                    mvarConfirm_Password = .OutputString    'Encoded password
                    mvarPassword = .OutputString
           ''         MsgBox mvarConfirm_Password
                End With

            ''-------------Send to Database-----------------



        conn.Open mvarConnstring
    
        Set cmd.ActiveConnection = conn
        cmd.CommandType = adCmdText
    
        '----------------------------------------------------------------------------------
        Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarPSL)
        cmd.Parameters.Append Param1
    
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, mvarUser_ID)
        cmd.Parameters.Append Param2
    
        Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 50, mvarUser_Name)
        cmd.Parameters.Append Param3
    
        Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 50, mvarU_TYPE)
        cmd.Parameters.Append Param4
    
        Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 50, mvarPassword)
        cmd.Parameters.Append Param5
    
    
        Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 50, mvarCREATED_BY)
        cmd.Parameters.Append Param6
        
        Set Param7 = cmd.CreateParameter("param7", adDate, adParamInput, 50, mvarUDT)
        cmd.Parameters.Append Param7
        
    
        Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 50, mvarAccess)
        cmd.Parameters.Append Param8
    
        Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 50, mvarConfirm_Password)
        cmd.Parameters.Append Param9


    '----------------------------------------------------------------------------------

        ' Enable PLSQLRSet property
        cmd.Properties("PLSQLRSet") = True
        
        cmd.CommandText = "{CALL Password_UserSetup_save(?,?,?,?,?,?,?,?,?)}"
        
        Set RS = cmd.Execute
        
        ' Disable PLSQLRSet property
        cmd.Properties("PLSQLRSet") = False
        
        Set RS = cmd.Execute
          
        Save = True

    Else

        MsgBox "Password does not match", vbCritical + vbOKOnly, "Uesr"
        Save = False
    End If

End If
Exit Function
Errdes:
MsgBox Err.Description, vbInformation, "Daffodil Sotware Ltd"
End Function

