VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarConnstring As String 'local copy
Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property
Public Function Employee_Information(mvarEmp_ID As String) As Recordset
  On Error Resume Next

    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
   
    ''Here Param1 is declared as Parameter type variable for taking input parameter
    '' and Param2 is for retrieving records fetched by the command text.
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamOutput)
    cmd.Parameters.Append Param2
    '----------------------------------------------------------------------------------
    ' Enable PLSQLRSet property
    
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL PKG_RPT_CR_VB.Emp_Job_Info_GetX(?,?)}"
   

    Set RS = cmd.Execute

             
     Set Employee_Information = RS

    ' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
    
End Function
Public Function GetX() As Recordset
    
   On Error Resume Next
    
    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    
    
    
    ''Here Param1 is declared as Parameter type variable for taking input parameter
    '' and Param2 is for retrieving records fetched by the command text.
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, "1001")
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamOutput)
    cmd.Parameters.Append Param2
    '----------------------------------------------------------------------------------
    ' Enable PLSQLRSet property
    
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL PKG_RPT_CR_VB.Rpt_Emp_Info(?,?)}"

    Set RS = cmd.Execute

    ' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
    
    Set GetX = RS
      
End Function
Public Function Emp_Salary_Payment_Mode() As Recordset
'On Error Resume Next
    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset
        
    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn

    cmd.CommandType = adCmdText
    
    cmd.CommandText = "SELECT EMP_INFO.EMP_ID, EMP_INFO.EMP_NM, " + _
                        " EMP_INFO.EMP_FAT_NM, EMP_INFO.DOB, " + _
                        " EMP_JOB_INFO.DEPT, EMP_JOB_INFO.BASIC_SAL," + _
                        " EMP_JOB_INFO.MODE_OF_PAYMENT, " + _
                        " EMP_JOB_INFO.BANKNAME, EMP_JOB_INFO.BRANCH_NAME, " + _
                        " ST_DEPT.DEPT_NM " + _
                        " From EMP_JOB_INFO, Emp_info, ST_DEPT " + _
                        " WHERE EMP_JOB_INFO.EMP_ID = EMP_INFO.EMP_ID AND " + _
                        " EMP_JOB_INFO.DEPT = ST_DEPT.DEPT_CODE AND " + _
                        " EMP_JOB_INFO.Mode_of_payment = 0"
    
    cmd.Properties("IRowsetChange") = True
    cmd.Properties("Updatability") = 7
    
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic

    If Not (RS.EOF Or RS.BOF) Then
        Set Emp_Salary_Payment_Mode = RS
        Exit Function
    End If
   
End Function
