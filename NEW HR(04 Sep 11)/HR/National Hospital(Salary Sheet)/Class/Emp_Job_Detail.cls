VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmp_Job_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarConnstring As String
Private mvarEmp_ID As String
Private mvarEmp_Nm As String
Private mvarDESIGNATION As String
Private mvarDept As String
Private mvarJobType As String
Private mvarJdate As Date
Private mvarPdate As Date
Private mvarFile_ref_number As String
Private mvarService_bk_number As String
Private mvarScale_code As String
Private mvarMode_of_payment As String
Private mvarBasic_Sal As Double
Private mvarbank_name As String
Private mvarBranch_name As String
Private mvarAcc_No As String
Private mvarPf_mem As String
Private mvarPf_mem_no As String
Private mvarReseponsibility As String
Private mvarPre_Balance As Double
Private mvarBMDCREGI As String 'local copy
Private mvarEmpClass As String 'local copy
Private mvarEmpPositionSerial As Integer
Private mvarEmpDesignationLevel As Integer
Public Property Let EmpDesignationLevel(ByVal vData As Integer)
       mvarEmpDesignationLevel = vData
End Property
Public Property Get EmpDesignationLevel() As Integer
      EmpDesignationLevel = mvarEmpDesignationLevel
End Property

Public Property Let EmpPositionSerial(ByVal vData As Integer)
       mvarEmpPositionSerial = vData
End Property
Public Property Get EmpPositionSerial() As Integer
      EmpPositionSerial = mvarEmpPositionSerial
End Property
Public Property Let EmpClass(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmpClass = 5
    mvarEmpClass = vData
End Property


Public Property Get EmpClass() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmpClass
    EmpClass = mvarEmpClass
End Property



Public Property Let BMDCREGI(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BMDCREGI = 5
    mvarBMDCREGI = vData
End Property


Public Property Get BMDCREGI() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BMDCREGI
    BMDCREGI = mvarBMDCREGI
End Property



Public Property Get Emp_Nm() As String
   Emp_Nm = mvarEmp_Nm
End Property
Public Property Let Reseponsibility(ByVal vData As String)
    mvarReseponsibility = vData
End Property
Public Property Get Reseponsibility() As String
    Reseponsibility = mvarReseponsibility
End Property
Public Function Save() As Variant
On Error GoTo Errdes
If mvarEmp_ID = Empty Then Exit Function

    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim Param3 As New ADODB.Parameter
    Dim Param4 As New ADODB.Parameter
    Dim Param5 As New ADODB.Parameter
    Dim Param6 As New ADODB.Parameter
    Dim Param7 As New ADODB.Parameter
    Dim Param8 As New ADODB.Parameter
    Dim Param9 As New ADODB.Parameter
    Dim Param10 As New ADODB.Parameter
    Dim Param11 As New ADODB.Parameter
    Dim Param12 As New ADODB.Parameter
    Dim Param13 As New ADODB.Parameter
    Dim Param14 As New ADODB.Parameter
    Dim Param15 As New ADODB.Parameter
    Dim Param16 As New ADODB.Parameter
    Dim Param17 As New ADODB.Parameter
    Dim Param18 As New ADODB.Parameter
    Dim Param19 As New ADODB.Parameter
    Dim Param20 As New ADODB.Parameter
    Dim Param21 As New ADODB.Parameter
    Dim Param22 As New ADODB.Parameter
    
   
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 50, mvarDESIGNATION)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 50, mvarDept)
    cmd.Parameters.Append Param3
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 50, mvarJobType)
    cmd.Parameters.Append Param4

    Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, , mvarJdate)
    cmd.Parameters.Append Param5
    
    
    Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, , mvarPdate)
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 50, mvarFile_ref_number)
    cmd.Parameters.Append Param7

    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 50, mvarService_bk_number)
    cmd.Parameters.Append Param8
    
    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 200, mvarReseponsibility)
    cmd.Parameters.Append Param9
    
    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 10, mvarScale_code)
    cmd.Parameters.Append Param10

    Set Param11 = cmd.CreateParameter("param11", adDouble, adParamInput, 10, mvarBasic_Sal)
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adVarChar, adParamInput, 2, mvarMode_of_payment)
    cmd.Parameters.Append Param12
    
    Set Param13 = cmd.CreateParameter("param13", adVarChar, adParamInput, 50, mvarbank_name)
    cmd.Parameters.Append Param13
    
    Set Param14 = cmd.CreateParameter("param14", adVarChar, adParamInput, 100, mvarBranch_name)
    cmd.Parameters.Append Param14
    
    Set Param15 = cmd.CreateParameter("param15", adVarChar, adParamInput, 10, Trim(mvarAcc_No))
    cmd.Parameters.Append Param15
    
    Set Param16 = cmd.CreateParameter("param16", adVarChar, adParamInput, 10, mvarPf_mem)
    cmd.Parameters.Append Param16
    
    Set Param17 = cmd.CreateParameter("param17", adVarChar, adParamInput, 10, mvarPf_mem_no)
    cmd.Parameters.Append Param17

    Set Param18 = cmd.CreateParameter("param18", adDouble, adParamInput, 10, mvarPre_Balance)
    cmd.Parameters.Append Param18
    
    Set Param19 = cmd.CreateParameter("param19", adVarChar, adParamInput, 10, mvarBMDCREGI)
    cmd.Parameters.Append Param19
    
    Set Param20 = cmd.CreateParameter("Param20", adVarChar, adParamInput, 2, mvarEmpClass)
    cmd.Parameters.Append Param20
    
    Set Param21 = cmd.CreateParameter("Param21", adInteger, adParamInput, 5, mvarEmpPositionSerial)
    cmd.Parameters.Append Param21
   
    Set Param22 = cmd.CreateParameter("Param22", adInteger, adParamInput, 5, mvarEmpDesignationLevel)
    cmd.Parameters.Append Param22
    
    ''----------------------------------------------------------------------------------

'' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Pkg_Emp_Info.Emp_Job_Info_IU(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?,?,?)}"
   
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
Exit Function
Errdes:
MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamOutput)
    cmd.Parameters.Append Param2
    '----------------------------------------------------------------------------------
    ' Enable PLSQLRSet property
    
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL PKG_Emp_Info.Emp_Job_Info_GetX(?,?)}"

    Set RS = cmd.Execute

    ' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
    
    Set GetX = RS
        
    If Not (RS.EOF And RS.BOF) Then
       mvarEmp_ID = RS("Emp_Id")
       mvarEmp_Nm = RS("Emp_Nm")
       mvarDESIGNATION = RS("Designation")
       mvarDept = RS("Dept_Nm")
       mvarJobType = RS("JType_Nm")
       mvarJdate = RS("Jdate")
       mvarPdate = RS("Pdate")
       mvarFile_ref_number = RS("File_ref_no")
       mvarService_bk_number = RS("Service_bk_nm")
       mvarReseponsibility = RS("Responsibility")
       mvarScale_code = RS("Scale_Code")
       mvarBasic_Sal = RS("Basic_sal")
       mvarMode_of_payment = RS("Mode_of_payment")
       mvarbank_name = RS("bankname")
       mvarBranch_name = RS("Branch_name")
       mvarAcc_No = RS("Acc_No")
       mvarPf_mem = RS("Pf_mem")
       mvarPf_mem_no = RS("Pf_mem_no")
       mvarPre_Balance = RS("Pre_Bal")
       mvarBMDCREGI = RS("BMDCREGI")
       mvarEmpClass = RS("EmpClass")
       mvarEmpPositionSerial = RS("Emp_Position")
       
    Else
    
       mvarEmp_ID = ""
       mvarDESIGNATION = ""
       mvarDept = ""
       mvarJobType = ""
       mvarJdate = ""
       mvarPdate = ""
       mvarFile_ref_number = ""
       mvarService_bk_number = ""
      ' mvarresponcsibility = ""
       mvarScale_code = ""
       mvarBasic_Sal = ""
       mvarMode_of_payment = ""
       mvarbank_name = ""
       mvarBranch_name = ""
       mvarAcc_No = ""
       mvarPf_mem = ""
       mvarPf_mem_no = ""
       mvarPre_Balance = ""
       mvarBMDCREGI = ""
       mvarEmpClass = ""
       mvarEmpPositionSerial = ""
    End If
    
End Function
Public Function Delete() As Variant
On Error GoTo DeleteError
Dim conn As New Connection
Dim cmd As New Command

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Emp_Job_info_Del"
    cmd(1) = mvarEmp_ID
    cmd.Execute
    Exit Function
DeleteError:
    Err.Raise vbObjectError + 513, "Delete Error", "Problems in Deleting Data from the Emp_Per_info Table" + vbCr + vbCr + Err.Description
End Function
Public Property Let Pre_Balance(ByVal vData As Double)
    mvarPre_Balance = vData
End Property
Public Property Get Pre_Balance() As Double
    Pre_Balance = mvarPre_Balance
End Property
Public Property Let Pf_mem_no(ByVal vData As String)
    mvarPf_mem_no = vData
End Property
Public Property Get Pf_mem_no() As String
    Pf_mem_no = mvarPf_mem_no
End Property
Public Property Let Pf_mem(ByVal vData As String)
    mvarPf_mem = vData
End Property
Public Property Get Pf_mem() As String
    Pf_mem = mvarPf_mem
End Property
Public Property Let Acc_No(ByVal vData As String)
    mvarAcc_No = vData
End Property
Public Property Get Acc_No() As String
    Acc_No = mvarAcc_No
End Property
Public Property Let Branch_name(ByVal vData As String)
    mvarBranch_name = vData
End Property
Public Property Get Branch_name() As String
    Branch_name = mvarBranch_name
End Property
Public Property Let bank_name(ByVal vData As String)
    mvarbank_name = vData
End Property
Public Property Get bank_name() As String
    bank_name = mvarbank_name
End Property
Public Property Let Mode_of_payment(ByVal vData As String)
    mvarMode_of_payment = vData
End Property
Public Property Get Mode_of_payment() As String
    Mode_of_payment = mvarMode_of_payment
End Property
Public Property Let Basic_Sal(ByVal vData As Double)
    mvarBasic_Sal = vData
End Property
Public Property Get Basic_Sal() As Double
    Basic_Sal = mvarBasic_Sal
End Property
Public Property Let Scale_code(ByVal vData As String)
    mvarScale_code = vData
End Property
Public Property Get Scale_code() As String
    Scale_code = mvarScale_code
End Property
Public Property Let Service_bk_number(ByVal vData As String)
    mvarService_bk_number = vData
End Property
Public Property Get Service_bk_number() As String
    Service_bk_number = mvarService_bk_number
End Property
Public Property Let File_ref_number(ByVal vData As String)
    mvarFile_ref_number = vData
End Property
Public Property Get File_ref_number() As String
    File_ref_number = mvarFile_ref_number
End Property
Public Property Let Pdate(ByVal vData As Date)
    mvarPdate = vData
End Property
Public Property Get Pdate() As Date
     Pdate = mvarPdate
End Property
Public Property Let Jdate(ByVal vData As Date)
    mvarJdate = vData
End Property
Public Property Get Jdate() As Date
    Jdate = mvarJdate
End Property
Public Property Let JobType(ByVal vData As String)
    mvarJobType = vData
End Property
Public Property Get JobType() As String
    JobType = mvarJobType
End Property
Public Property Let Dept(ByVal vData As String)
    mvarDept = vData
End Property
Public Property Get Dept() As String
    Dept = mvarDept
End Property
Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property
Public Property Let designation(ByVal vData As String)
    mvarDESIGNATION = vData
End Property
Public Property Get designation() As String
    designation = mvarDESIGNATION
End Property
Public Property Let Emp_ID(ByVal vData As String)
    mvarEmp_ID = vData
End Property
Public Property Get Emp_ID() As String
    Emp_ID = mvarEmp_ID
End Property
Public Function Get_Employee() As Recordset
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
    cmd.CommandText = "{CALL PKG_MISC.Get_Employee_Info(?,?)}"

    Set RS = cmd.Execute

    ' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
    
    Set Get_Employee = RS
    
    If Not (RS.EOF And RS.BOF) Then
       mvarEmp_ID = RS("Emp_Id")
       mvarEmp_Nm = RS("Emp_Nm")
       mvarDESIGNATION = RS("Designation")
       mvarDept = RS("Dept_Nm")
       
    Else
       mvarEmp_ID = ""
       mvarEmp_Nm = ""
       mvarDESIGNATION = ""
       mvarDept = ""
    End If
    
End Function



Public Sub Show_Message()
       
    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset
    
    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "Select * from Temp"
    cmd.Properties("IRowsetChange") = True
    cmd.Properties("Updatability") = 7
    
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic

    MsgBox RS.Fields(0), vbInformation + vbOKOnly, "Message"

End Sub
