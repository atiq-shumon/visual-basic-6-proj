VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Salary_Monthly"
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
Private mvarPAY_MONTH As Integer
Private mvarPAY_YEAR As Integer
Private mvarDESIGNATION As String
Private mvarDEPT_NM As String
Private mvarATTN As Integer
Private mvarBASIC As Currency
Private mvarH_RENT As Currency
Private mvarMED As Currency
Private mvarCONV As Currency
Private mvarTFN As Currency
Private mvarTELEPHONE As Currency
Private mvarDA As Currency
Private mvarARREAR As Currency
Private mvarPF_Ded As Currency
Private mvarR_STAMP As Currency
Private mvarADV_AMOUNT  As Currency
Private mvarLN_AMOUNT As Currency
Private mvarRemarks   As String
Private mvarOTHERS_ALLOWANCE As Currency 'local copy
Private mvarOTHERS_DEDUCTION As Currency 'local copy
Private mvarDisburse_What As String
Private mvarDisburse_To As String
Private mvarDisburse_Method As String
Private mvarDisburse_date As String
Private mvarSalay_Month_Get As String 'local copy
Private mvarSalary_YearGet As String 'local copy
Private mvarSalaryType As String
Public Property Let Salary_YearGet(ByVal vData As String)
    mvarSalary_YearGet = vData
End Property
Public Property Get Salary_YearGet() As String
    Salary_YearGet = mvarSalary_YearGet
End Property
Public Property Let Salay_Month_Get(ByVal vData As String)
    mvarSalay_Month_Get = vData
End Property
Public Property Get Salay_Month_Get() As String
    Salay_Month_Get = mvarSalay_Month_Get
End Property
Public Property Get Emp_Nm() As String
    Emp_Nm = mvarEmp_Nm
End Property
Public Property Get DESIGNATION() As String
    DESIGNATION = mvarDESIGNATION
End Property
Public Property Get Department() As String
    Department = mvarDEPT_NM
End Property
Public Property Let Others_Ded(ByVal vData As Currency)
    mvarOTHERS_DEDUCTION = vData
End Property

Public Property Get Others_Ded() As Currency
    Others_Ded = mvarOTHERS_DEDUCTION
End Property
Public Property Let OTHERS_ALLOWANCE(ByVal vData As Currency)
    mvarOTHERS_ALLOWANCE = vData
End Property
Public Property Get OTHERS_ALLOWANCE() As Currency
    OTHERS_ALLOWANCE = mvarOTHERS_ALLOWANCE
End Property
Public Property Let Emp_ID(ByVal vData As String)
   mvarEmp_ID = vData
End Property
Public Property Get Emp_ID() As String
   Emp_ID = mvarEmp_ID
End Property
Public Property Let PAY_MONTH(ByVal vData As Long)
   mvarPAY_MONTH = vData
End Property
Public Property Get PAY_MONTH() As Long
   PAY_MONTH = mvarPAY_MONTH
End Property
Public Property Let PAY_YEAR(ByVal vData As Long)
   mvarPAY_YEAR = vData
End Property
Public Property Get PAY_YEAR() As Long
   PAY_YEAR = mvarPAY_YEAR
End Property
Public Property Let ATTN(ByVal vData As Long)
   mvarATTN = vData
End Property
Public Property Get ATTN() As Long
   ATTN = mvarATTN
End Property
Public Property Let BASIC(ByVal vData As Currency)
   mvarBASIC = vData
End Property
Public Property Get BASIC() As Currency
   BASIC = mvarBASIC
End Property
Public Property Let H_RENT(ByVal vData As Currency)
   mvarH_RENT = vData
End Property
Public Property Get H_RENT() As Currency
   H_RENT = mvarH_RENT
End Property
Public Property Let MED(ByVal vData As Currency)
   mvarMED = vData
End Property
Public Property Get MED() As Currency
   MED = mvarMED
End Property
Public Property Let CONV(ByVal vData As Currency)
   mvarCONV = vData
End Property
Public Property Get CONV() As Currency
   CONV = mvarCONV
End Property
Public Property Let TFN(ByVal vData As Currency)
   mvarTFN = vData
End Property
Public Property Get TFN() As Currency
   TFN = mvarTFN
End Property
Public Property Let DA(ByVal vData As Currency)
   mvarDA = vData
End Property
Public Property Get DA() As Currency
   DA = mvarDA
End Property
Public Property Let PF_Ded(ByVal vData As Currency)
   mvarPF_Ded = vData
End Property
Public Property Get PF_Ded() As Currency
   PF_Ded = mvarPF_Ded
End Property
Public Property Let ADV_AMOUNT(ByVal vData As Currency)
   mvarADV_AMOUNT = vData
End Property
Public Property Get ADV_AMOUNT() As Currency
   ADV_AMOUNT = mvarADV_AMOUNT
End Property
Public Property Let R_STAMP(ByVal vData As Currency)
   mvarR_STAMP = vData
End Property
Public Property Get R_STAMP() As Currency
   R_STAMP = mvarR_STAMP
End Property
Public Property Let Ln_Amount(ByVal vData As Currency)
   mvarLN_AMOUNT = vData
End Property
Public Property Get Ln_Amount() As Currency
   Ln_Amount = mvarLN_AMOUNT
End Property
'Public Property Let PF_Loan(ByVal vData As Currency)
'   mvarPF_Loan = vData
'End Property
'Public Property Get PF_Loan() As Currency
'   PF_Loan = mvarPF_Loan
'End Property
Public Property Let TELEPHONE(ByVal vData As Currency)
   mvarTELEPHONE = vData
End Property
Public Property Get TELEPHONE() As Currency
   TELEPHONE = mvarTELEPHONE
End Property
Public Property Let ARREAR(ByVal vData As Currency)
   mvarARREAR = vData
End Property
Public Property Get ARREAR() As Currency
   ARREAR = mvarARREAR
End Property
Public Property Let Remarks(ByVal vData As String)
   mvarRemarks = vData
End Property
Public Property Get Remarks() As String
   Remarks = mvarRemarks
End Property
Public Property Let Connstring(ByVal vData As String)
   mvarConnstring = vData
End Property
Public Property Get Connstring() As String
   Connstring = mvarConnstring
End Property
Public Property Let Disburse_What(ByVal vData As String)
    mvarDisburse_What = vData
End Property
Public Property Let Disburse_To(ByVal vData As String)
    mvarDisburse_To = vData
End Property
Public Property Let Disburse_Method(ByVal vData As String)
    mvarDisburse_Method = vData
End Property
Public Property Let Disburse_Date(ByVal vData As String)
    mvarDisburse_date = vData
End Property
Public Property Let SalaryType(ByVal vData As String)
    mvarSalaryType = vData
End Property
Public Property Get SalaryType() As String
    SalaryType = mvarSalaryType
End Property

Public Function Save()
   If mvarEmp_ID = Empty Then Exit Function

    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim Rs As New ADODB.Recordset

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
    
        
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adNumeric, adParamInput, 2, mvarPAY_MONTH)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adNumeric, adParamInput, 4, mvarPAY_YEAR)
    cmd.Parameters.Append Param3
    Set Param4 = cmd.CreateParameter("param4", adNumeric, adParamInput, 2, mvarATTN)
    cmd.Parameters.Append Param4
 
    Set Param5 = cmd.CreateParameter("param5", adNumeric, adParamInput, 10, mvarBASIC)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adNumeric, adParamInput, 10, mvarH_RENT)
    cmd.Parameters.Append Param6
    Set Param7 = cmd.CreateParameter("param7", adNumeric, adParamInput, 10, mvarMED)
    cmd.Parameters.Append Param7

    Set Param8 = cmd.CreateParameter("param8", adNumeric, adParamInput, 10, mvarCONV)
    cmd.Parameters.Append Param8
    
    Set Param9 = cmd.CreateParameter("param9", adNumeric, adParamInput, 10, mvarTFN)
    cmd.Parameters.Append Param9
    
    Set Param10 = cmd.CreateParameter("param10", adNumeric, adParamInput, 10, mvarTELEPHONE)
    cmd.Parameters.Append Param10

    Set Param11 = cmd.CreateParameter("param11", adNumeric, adParamInput, 10, mvarDA)
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adNumeric, adParamInput, 10, mvarARREAR)
    cmd.Parameters.Append Param12
      
    Set Param13 = cmd.CreateParameter("param13", adNumeric, adParamInput, 10, mvarOTHERS_ALLOWANCE)
    cmd.Parameters.Append Param13
    Set Param14 = cmd.CreateParameter("param14", adNumeric, adParamInput, 10, mvarOTHERS_DEDUCTION)
    cmd.Parameters.Append Param14
    
    Set Param15 = cmd.CreateParameter("param15", adNumeric, adParamInput, 10, mvarR_STAMP)
    cmd.Parameters.Append Param15
    
    Set Param16 = cmd.CreateParameter("param16", adVarChar, adParamInput, 254, mvarRemarks)
    cmd.Parameters.Append Param16
  
    Set Param17 = cmd.CreateParameter("param17", adNumeric, adParamInput, 10, mvarADV_AMOUNT)
    cmd.Parameters.Append Param17

    Set Param18 = cmd.CreateParameter("param18", adNumeric, adParamInput, 10, mvarLN_AMOUNT)
    cmd.Parameters.Append Param18
    
    Set Param19 = cmd.CreateParameter("param19", adNumeric, adParamInput, 10, mvarPF_Ded)
    cmd.Parameters.Append Param19
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Pkg_Salary.Salary_IU(?, ?, ?,?,?, ?, ?,?,?, ?, ?,?,?, ?, ?,?,?, ?, ?)}"
    
    Set Rs = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False


End Function
Public Function Delete()

If mvarEmp_ID = Empty Then Exit Function

    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim Rs As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim Param3 As New ADODB.Parameter
    Dim Param4 As New ADODB.Parameter
     
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 15, mvarSalay_Month_Get)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, mvarSalary_YearGet)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 1, mvarSalaryType)
    cmd.Parameters.Append Param4

'----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
           
   ' cmd.CommandText = "{CALL salary_preparation_Delete(?, ?, ?)}"
    
    If BonusPreparationStatus = 0 Then
        cmd.CommandText = "{CALL salary_preparation_Delete(?, ?, ?,?)}"
    Else
        cmd.CommandText = "{CALL Bonus_preparation_Delete(?, ?, ?,?)}"
    End If
    
    
    
    Set Rs = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function
Public Function GetX() As Recordset
  On Error Resume Next

    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim Rs As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim Param3 As New ADODB.Parameter
    Dim Param4 As New ADODB.Parameter
    
    ''Here Param1 is declared as Parameter type variable for taking input parameter
    '' and Param2 is for retrieving records fetched by the command text.
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adInteger, adParamInput, 2, mvarPAY_MONTH)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adInteger, adParamInput, 4, mvarPAY_YEAR)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adSmallInt, adParamOutput)
    cmd.Parameters.Append Param4
    '----------------------------------------------------------------------------------
    ' Enable PLSQLRSet property
    
   
    
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL PKG_Salary.Salary_Get_X(?,?,?,?)}"

    Set Rs = cmd.Execute

    ' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
    
    Set GetX = Rs
    
    If Not (Rs.EOF And Rs.BOF) Then
        
        mvarEmp_ID = Rs("Emp_Id")
        mvarEmp_Nm = Rs("Emp_Nm")
        mvarPAY_MONTH = Rs("PAY_MONTH")
        mvarPAY_YEAR = Rs("PAY_Year")
        mvarDESIGNATION = Rs("Designation")
        mvarDEPT_NM = Rs("Dept_Nm")
        mvarATTN = Rs("ATTN")
        mvarBASIC = Rs("BASIC")
        mvarH_RENT = Rs("H_Rent")
        mvarMED = Rs("Med")
        mvarCONV = Rs("Conv")
        mvarTFN = Rs("Tfn")
        mvarTELEPHONE = Rs("Telephone")
        mvarDA = Rs("DA")
        mvarARREAR = Rs("ARREAR")
        mvarPF_Ded = Rs("PF_Deduction")
        mvarR_STAMP = Rs("R_Stamp")
        mvarADV_AMOUNT = Rs("ADV_AMOUNT")
        mvarLN_AMOUNT = Rs("LN_AMOUNT")
        mvarRemarks = Rs("REMARKS")
        mvarPF_Ded = Rs("PF_Deduction")
                    
    Else
        
        mvarEmp_ID = ""
        mvarEmp_Nm = ""
        mvarPAY_MONTH = 0
        mvarPAY_YEAR = 0
        mvarDESIGNATION = ""
        mvarDEPT_NM = ""
        mvarATTN = 0
        mvarBASIC = 0
        mvarH_RENT = 0
        mvarMED = 0
        mvarCONV = 0
        mvarTFN = 0
        mvarTELEPHONE = 0
        mvarDA = 0
        mvarARREAR = 0
        mvarPF_Ded = 0
        mvarR_STAMP = 0
        mvarADV_AMOUNT = 0
        mvarLN_AMOUNT = 0
        mvarRemarks = ""
    End If
End Function

Public Sub Show_Message()
       
    Dim conn As New Connection
    Dim cmd As New Command
    Dim Rs As New Recordset
    
    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "Select * from Temp"
    cmd.Properties("IRowsetChange") = True
    cmd.Properties("Updatability") = 7
    
    Rs.CursorLocation = adUseClient
    Rs.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic

    MsgBox Rs.Fields(0), vbInformation + vbOKOnly, "Message"

End Sub
Public Function Disburse()
    
       On Error Resume Next
        Dim conn As New Connection
        Dim cmd As New Command
        Dim Rs As New Recordset
        
        conn.Open mvarConnstring
        Set cmd.ActiveConnection = conn
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "Disburse"
        cmd(1) = mvarDisburse_What
        cmd(2) = mvarDisburse_Method
        cmd(3) = mvarDisburse_To
        cmd(4) = mvarPAY_MONTH
        cmd(5) = mvarPAY_YEAR
        cmd(6) = mvarDisburse_date
        
        Set Rs = cmd.Execute
        
        MsgBox Rs!Message, vbOKOnly + vbExclamation
        
        Exit Function
SaveError:
        
        Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the Salary_Monthly Table" + vbCr + vbCr + Err.Description
End Function
