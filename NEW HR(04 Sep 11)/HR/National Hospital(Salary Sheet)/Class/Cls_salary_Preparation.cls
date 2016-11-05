VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_salary_Preparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarEmp_ID As String 'local copy
Private mvarPAY_MONTH As String 'local copy
Private mvarPAY_YEAR As String 'local copy
Private mvarDESIGNATION As String 'local copy
Private mvarDEPT_NM As String 'local copy
Private mvarATTN As String 'local copy
Private mvarLEAVE As String 'local copy
Private mvarBASIC As Double 'local copy
Private mvarH_RENT As Double 'local copy
Private mvarMED As Double 'local copy
Private mvarCONV As Double 'local copy
Private mvarTFN As Double 'local copy
Private mvarDA As Double 'local copy
Private mvarSDA As Double 'local copy
Private mvarARREAR As Double 'local copy
Private mvarBonus As Double 'local copy
Private mvarOTHERS_ALLOWANCE As Double 'local copy
Private mvarOTHERS_ADDITION As Double 'local copy
Private mvarPF_CONTRI_DEDUCTION As Double 'local copy
Private mvarPF_LN_AMOUNT As Double 'local copy
Private mvarSALARY_ADVANCE As Double 'local copy
Private mvarR_STAMP As Double 'local copy
Private mvarOTHERS_DEDUCTION As Double 'local copy
Private mvarNET_PAYABLE As Double 'local copy
Private mvarCREATE_DATE As Date 'local copy
Private mvarCREATE_BY As String 'local copy
Private mvarUPDATE_DATE As Date 'local copy
Private mvarEmp_Nm As String 'local copy
Private mvarConnstring As String 'local copy
Private mvarRemarks As String 'local copy
Private mvarEMPLOYER_CONTRIBUTION As Double 'local copy
Private mvarEMPLOYEE_CONTRIBUTION As Double 'local copy
Private mvarCREATED_BY As String 'local copy
Private mvarCREATED_DATE As Date 'local copy
Private mvarWORKING_DAY As String 'local copy
Private mvarSALARY_DISBURSE As String 'local copy
Private mvarAccountHitdate As Date 'local copy
Private mvarNarrationofAcc As String 'local copy
Private mvarAcc_Code As String 'local copy
Private mvarAcc_NetPayment As Double 'local copy
Private mvarAcc_NetPayCr As Double 'local copy
Private mvarAcc_NetPayDr As Double 'local copy
Private mvarAcc_PaymenyCPBD As String 'local copy
Private mvarAcc_Check_No As String 'local copy
Private mvarAcc_UserID As String 'local copy
Private mvarAcc_GetNull As String 'local copy
Private mvarEnfoFPF_BgYear As String 'local copy
Private mvarEnfoFPF_EndYear As String 'local copy
Private mvarVou_No As String 'local copy
Private mvarLoanRefundedDate As Date 'local copy
Private mvarNoOfInstallmentPaid As String 'local copy
Private mvarAmountPaid As Integer 'local copy
Private mvarCurrentBalance As Integer 'local copy
Private mvarNotes As String 'local copy
Private mvarLoanRefundNo As String 'local copy
Private mvarEntrDate As String 'local copy
Private mvaradddeduct_other As Double 'local copy
Private mvarSalaryType As String 'local copy
Private mvarSalaryBonusBasic As String 'local copy
Private mvarDressAllowance As Double 'local copy
Private mvarNDFund As Double 'local copy
Private mvarEducationAsstAllowance As Double
Private mvarEmpPositionSerial As Integer
Private mvarEmpDesignationLevel As Integer
Public Property Let EmployeePositionSerial(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.adddeduct_other = 5
    mvarEmpPositionSerial = vData
End Property
Public Property Get EmployeePositionSerial() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.adddeduct_other
    EmployeePositionSerial = mvarEmpPositionSerial
End Property
Public Property Let EmployeeDesignationLevel(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.adddeduct_other = 5
    mvarEmpDesignationLevel = vData
End Property
Public Property Get EmployeeDesignationLevel() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.adddeduct_other
    EmployeeDesignationLevel = mvarEmpDesignationLevel
End Property
Public Property Let DressAllowance(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.adddeduct_other = 5
    mvarDressAllowance = vData
End Property
Public Property Get DressAllowance() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.adddeduct_other
    DressAllowance = mvarDressAllowance
End Property
Public Property Let EducationAsstAllowance(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.adddeduct_other = 5
    mvarEducationAsstAllowance = vData
End Property
Public Property Get EducationAsstAllowance() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.adddeduct_other
    EducationAsstAllowance = mvarEducationAsstAllowance
End Property

Public Property Let NDFundDeduct(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.adddeduct_other = 5
    mvarNDFund = vData
End Property
Public Property Get NDFundDeduct() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.adddeduct_other
    NDFundDeduct = mvarNDFund
End Property
Public Property Let adddeduct_other(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.adddeduct_other = 5
    mvaradddeduct_other = vData
End Property


Public Property Get adddeduct_other() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.adddeduct_other
    adddeduct_other = mvaradddeduct_other
End Property


Public Property Let EntrDate(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EntrDate = 5
    mvarEntrDate = vData
End Property
Public Property Get EntrDate() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EntrDate
    EntrDate = mvarEntrDate
End Property



Public Property Let LoanRefundNo(ByVal vData As String)
    mvarLoanRefundNo = vData
End Property
Public Property Get LoanRefundNo() As String
    LoanRefundNo = mvarLoanRefundNo
End Property
Public Property Let Notes(ByVal vData As String)
    mvarNotes = vData
End Property
Public Property Get Notes() As String
    Notes = mvarNotes
End Property
Public Property Let CurrentBalance(ByVal vData As Integer)
    mvarCurrentBalance = vData
End Property
Public Property Get CurrentBalance() As Integer
    CurrentBalance = mvarCurrentBalance
End Property
Public Property Let AmountPaid(ByVal vData As Integer)
    mvarAmountPaid = vData
End Property
Public Property Get AmountPaid() As Integer
    AmountPaid = mvarAmountPaid
End Property
Public Property Let SalaryType(ByVal vData As String)
    mvarSalaryType = vData
End Property
Public Property Get SalaryType() As String
    SalaryType = mvarSalaryType
End Property
Public Property Let NoOfInstallmentPaid(ByVal vData As String)
    mvarNoOfInstallmentPaid = vData
End Property
Public Property Get NoOfInstallmentPaid() As String
    NoOfInstallmentPaid = mvarNoOfInstallmentPaid
End Property
Public Property Let LoanRefundedDate(ByVal vData As Date)
    mvarLoanRefundedDate = vData
End Property
Public Property Get LoanRefundedDate() As Date
    LoanRefundedDate = mvarLoanRefundedDate
End Property
Public Property Let Vou_No(ByVal vData As String)
    mvarVou_No = vData
End Property
Public Property Get Vou_No() As String
    Vou_No = mvarVou_No
End Property
Public Property Let EnfoFPF_EndYear(ByVal vData As String)
    mvarEnfoFPF_EndYear = vData
End Property
Public Property Get EnfoFPF_EndYear() As String
    EnfoFPF_EndYear = mvarEnfoFPF_EndYear
End Property
Public Property Let EnfoFPF_BgYear(ByVal vData As String)
    mvarEnfoFPF_BgYear = vData
End Property
Public Property Get EnfoFPF_BgYear() As String
    EnfoFPF_BgYear = mvarEnfoFPF_BgYear
End Property
Public Property Let Acc_GetNull(ByVal vData As String)
   mvarAcc_GetNull = vData
End Property
Public Property Get Acc_GetNull() As String
    Acc_GetNull = mvarAcc_GetNull
End Property
Public Property Let Acc_UserID(ByVal vData As String)
    mvarAcc_UserID = vData
End Property
Public Property Get Acc_UserID() As String
    Acc_UserID = mvarAcc_UserID
End Property
Public Property Let Acc_Check_No(ByVal vData As String)
    mvarAcc_Check_No = vData
End Property
Public Property Get Acc_Check_No() As String
    Acc_Check_No = mvarAcc_Check_No
End Property
Public Property Let Acc_PaymenyCPBD(ByVal vData As String)
    mvarAcc_PaymenyCPBD = vData
End Property
Public Property Get Acc_PaymenyCPBD() As String
    Acc_PaymenyCPBD = mvarAcc_PaymenyCPBD
End Property
Public Property Let Acc_NetPayDr(ByVal vData As Double)
    mvarAcc_NetPayDr = vData
End Property
Public Property Get Acc_NetPayDr() As Double
    Acc_NetPayDr = mvarAcc_NetPayDr
End Property
Public Property Let Acc_NetPayCr(ByVal vData As Double)
    mvarAcc_NetPayCr = vData
End Property
Public Property Get Acc_NetPayCr() As Double
    Acc_NetPayCr = mvarAcc_NetPayCr
End Property
Public Property Let Acc_NetPayment(ByVal vData As Double)
    mvarAcc_NetPayment = vData
End Property
Public Property Get SalaryBonusBasic() As Double
    SalaryBonusBasic = mvarSalaryBonusBasic
End Property
Public Property Let SalaryBonusBasic(ByVal vData As Double)
    mvarSalaryBonusBasic = vData
End Property

Public Property Get Acc_NetPayment() As Double
    Acc_NetPayment = mvarAcc_NetPayment
End Property
Public Property Let Acc_Code(ByVal vData As String)
    mvarAcc_Code = vData
End Property
Public Property Get Acc_Code() As String
    Acc_Code = mvarAcc_Code
End Property
Public Property Let NarrationofAcc(ByVal vData As String)
    mvarNarrationofAcc = vData
End Property
Public Property Get NarrationofAcc() As String
    NarrationofAcc = mvarNarrationofAcc
End Property
Public Property Let AccountHitdate(ByVal vData As Date)
    mvarAccountHitdate = vData
End Property
Public Property Get AccountHitdate() As Date
    AccountHitdate = mvarAccountHitdate
End Property
Public Property Let SALARY_DISBURSE(ByVal vData As String)
    mvarSALARY_DISBURSE = vData
End Property
Public Property Get SALARY_DISBURSE() As String
    SALARY_DISBURSE = mvarSALARY_DISBURSE
End Property
Public Property Let WORKING_DAY(ByVal vData As String)
    mvarWORKING_DAY = vData
End Property
Public Property Get WORKING_DAY() As String
    WORKING_DAY = mvarWORKING_DAY
End Property
Public Property Let CREATED_DATE(ByVal vData As Date)
    mvarCREATED_DATE = vData
End Property
Public Property Get CREATED_DATE() As Date
    CREATED_DATE = mvarCREATED_DATE
End Property
Public Property Let CREATED_BY(ByVal vData As String)
    mvarCREATED_BY = vData
End Property
Public Property Get CREATED_BY() As String
    CREATED_BY = mvarCREATED_BY
End Property
Public Property Let Employee_Contribution(ByVal vData As Double)
    mvarEMPLOYEE_CONTRIBUTION = vData
End Property
Public Property Get Employee_Contribution() As Double
    Employee_Contribution = mvarEMPLOYEE_CONTRIBUTION
End Property
Public Property Let EMPLOYER_CONTRIBUTION(ByVal vData As Double)
    mvarEMPLOYER_CONTRIBUTION = vData
End Property
Public Property Get EMPLOYER_CONTRIBUTION() As Double
    EMPLOYER_CONTRIBUTION = mvarEMPLOYER_CONTRIBUTION
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
Public Property Let Emp_Nm(ByVal vData As String)
    mvarEmp_Nm = vData
End Property
Public Property Get Emp_Nm() As String
    Emp_Nm = mvarEmp_Nm
End Property
Public Property Let UPDATE_DATE(ByVal vData As Date)
    mvarUPDATE_DATE = vData
End Property
Public Property Get UPDATE_DATE() As Date
    UPDATE_DATE = mvarUPDATE_DATE
End Property
Public Property Let CREATE_BY(ByVal vData As String)
    mvarCREATE_BY = vData
End Property
Public Property Get CREATE_BY() As String
    CREATE_BY = mvarCREATE_BY
End Property
Public Property Let CREATE_DATE(ByVal vData As Date)
    mvarCREATE_DATE = vData
End Property
Public Property Get CREATE_DATE() As Date
    CREATE_DATE = mvarCREATE_DATE
End Property
Public Property Let NET_PAYABLE(ByVal vData As Double)
    mvarNET_PAYABLE = vData
End Property
Public Property Get NET_PAYABLE() As Double
    NET_PAYABLE = mvarNET_PAYABLE
End Property
Public Property Let OTHERS_DEDUCTION(ByVal vData As Double)
    mvarOTHERS_DEDUCTION = vData
End Property
Public Property Get OTHERS_DEDUCTION() As Double
    OTHERS_DEDUCTION = mvarOTHERS_DEDUCTION
End Property
Public Property Let R_STAMP(ByVal vData As Double)
    mvarR_STAMP = vData
End Property
Public Property Get R_STAMP() As Double
    R_STAMP = mvarR_STAMP
End Property
Public Property Let SALARY_ADVANCE(ByVal vData As Double)
    mvarSALARY_ADVANCE = vData
End Property
Public Property Get SALARY_ADVANCE() As Double
    SALARY_ADVANCE = mvarSALARY_ADVANCE
End Property
Public Property Let PF_LN_AMOUNT(ByVal vData As Double)
    mvarPF_LN_AMOUNT = vData
End Property
Public Property Get PF_LN_AMOUNT() As Double
    PF_LN_AMOUNT = mvarPF_LN_AMOUNT
End Property
Public Property Let PF_CONTRI_DEDUCTION(ByVal vData As Double)
    mvarPF_CONTRI_DEDUCTION = vData
End Property
Public Property Get PF_CONTRI_DEDUCTION() As Double
    PF_CONTRI_DEDUCTION = mvarPF_CONTRI_DEDUCTION
End Property
Public Property Let OTHERS_ADDITION(ByVal vData As Double)
    mvarOTHERS_ADDITION = vData
End Property
Public Property Get OTHERS_ADDITION() As Double
    OTHERS_ADDITION = mvarOTHERS_ADDITION
End Property
Public Property Let OTHERS_ALLOWANCE(ByVal vData As Double)
    mvarOTHERS_ALLOWANCE = vData
End Property
Public Property Get OTHERS_ALLOWANCE() As Double
    OTHERS_ALLOWANCE = mvarOTHERS_ALLOWANCE
End Property
Public Property Let Bonus(ByVal vData As Double)
    mvarBonus = vData
End Property
Public Property Get Bonus() As Double
    Bonus = mvarBonus
End Property
Public Property Let ARREAR(ByVal vData As Double)
    mvarARREAR = vData
End Property
Public Property Get ARREAR() As Double
    ARREAR = mvarARREAR
End Property
Public Property Let DA(ByVal vData As Double)
    mvarDA = vData
End Property
Public Property Get DA() As Double
    DA = mvarDA
End Property
Public Property Let SDA(ByVal vData As Double)
    mvarSDA = vData
End Property
Public Property Get SDA() As Double
    SDA = mvarSDA
End Property
Public Property Let TFN(ByVal vData As Double)
   mvarTFN = vData
End Property
Public Property Get TFN() As Double
    TFN = mvarTFN
End Property
Public Property Let CONV(ByVal vData As Double)
    mvarCONV = vData
End Property
Public Property Get CONV() As Double
    CONV = mvarCONV
End Property
Public Property Let MED(ByVal vData As Double)
    mvarMED = vData
End Property
Public Property Get MED() As Double
    MED = mvarMED
End Property
Public Property Let H_RENT(ByVal vData As Double)
    mvarH_RENT = vData
End Property
Public Property Get H_RENT() As Double
    H_RENT = mvarH_RENT
End Property
Public Property Let BASIC(ByVal vData As Double)
    mvarBASIC = vData
End Property
Public Property Get BASIC() As Double
    BASIC = mvarBASIC
End Property
Public Property Let LEAVE(ByVal vData As String)
    mvarLEAVE = vData
End Property
Public Property Get LEAVE() As String
    LEAVE = mvarLEAVE
End Property
Public Property Let ATTN(ByVal vData As String)
    mvarATTN = vData
End Property
Public Property Get ATTN() As String
    ATTN = mvarATTN
End Property
Public Property Let DEPT_NM(ByVal vData As String)
    mvarDEPT_NM = vData
End Property
Public Property Get DEPT_NM() As String
    DEPT_NM = mvarDEPT_NM
End Property
Public Property Let designation(ByVal vData As String)
    mvarDESIGNATION = vData
End Property
Public Property Get designation() As String
    designation = mvarDESIGNATION
End Property
Public Property Let PAY_YEAR(ByVal vData As String)
   mvarPAY_YEAR = vData
End Property
Public Property Get PAY_YEAR() As String
    PAY_YEAR = mvarPAY_YEAR
End Property
Public Property Let PAY_MONTH(ByVal vData As String)
    mvarPAY_MONTH = vData
End Property
Public Property Get PAY_MONTH() As String
    PAY_MONTH = mvarPAY_MONTH
End Property
Public Property Let Emp_ID(ByVal vData As String)
    mvarEmp_ID = vData
End Property
Public Property Get Emp_ID() As String
    Emp_ID = mvarEmp_ID
End Property
Public Function Save()
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
    Dim Param23 As New ADODB.Parameter
      
    Dim Param24 As New ADODB.Parameter
    Dim Param25 As New ADODB.Parameter
    Dim Param26 As New ADODB.Parameter
    Dim Param27 As New ADODB.Parameter
    Dim Param28 As New ADODB.Parameter
    Dim Param29 As New ADODB.Parameter
    Dim Param30 As New ADODB.Parameter
    Dim Param31 As New ADODB.Parameter
    Dim Param32 As New ADODB.Parameter
    Dim Param33 As New ADODB.Parameter
    Dim Param34 As New ADODB.Parameter
    Dim Param35 As New ADODB.Parameter
    Dim Param36 As New ADODB.Parameter
    Dim Param37 As New ADODB.Parameter
    Dim Param38 As New ADODB.Parameter
    Dim Param39 As New ADODB.Parameter
    
    
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 100, mvarEmp_Nm)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 20, mvarPAY_MONTH)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 6, mvarPAY_YEAR)
    cmd.Parameters.Append Param4

    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 50, mvarDESIGNATION)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 50, mvarDEPT_NM)
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 10, mvarATTN)
    cmd.Parameters.Append Param7
'---------------------------------------------------
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 10, mvarLEAVE)
    cmd.Parameters.Append Param8
    
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 10, mvarBASIC)
    cmd.Parameters.Append Param9
    
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, mvarH_RENT)
    cmd.Parameters.Append Param10

    Set Param11 = cmd.CreateParameter("param11", adDouble, adParamInput, 10, mvarMED)
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adDouble, adParamInput, 10, mvarCONV)
    cmd.Parameters.Append Param12
    
    Set Param13 = cmd.CreateParameter("param13", adDouble, adParamInput, 10, mvarTFN)
    cmd.Parameters.Append Param13
    
    Set Param14 = cmd.CreateParameter("param14", adDouble, adParamInput, 10, mvarDA)
    cmd.Parameters.Append Param14
    
    Set Param15 = cmd.CreateParameter("param15", adDouble, adParamInput, 10, mvarARREAR)
    cmd.Parameters.Append Param15
    
    Set Param16 = cmd.CreateParameter("param16", adDouble, adParamInput, 10, mvarBonus)
    cmd.Parameters.Append Param16
    
    Set Param17 = cmd.CreateParameter("param17", adDouble, adParamInput, 10, mvarOTHERS_ALLOWANCE)
    cmd.Parameters.Append Param17
    
    Set Param18 = cmd.CreateParameter("param18", adDouble, adParamInput, 10, mvarOTHERS_ADDITION)
    cmd.Parameters.Append Param18
    
        
    Set Param19 = cmd.CreateParameter("param19", adDouble, adParamInput, 10, mvarPF_CONTRI_DEDUCTION)
    cmd.Parameters.Append Param19
    
    Set Param20 = cmd.CreateParameter("param20", adDouble, adParamInput, 10, mvarPF_LN_AMOUNT)
    cmd.Parameters.Append Param20
    
    Set Param21 = cmd.CreateParameter("param21", adDouble, adParamInput, 10, mvarSALARY_ADVANCE)
    cmd.Parameters.Append Param21

    Set Param22 = cmd.CreateParameter("param22", adDouble, adParamInput, 10, mvarR_STAMP)
    cmd.Parameters.Append Param22
    
    Set Param23 = cmd.CreateParameter("param23", adDouble, adParamInput, 10, mvarOTHERS_DEDUCTION)
    cmd.Parameters.Append Param23
    
    Set Param24 = cmd.CreateParameter("param24", adDouble, adParamInput, 10, mvarNET_PAYABLE)
    cmd.Parameters.Append Param24
    
    Set Param25 = cmd.CreateParameter("param25", adDate, adParamInput, 10, mvarCREATE_DATE)
    cmd.Parameters.Append Param25
    
    Set Param26 = cmd.CreateParameter("param26", adVarChar, adParamInput, 11, mvarCREATE_BY)
    cmd.Parameters.Append Param26
    
    Set Param27 = cmd.CreateParameter("param27", adDate, adParamInput, 10, mvarUPDATE_DATE)
    cmd.Parameters.Append Param27
    
    Set Param28 = cmd.CreateParameter("param28", adVarChar, adParamInput, 150, mvarRemarks)
    cmd.Parameters.Append Param28
      
    Set Param29 = cmd.CreateParameter("param29", adVarChar, adParamInput, 150, mvarWORKING_DAY)
    cmd.Parameters.Append Param29
    
    Set Param30 = cmd.CreateParameter("param30", adVarChar, adParamInput, 2, mvarSALARY_DISBURSE)
    cmd.Parameters.Append Param30
    
    Set Param31 = cmd.CreateParameter("param31", adDouble, adParamInput, 10, mvaradddeduct_other)
    cmd.Parameters.Append Param31
    
     Set Param32 = cmd.CreateParameter("param32", adDouble, adParamInput, 10, mvarSDA)
    cmd.Parameters.Append Param32
       
    Set Param33 = cmd.CreateParameter("param33", adChar, adParamInput, 1, mvarSalaryType)
    cmd.Parameters.Append Param33
   
    Set Param34 = cmd.CreateParameter("param34", adDouble, adParamInput, 10, mvarSalaryBonusBasic)
    cmd.Parameters.Append Param34
   
    Set Param35 = cmd.CreateParameter("param35", adDouble, adParamInput, 10, mvarDressAllowance)
    cmd.Parameters.Append Param35
    
    Set Param36 = cmd.CreateParameter("param36", adDouble, adParamInput, 10, mvarNDFund)
    cmd.Parameters.Append Param36
    
    Set Param37 = cmd.CreateParameter("param37", adDouble, adParamInput, 10, mvarEducationAsstAllowance)
    cmd.Parameters.Append Param37
   
    Set Param38 = cmd.CreateParameter("param38", adInteger, adParamInput, 10, mvarEmpPositionSerial)
    cmd.Parameters.Append Param38
    
    Set Param39 = cmd.CreateParameter("param39", adInteger, adParamInput, 10, mvarEmpDesignationLevel)
    cmd.Parameters.Append Param39
   
  
   
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Salary_Preparation_Save(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
      
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function

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
Public Function PF_Save()

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
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, mvarPAY_MONTH)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 4, mvarPAY_YEAR)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, mvarEMPLOYER_CONTRIBUTION)
    cmd.Parameters.Append Param4

    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 10, mvarEMPLOYEE_CONTRIBUTION)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 11, mvarCREATE_BY)
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adDate, adParamInput, 10, mvarUPDATE_DATE)
    cmd.Parameters.Append Param7
      
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL ProVident_Save(?,?,?,?,?,?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function
Public Function Salary_Update_Save()

    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim Param3 As New ADODB.Parameter
    
    
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, mvarPAY_MONTH)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 20, mvarPAY_YEAR)
    cmd.Parameters.Append Param3
    
   
   
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Salary_Preparation_TableUpdate(?,?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function
Public Function Accounts_Hit_Save()
On Error GoTo Errdes
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
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarVou_No)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, mvarAccountHitdate)
    cmd.Parameters.Append Param2

    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 150, mvarNarrationofAcc)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 20, mvarAcc_Code)
    cmd.Parameters.Append Param4
    
   
    Set Param5 = cmd.CreateParameter("Param5", adDouble, adParamInput, 10, mvarAcc_NetPayDr)
    cmd.Parameters.Append Param5

    Set Param6 = cmd.CreateParameter("Param6", adDouble, adParamInput, 20, mvarAcc_NetPayCr)
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("Param7", adVarChar, adParamInput, 20, mvarAcc_PaymenyCPBD)
    cmd.Parameters.Append Param7
   
    
    Set Param8 = cmd.CreateParameter("Param8", adVarChar, adParamInput, 10, mvarAcc_Check_No)
    cmd.Parameters.Append Param8

    Set Param9 = cmd.CreateParameter("Param9", adVarChar, adParamInput, 20, mvarAcc_GetNull)
    cmd.Parameters.Append Param9
    
    Set Param10 = cmd.CreateParameter("Param10", adVarChar, adParamInput, 20, 0)
    cmd.Parameters.Append Param10
    
    Set Param11 = cmd.CreateParameter("Param11", adVarChar, adParamInput, 20, mvarAcc_UserID)
    cmd.Parameters.Append Param11
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Account.Save_Vou_Payroll(?,?,?,?,?,?,?,?,?,?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
Exit Function
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH."

End Function
Public Function PF_EndofYear_Save()
On Error GoTo Errdes
    Dim conn1 As New ADODB.Connection
    Dim cmd1 As New ADODB.Command
    Dim rs1 As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim Param3 As New ADODB.Parameter
    conn1.Open mvarConnstring
    
    Set cmd1.ActiveConnection = conn1
    cmd1.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
   If Len(Trim(Emp_ID_Value)) = 0 Then
        Set Param1 = cmd1.CreateParameter("param1", adInteger, adParamInput, 10, 1)
        cmd1.Parameters.Append Param1
        
        Set Param2 = cmd1.CreateParameter("param2", adVarChar, adParamInput, 10, Emp_ID_Value)
        cmd1.Parameters.Append Param2
    Else
        Set Param1 = cmd1.CreateParameter("param1", adInteger, adParamInput, 10, 0)
        cmd1.Parameters.Append Param1
        
        Set Param2 = cmd1.CreateParameter("param2", adVarChar, adParamInput, 10, Emp_ID_Value)
        cmd1.Parameters.Append Param2
    
    
    End If
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd1.Properties("PLSQLRSet") = True
      
   cmd1.CommandText = "{CALL PF_Bal_EndofYear_Save1(?,?)}"
    
    
    Set rs1 = cmd1.Execute
    
' Disable PLSQLRSet property
cmd1.Properties("PLSQLRSet") = False
Exit Function
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH."
End Function
Public Function Post_to_VOUCHER_Save()
On Error GoTo Errdes
    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarVou_No)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("Param2", adVarChar, adParamInput, 20, mvarAcc_PaymenyCPBD)
    cmd.Parameters.Append Param2
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Account.PostVou(?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
Exit Function
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH."

End Function


Public Function Loan_Sub_Save() As String
'If mvarLeave_Code = Empty Then Exit Function

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
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, mvarLoanRefundedDate)
    cmd.Parameters.Append Param2

    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, mvarNoOfInstallmentPaid)
    cmd.Parameters.Append Param3

    Set Param4 = cmd.CreateParameter("param4", adInteger, adParamInput, 10, mvarAmountPaid)
    cmd.Parameters.Append Param4

    Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 5, mvarCurrentBalance)
    cmd.Parameters.Append Param5

    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 250, mvarNotes)
    cmd.Parameters.Append Param6


    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 10, mvarLoanRefundNo)
    cmd.Parameters.Append Param7
    

    Set Param8 = cmd.CreateParameter("param8", adDate, adParamInput, 15, mvarEntrDate)
    cmd.Parameters.Append Param8
    


' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL LoanInformation_Sub_save( ?,?,?,?,?,?,?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function
