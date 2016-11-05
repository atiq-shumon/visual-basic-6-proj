VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoanRefundInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarEmp_ID As String 'local copy
Private mvarLoan_Id As String 'local copy
Private mvarLaonIssedDate As Date 'local copy
Private mvarNoOfInstallment As Double 'local copy
Private mvarIssuedAmount As Double 'local copy
Private mvarSlabInstallmentAmount As Double 'local copy
Private mvarNotes As String 'local copy
Private mvarLoanRefundedDate As Date 'local copy
Private mvarNoOfInstallmentPaid As Double 'local copy
Private mvarAmountPaid As Double 'local copy
Private mvarCurrentBalance As Double 'local copy
Private mvarLoanRefundNo As String 'local copy
Private mvarConnstring As String 'local copy
'local variable(s) to hold property value(s)
Private mvarEntrDate As String 'local copy
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



Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property
Public Property Let LoanRefundNo(ByVal vData As String)
    mvarLoanRefundNo = vData
End Property
Public Property Get LoanRefundNo() As String
    LoanRefundNo = mvarLoanRefundNo
End Property
Public Property Let CurrentBalance(ByVal vData As Double)
    mvarCurrentBalance = vData
End Property
Public Property Get CurrentBalance() As Double
    CurrentBalance = mvarCurrentBalance
End Property
Public Property Let AmountPaid(ByVal vData As Double)
    mvarAmountPaid = vData
End Property
Public Property Get AmountPaid() As Double
    AmountPaid = mvarAmountPaid
End Property
Public Property Let NoOfInstallmentPaid(ByVal vData As Double)
    mvarNoOfInstallmentPaid = vData
End Property
Public Property Get NoOfInstallmentPaid() As Double
    NoOfInstallmentPaid = mvarNoOfInstallmentPaid
End Property
Public Property Let LoanRefundedDate(ByVal vData As Date)
    mvarLoanRefundedDate = vData
End Property
Public Property Get LoanRefundedDate() As Date
    LoanRefundedDate = mvarLoanRefundedDate
End Property
Public Property Let Notes(ByVal vData As String)
    mvarNotes = vData
End Property
Public Property Get Notes() As String
    Notes = mvarNotes
End Property
Public Property Let SlabInstallmentAmount(ByVal vData As Double)
    mvarSlabInstallmentAmount = vData
End Property
Public Property Get SlabInstallmentAmount() As Double
    SlabInstallmentAmount = mvarSlabInstallmentAmount
End Property
Public Property Let IssuedAmount(ByVal vData As Double)
    mvarIssuedAmount = vData
End Property
Public Property Get IssuedAmount() As Double
    IssuedAmount = mvarIssuedAmount
End Property
Public Property Let NoOfInstallment(ByVal vData As Double)
    mvarNoOfInstallment = vData
End Property
Public Property Get NoOfInstallment() As Double
    NoOfInstallment = mvarNoOfInstallment
End Property
Public Property Let LaonIssedDate(ByVal vData As Date)
    mvarLaonIssedDate = vData
End Property
Public Property Get LaonIssedDate() As Date
    LaonIssedDate = mvarLaonIssedDate
End Property
Public Property Let Loan_Id(ByVal vData As String)
    mvarLoan_Id = vData
End Property
Public Property Get Loan_Id() As String
    Loan_Id = mvarLoan_Id
End Property
Public Property Let Emp_ID(ByVal vData As String)
    mvarEmp_ID = vData
End Property
Public Property Get Emp_ID() As String
    Emp_ID = mvarEmp_ID
End Property
Public Function Save() As String
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
        
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, mvarLoan_Id)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 15, mvarLaonIssedDate)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adInteger, adParamInput, 10, mvarNoOfInstallment)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 5, mvarIssuedAmount)
    cmd.Parameters.Append Param5
    
    
    Set Param6 = cmd.CreateParameter("param6", adInteger, adParamInput, 10, mvarSlabInstallmentAmount)
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 350, mvarNotes)
    cmd.Parameters.Append Param7
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL LoanInformation_Main_save( ?,?,?,?,?,?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
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


