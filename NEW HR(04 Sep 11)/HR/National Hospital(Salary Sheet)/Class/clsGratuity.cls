VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGratuity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sGratuityReceiveId As String
Private sSourceOfFund As String
Private sDescription As String
Private dPaymentReceiveDate As Date
Private nReceiveAmount  As Double
Private sPaymentReceivedType As String
Private sCheckNo As String
Private sAccountNo As String
Private sBankCode As String
Private sVoucherNo As String
Private sAccountType As String
Private mvarConnstring As String


Public Property Let GratuityReceiveId(ByVal vData As String)
    sGratuityReceiveId = vData
End Property
Public Property Get GratuityReceiveId() As String
    GratuityReceiveId = sGratuityReceiveId
End Property
Public Property Let SourceOfFund(ByVal vData As String)
    sSourceOfFund = vData
End Property
Public Property Get SourceOfFund() As String
    SourceOfFund = sSourceOfFund
End Property


Public Property Let Description(ByVal vData As String)
    sDescription = vData
End Property
Public Property Get Description() As String
    Description = sDescription
End Property
Public Property Let PaymentReceiveDate(ByVal vData As Date)
    dPaymentReceiveDate = vData
End Property
Public Property Get PaymentReceiveDate() As Date
    PaymentReceiveDate = dPaymentReceiveDate
End Property
Public Property Let ReceiveAmount(ByVal vData As Double)
    nReceiveAmount = vData
End Property
Public Property Get ReceiveAmount() As Double
    ReceiveAmount = nReceiveAmount
End Property
Public Property Let PaymentReceivedType(ByVal vData As String)
    sPaymentReceivedType = vData
End Property
Public Property Get PaymentReceivedType() As String
    PaymentReceivedType = sPaymentReceivedType
End Property
Public Property Let CheckNo(ByVal vData As String)
    sCheckNo = vData
End Property
Public Property Get CheckNo() As String
    CheckNo = sCheckNo
End Property
Public Property Let AccountNo(ByVal vData As String)
    sAccountNo = vData
End Property
Public Property Get AccountNo() As String
    AccountNo = sAccountNo
End Property
Public Property Let BankCode(ByVal vData As String)
    sBankCode = vData
End Property
Public Property Get BankCode() As String
    BankCode = sBankCode
End Property
Public Property Let VoucherNo(ByVal vData As String)
    sVoucherNo = vData
End Property
Public Property Get VoucherNo() As String
    VoucherNo = sVoucherNo
End Property
Public Property Let AccountType(ByVal vData As String)
    sAccountType = vData
End Property
Public Property Get AccountType() As String
    AccountType = sAccountType
End Property


Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property
Public Function Save() As String
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
       
       
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, sGratuityReceiveId)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 15, sSourceOfFund)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 250, sDescription)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 11, dPaymentReceiveDate)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adNumeric, adParamInput, 10, nReceiveAmount)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 5, sPaymentReceivedType)
    cmd.Parameters.Append Param6
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 50, sCheckNo)
    cmd.Parameters.Append Param7
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 100, sAccountNo)
    cmd.Parameters.Append Param8
    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 15, sBankCode)
    cmd.Parameters.Append Param9
     
    Set Param10 = cmd.CreateParameter("Param10", adVarChar, adParamInput, 50, sVoucherNo)
    cmd.Parameters.Append Param10
     
    Set Param11 = cmd.CreateParameter("Param11", adVarChar, adParamInput, 15, sAccountType)
    cmd.Parameters.Append Param11
     
    '----------------------------------------------------------------------------------
' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL GRATUITY_RECEIVE_SAVE( ?,?,?,?,?,?,?,?,?,?,?)}"
    Set Rs = cmd.Execute
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function

Public Function Delete()
    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim Rs As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
       
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, sGratuityReceiveId)
    cmd.Parameters.Append Param1
    '----------------------------------------------------------------------------------
' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL GRATUITY_RECEIVE_DELETE(?)}"
    Set Rs = cmd.Execute
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function


Public Function Save_Gratuity_Open_Close() As String
    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim Rs As New ADODB.Recordset
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL GRA_OPENING_CLOSING_SAVE()}"
    Set Rs = cmd.Execute
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function

