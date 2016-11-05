VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMemberFund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sTrackId As String
Private sAccountNo As String
Private sAccountType As String
Private sBankCode As String
Private nAmount  As Double
Private mvarConnstring As String

Public Property Let TrackId(ByVal vData As String)
    sTrackId = vData
End Property
Public Property Get TrackId() As String
    TrackId = sTrackId
End Property
Public Property Let BankCode(ByVal vData As String)
    sBankCode = vData
End Property
Public Property Get BankCode() As String
    BankCode = sBankCode
End Property
Public Property Let AccountNo(ByVal vData As String)
    sAccountNo = vData
End Property
Public Property Get AccountNo() As String
    AccountNo = sAccountNo
End Property
Public Property Let Amount(ByVal vData As Double)
    nAmount = vData
End Property
Public Property Get Amount() As Double
    Amount = nAmount
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
      
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, sTrackId)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("Param2", adVarChar, adParamInput, 15, sAccountType)
    cmd.Parameters.Append Param2

    Set Param3 = cmd.CreateParameter("Param3", adVarChar, adParamInput, 150, sAccountNo)
    cmd.Parameters.Append Param3

    Set Param4 = cmd.CreateParameter("Param4", adVarChar, adParamInput, 15, sBankCode)
    cmd.Parameters.Append Param4

    Set Param5 = cmd.CreateParameter("param5", adNumeric, adParamInput, 10, nAmount)
    cmd.Parameters.Append Param5
             
    '----------------------------------------------------------------------------------
' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL MEMBER_FUND_SAVE( ?,?,?,?,?)}"
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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, sTrackId)
    cmd.Parameters.Append Param1
    '----------------------------------------------------------------------------------
' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL MEMBER_FUND_DELETE(?)}"
    Set Rs = cmd.Execute
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function




