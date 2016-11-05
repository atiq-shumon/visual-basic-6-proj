VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Overtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarEmp_ID As String 'local copy
Private mvarPAYDATE As Date 'local copy
Private mvarYEARFORPAYMENT As String 'local copy
Private mvarMONTHFORPAYMENT As String 'local copy
Private mvarPayScale As String 'local copy
Private mvarOVERTIMEHOURPERDAY As String 'local copy
Private mvarAMOUNT As String 'local copy
Private mvarOTHERSAMOUNT As String 'local copy
Private mvarREVSTAMP As String 'local copy
Private mvarDEDUCTION As String 'local copy
Private mvarNOOFDAYS As String 'local copy
Private mvarPAYMENTTYPR As String 'local copy
Private mvarOTTYPE As String 'local copy
Private mvarConnstring As String 'local copy
Private mvarNETPAYABLE As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarMONEY_TAKEN As String 'local copy
'local variable(s) to hold property value(s)
Private mvarRemarks As String 'local copy
'local variable(s) to hold property value(s)
Private mvarNoofHrsOvertime As String 'local copy
Public Property Let NoofHrsOvertime(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NoofHrsOvertime = 5
    mvarNoofHrsOvertime = vData
End Property


Public Property Get NoofHrsOvertime() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NoofHrsOvertime
    NoofHrsOvertime = mvarNoofHrsOvertime
End Property



Public Property Let Remarks(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Remarks = 5
    mvarRemarks = vData
End Property


Public Property Get Remarks() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Remarks
    Remarks = mvarRemarks
End Property



Public Property Let MONEY_TAKEN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MONEY_TAKEN = 5
    mvarMONEY_TAKEN = vData
End Property


Public Property Get MONEY_TAKEN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MONEY_TAKEN
    MONEY_TAKEN = mvarMONEY_TAKEN
End Property



Public Property Let NETPAYABLE(ByVal vData As Integer)
    mvarNETPAYABLE = vData
End Property
Public Property Get NETPAYABLE() As Integer
    NETPAYABLE = mvarNETPAYABLE
End Property
Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property
Public Property Let OTTYPE(ByVal vData As String)
    mvarOTTYPE = vData
End Property
Public Property Get OTTYPE() As String
    OTTYPE = mvarOTTYPE
End Property
Public Property Let PAYMENTTYPR(ByVal vData As String)
    mvarPAYMENTTYPR = vData
End Property
Public Property Get PAYMENTTYPR() As String
    PAYMENTTYPR = mvarPAYMENTTYPR
End Property
Public Property Let NOOFDAYS(ByVal vData As String)
    mvarNOOFDAYS = vData
End Property
Public Property Get NOOFDAYS() As String
    NOOFDAYS = mvarNOOFDAYS
End Property
Public Property Let DEDUCTION(ByVal vData As String)
    mvarDEDUCTION = vData
End Property
Public Property Get DEDUCTION() As String
    DEDUCTION = mvarDEDUCTION
End Property
Public Property Let REVSTAMP(ByVal vData As String)
    mvarREVSTAMP = vData
End Property
Public Property Get REVSTAMP() As String
    REVSTAMP = mvarREVSTAMP
End Property
Public Property Let OTHERSAMOUNT(ByVal vData As String)
    mvarOTHERSAMOUNT = vData
End Property
Public Property Get OTHERSAMOUNT() As String
    OTHERSAMOUNT = mvarOTHERSAMOUNT
End Property
Public Property Let AMOUNT(ByVal vData As String)
    mvarAMOUNT = vData
End Property
Public Property Get AMOUNT() As String
    AMOUNT = mvarAMOUNT
End Property
Public Property Let OVERTIMEHOURPERDAY(ByVal vData As String)
    mvarOVERTIMEHOURPERDAY = vData
End Property
Public Property Get OVERTIMEHOURPERDAY() As String
    OVERTIMEHOURPERDAY = mvarOVERTIMEHOURPERDAY
End Property
Public Property Let PayScale(ByVal vData As String)
    mvarPayScale = vData
End Property
Public Property Get PayScale() As String
    PayScale = mvarPayScale
End Property
Public Property Let MONTHFORPAYMENT(ByVal vData As String)
    mvarMONTHFORPAYMENT = vData
End Property
Public Property Get MONTHFORPAYMENT() As String
    MONTHFORPAYMENT = mvarMONTHFORPAYMENT
End Property
Public Property Let YEARFORPAYMENT(ByVal vData As String)
    mvarYEARFORPAYMENT = vData
End Property
Public Property Get YEARFORPAYMENT() As String
    YEARFORPAYMENT = mvarYEARFORPAYMENT
End Property
Public Property Let PAYDATE(ByVal vData As Date)
    mvarPAYDATE = vData
End Property
Public Property Get PAYDATE() As Date
    PAYDATE = mvarPAYDATE
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
                     
        
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, mvarPAYDATE)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, mvarYEARFORPAYMENT)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 50, mvarMONTHFORPAYMENT)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 500, mvarPayScale)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 150, mvarOVERTIMEHOURPERDAY)
    cmd.Parameters.Append Param6

    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 10, mvarAMOUNT)
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adDouble, adParamInput, 10, mvarOTHERSAMOUNT)
    cmd.Parameters.Append Param8
    
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 10, mvarREVSTAMP)
    cmd.Parameters.Append Param9
    
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, mvarDEDUCTION)
    cmd.Parameters.Append Param10

    Set Param11 = cmd.CreateParameter("param11", adDouble, adParamInput, 10, mvarNETPAYABLE)
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adDouble, adParamInput, 2, mvarNOOFDAYS)
    cmd.Parameters.Append Param12
    
    Set Param13 = cmd.CreateParameter("param13", adVarChar, adParamInput, 2, mvarPAYMENTTYPR)
    cmd.Parameters.Append Param13
    
    Set Param14 = cmd.CreateParameter("param14", adVarChar, adParamInput, 2, mvarOTTYPE)
    cmd.Parameters.Append Param14
    
    
    Set Param15 = cmd.CreateParameter("param15", adVarChar, adParamInput, 2, 0)
    cmd.Parameters.Append Param15
    
    
    Set Param16 = cmd.CreateParameter("param16", adVarChar, adParamInput, 200, mvarRemarks)
    cmd.Parameters.Append Param16
    '----------------------------------------------------------------------------------
    Set Param17 = cmd.CreateParameter("param17", adDouble, adParamInput, 2, mvarNoofHrsOvertime)
    cmd.Parameters.Append Param17



' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Over_TimePreparation_save(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
End Function
Public Function Delete_Overtime()
    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
            
        
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, mvarPAYDATE)
    cmd.Parameters.Append Param2
    
    '----------------------------------------------------------------------------------




' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Overtime_Preparation_Delete(?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
End Function


