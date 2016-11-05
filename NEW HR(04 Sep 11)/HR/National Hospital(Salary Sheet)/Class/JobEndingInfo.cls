VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JobEndingInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarEmp_ID As String
Private mvarEmp_Name As String
Private mvarDESIGNATION As String
Private mvarJobEnding_Type As String
Private mvarJobEndingStatus As String
Private mvarDesciption As String
Private mvarConnstring As String
Private mvarDepartment As String
Private mvarJobEndingDate As Date
Public Property Let JobEndingDate(ByVal vData As Date)
    mvarJobEndingDate = vData
End Property
Public Property Get JobEndingDate() As Date
    JobEndingDate = mvarJobEndingDate
End Property
Public Property Let Department(ByVal vData As String)
    mvarDepartment = vData
End Property
Public Property Get Department() As String
    Department = mvarDepartment
End Property
Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property
Public Property Let Desciption(ByVal vData As String)
    mvarDesciption = vData
End Property
Public Property Get Desciption() As String
    Desciption = mvarDesciption
End Property
Public Property Let JobEndingStatus(ByVal vData As String)
    mvarJobEndingStatus = vData
End Property
Public Property Get JobEndingStatus() As String
    JobEndingStatus = mvarJobEndingStatus
End Property
Public Property Let JobEnding_Type(ByVal vData As String)
    mvarJobEnding_Type = vData
End Property
Public Property Get JobEnding_Type() As String
    JobEnding_Type = mvarJobEnding_Type
End Property
Public Property Let DESIGNATION(ByVal vData As String)
    mvarDESIGNATION = vData
End Property
Public Property Get DESIGNATION() As String
    DESIGNATION = mvarDESIGNATION
End Property
Public Property Let Emp_Name(ByVal vData As String)
    mvarEmp_Name = vData
End Property
Public Property Get Emp_Name() As String
    Emp_Name = mvarEmp_Name
End Property
Public Property Let Emp_ID(ByVal vData As String)
    mvarEmp_ID = vData
End Property
Public Property Get Emp_ID() As String
    Emp_ID = mvarEmp_ID
End Property
Public Function Save() As String
'On Error Resume Next
'If mvarCo_Nm = Empty Then Exit Function

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
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 100, mvarEmp_Name)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 100, mvarDESIGNATION)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 100, mvarJobEnding_Type)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 12, Format(mvarJobEndingDate, "DD-MMM-YYYY"))
    cmd.Parameters.Append Param5
       
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 150, mvarDesciption)
    cmd.Parameters.Append Param6
        
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 100, mvarJobEndingStatus)
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 100, mvarDepartment)
    cmd.Parameters.Append Param8
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL Job_Ending_save(?,?,?,?,?,?,?,?)}"
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False

End Function



