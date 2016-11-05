VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvaruserID As String 'local copy
'local variable(s) to hold property value(s)
Private mvarUserName As String 'local copy
Private mvarUserPassword As String 'local copy
Private mvarUserDepartment As String 'local copy
Private mvarUserActiveOrNot As Integer 'local copy
Private mvarUserAccessLevel As Integer 'local copy
Private conn As New ADODB.Connection
Private cmd As New ADODB.Command
Private rs As New ADODB.Recordset
Private conString As New CLSCONNECTION



Public Property Let UserAccessLevel(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserAccessLevel = 5
    mvarUserAccessLevel = vData
End Property


Public Property Get UserAccessLevel() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserAccessLevel
    UserAccessLevel = mvarUserAccessLevel
End Property



Public Property Let UserActiveOrNot(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserActiveOrNot = 5
    mvarUserActiveOrNot = vData
End Property


Public Property Get UserActiveOrNot() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserActiveOrNot
    UserActiveOrNot = mvarUserActiveOrNot
End Property



Public Property Let UserDepartment(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserDepartment = 5
    mvarUserDepartment = vData
End Property


Public Property Get UserDepartment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserDepartment
    UserDepartment = mvarUserDepartment
End Property



Public Property Let UserPassword(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserPassword = 5
    mvarUserPassword = vData
End Property


Public Property Get UserPassword() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserPassword
    UserPassword = mvarUserPassword
End Property



Public Property Let UserName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserName = 5
    mvarUserName = vData
End Property


Public Property Get UserName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserName
    UserName = mvarUserName
End Property



Public Property Let UserID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.userID = 5
    mvaruserID = vData
End Property
Public Property Get UserID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.userID
    UserID = mvaruserID
End Property
Public Function Save(security As clsSecurity) As String
    On Error GoTo errDesc
    Dim param0 As New ADODB.Parameter
    Dim param1 As New ADODB.Parameter
    Dim param2 As New ADODB.Parameter
    Dim param3 As New ADODB.Parameter
    Dim param4 As New ADODB.Parameter
    Dim param5 As New ADODB.Parameter
    Dim param6 As New ADODB.Parameter
         
         
    
    Set param0 = cmd.CreateParameter("param0", adInteger, adParamInput, 10, 1)
    cmd.Parameters.Append param0
    
    Set param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, security.UserID)
    cmd.Parameters.Append param1
    
    Set param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 100, security.UserName)
    cmd.Parameters.Append param2
    Set param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 30, security.UserPassword)
    cmd.Parameters.Append param3
    
    Set param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 15, security.UserActiveOrNot)
    cmd.Parameters.Append param4
    
    Set param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 15, security.UserDepartment)
    cmd.Parameters.Append param5
    
    Set param6 = cmd.CreateParameter("param6", adInteger, adParamInput, 3, security.UserAccessLevel)
    cmd.Parameters.Append param6
    
    conn.Open conString.ConnectionString
    cmd.ActiveConnection = conn
    
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "{Call s_u_d_security(?,?,?,?,?,?,?)}"
    
    cmd.Execute
    Save = "User Created Successfully"
    conn.Close
    Set cmd = Nothing
Exit Function
errDesc:
    Save = Err.Description
 


Public Function GetDepartments() As Collection
     Dim department(2) As String
     department(0) = "Acc"
     department(1) = "Per"
End Function



