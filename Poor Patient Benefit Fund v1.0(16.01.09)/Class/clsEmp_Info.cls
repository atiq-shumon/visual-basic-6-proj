VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmp_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarEmp_id As String 'local copy
Private mvarEmp_Name As String 'local copy
Private mvarAddress As String 'local copy
Private mvarConnection_String As String 'local copy
Public Property Let Connection_String(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Connection_String = 5
    mvarConnection_String = vData
End Property
Public Property Get Connection_String() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Connection_String
    Connection_String = mvarConnection_String
End Property
  
'Public Function GetAll() As Recordset
'
'   ' On Error Resume Next
'
'    Dim Conn As New Connection
'    Dim cmd As New Command
'    Dim RS As New Recordset
'
'    Conn.Open mvarConnection_String
'    Set cmd.ActiveConnection = Conn
'
'    cmd.CommandType = adCmdText
'    cmd.CommandText = "Select * from Emp_Info"
'    cmd.Properties("IRowsetChange") = True
'    cmd.Properties("Updatability") = 7
'
'    RS.CursorLocation = adUseClient
'    RS.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
'
'    If Not (RS.EOF Or RS.BOF) Then
'        Set GetAll = RS
'        Exit Function
'    End If
'
'End Function
'
'Public Function Save()
'''On Error Resume Next
'
'If mvarEmp_id = Empty Then Exit Function
'
'    Dim Conn As New ADODB.Connection
'    Dim cmd As New ADODB.Command
'    Dim RS As New ADODB.Recordset
'
'    Dim Param1 As New Parameter
'    Dim Param2 As New Parameter
'    Dim Param3 As New Parameter
'
'    Conn.Open mvarConnection_String
'
'    Set cmd.ActiveConnection = Conn
'    cmd.CommandType = adCmdText
'
'    '----------------------------------------------------------------------------------
'    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_id)
'    cmd.Parameters.Append Param1
'
'    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 50, mvarEmp_Name)
'    cmd.Parameters.Append Param2
'
'    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 100, mvarAddress)
'    cmd.Parameters.Append Param3
'    '----------------------------------------------------------------------------------
'' Enable PLSQLRSet property
'cmd.Properties("PLSQLRSet") = True
'
'cmd.CommandText = "{CALL Emp_Info_IU(?, ?, ?)}"
'Set RS = cmd.Execute
'
'' Disable PLSQLRSet property
'cmd.Properties("PLSQLRSet") = False
'End Function
'Public Function Delete()
'
'    Dim Conn As New Connection
'    Dim cmd As New Command
'    Dim RS As New Recordset
'
'    Conn.Open mvarConnection_String
'    Set cmd.ActiveConnection = Conn
'
'    cmd.CommandType = adCmdText
'    cmd.CommandText = "Delete from Emp_Info where Emp_Id='" + mvarEmp_id + "'"
'    RS.CursorLocation = adUseClient
'    RS.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
'
'    Exit Function
'
'End Function
'Public Function GetX() As Recordset
'
'On Error Resume Next
'
'If mvarEmp_id = Empty Then Exit Function
'
'Dim Conn As New Connection
'    Dim cmd As New Command
'    Dim RS As New Recordset
'
'    Conn.Open mvarConnection_String
'    Set cmd.ActiveConnection = Conn
'
'    cmd.CommandType = adCmdText
'    cmd.CommandText = "Select  from Emp_Info where emp_id= '" + mvarEmp_id + "'"
'    cmd.Properties("IRowsetChange") = True
'    cmd.Properties("Updatability") = 7
'
'    RS.CursorLocation = adUseClient
'    RS.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
'
'    If Not (RS.EOF And RS.BOF) Then
'        mvarEmp_id = RS("Emp_Id")
'        mvarEmp_Name = RS("Emp_Name")
'        mvarAddress = RS("Emp_Address")
'    Else
'        mvarEmp_id = ""
'        mvarEmp_Name = ""
'        mvarAddress = ""
'    End If
'
'
'End Function
'Public Sub Show_Message()
'
'    If mvarEmp_id = "" Then Exit Sub
'
'    Dim Conn As New Connection
'    Dim cmd As New Command
'    Dim RS As New Recordset
'
'    Conn.Open mvarConnection_String
'    Set cmd.ActiveConnection = Conn
'
'    cmd.CommandType = adCmdText
'    cmd.CommandText = "Select * from Temp"
'    cmd.Properties("IRowsetChange") = True
'    cmd.Properties("Updatability") = 7
'
'    RS.CursorLocation = adUseClient
'    RS.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
'
'    MsgBox RS.Fields(0), vbInformation + vbOKOnly, "Message"
'
'End Sub
'
'
'Public Property Let Address(ByVal vData As String)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.Address = 5
'    mvarAddress = vData
'End Property
'
'
'Public Property Get Address() As String
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.Address
'     Address = mvarAddress
'End Property
'
'
'
'Public Property Let Emp_Name(ByVal vData As String)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.Emp_Name = 5
'    mvarEmp_Name = vData
'End Property
'
'
'Public Property Get Emp_Name() As String
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.Emp_Name
'     Emp_Name = mvarEmp_Name
'End Property
'
'
'
'Public Property Let Emp_id(ByVal vData As String)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.Emp_id = 5
'    mvarEmp_id = vData
'End Property
'
'
'Public Property Get Emp_id() As String
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.Emp_id
'     Emp_id = mvarEmp_id
'End Property
'
