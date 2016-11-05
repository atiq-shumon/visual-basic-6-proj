VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "St_Desig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarConnstring As String
Private mvarDesig_Code As String
Private mvarDESIGNATION As String
Private mvarDesig_Level As Long
Private mvarTrack_Id As Long
Private mvarPool As Integer 'local copy
Private mvarEmp_Type As Integer 'local copy
Private mvarPW_Commission As Integer 'local copy
Public Property Let PW_Commission(ByVal vData As Integer)
    mvarPW_Commission = vData
End Property
Public Property Get PW_Commission() As Integer
     PW_Commission = mvarPW_Commission
End Property
Public Property Let Emp_Type(ByVal vData As Integer)
    mvarEmp_Type = vData
End Property
Public Property Get Emp_Type() As Integer
    Emp_Type = mvarEmp_Type
End Property
Public Property Let Pool(ByVal vData As Integer)
   mvarPool = vData
End Property
Public Property Get Pool() As Integer
    Pool = mvarPool
End Property
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
Public Property Let Desig_Code(ByVal vData As String)
   mvarDesig_Code = vData
End Property
Public Property Get Desig_Code() As String
   Desig_Code = mvarDesig_Code
End Property
Public Property Let DESIGNATION(ByVal vData As String)
   mvarDESIGNATION = vData
End Property
Public Property Get DESIGNATION() As String
   DESIGNATION = mvarDESIGNATION
End Property
Public Property Let Desig_Level(ByVal vData As Long)
   mvarDesig_Level = vData
End Property
Public Property Get Desig_Level() As Long
   Desig_Level = mvarDesig_Level
End Property
Public Property Let Connstring(ByVal vData As String)
   mvarConnstring = vData
End Property
Public Property Get Connstring() As String
   Connstring = mvarConnstring
End Property
Public Function Save()
''On ERROR Resume Next

If mvarDesig_Code = Empty Then Exit Function

    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim Param3 As New ADODB.Parameter
    Dim Param4 As New ADODB.Parameter
    Dim Param5 As New ADODB.Parameter
    Dim Param6 As New ADODB.Parameter
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarDesig_Code)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 50, mvarDESIGNATION)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adInteger, adParamInput, 2, mvarPool)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adInteger, adParamInput, 2, mvarDesig_Level)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 2, mvarEmp_Type)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adInteger, adParamInput, 2, mvarPW_Commission)
    cmd.Parameters.Append Param6
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Pkg_St_Company.St_desig_IU(?, ?, ?, ?, ?, ?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function
Public Function Delete()

    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset
 
    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "Delete from St_Desig where Desig_Code='" + mvarDesig_Code + "'"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
    Exit Function

End Function
Public Function GetX() As Recordset

On Error Resume Next

If mvarDesig_Code = Empty Then Exit Function

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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarDesig_Code)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamOutput)
    cmd.Parameters.Append Param2
    '----------------------------------------------------------------------------------
    ' Enable PLSQLRSet property
    
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL PKG_St_Company.St_Desig_GetX(?,?)}"

    Set RS = cmd.Execute

    ' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
    
    Set GetX = RS
    
    If Not (RS.EOF And RS.BOF) Then
       mvarDesig_Code = RS("Desig_Code")
       mvarDESIGNATION = RS("Designation")
       mvarDesig_Level = RS("Desig_Level")
       mvarEmp_Type = RS("Emp_Type")
       mvarPW_Commission = RS("PW_Commission")
       mvarPool = RS("Pool")
    Else
       mvarDESIGNATION = ""
       mvarDesig_Level = ""
       mvarEmp_Type = ""
       mvarPW_Commission = ""
       mvarPool = ""
          
    End If

End Function
Public Function GetAll() As Recordset
   On Error Resume Next

    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset
    
    
    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn

    cmd.CommandType = adCmdText
    cmd.CommandText = "Select * from St_Desig"
    cmd.Properties("IRowsetChange") = True
    cmd.Properties("Updatability") = 7
    
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
   
    If Not (RS.EOF Or RS.BOF) Then
        Set GetAll = RS
        Exit Function
    End If
    
End Function

