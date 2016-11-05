VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "St_JbType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarConnstring As String
Private mvarJType_Code As String
Private mvarJType_Nm As String
Private mvarDescription As String
Private mvarTrack_Id As Long
Public Property Let JType_Code(ByVal vData As String)
   mvarJType_Code = vData
End Property
Public Property Get JType_Code() As String
   JType_Code = mvarJType_Code
End Property
Public Property Let JType_Nm(ByVal vData As String)
   mvarJType_Nm = vData
End Property
Public Property Get JType_Nm() As String
   JType_Nm = mvarJType_Nm
End Property
Public Property Let Description(ByVal vData As String)
   mvarDescription = vData
End Property
Public Property Get Description() As String
   Description = mvarDescription
End Property
Public Property Let Track_Id(ByVal vData As Long)
   mvarTrack_Id = vData
End Property
Public Property Get Track_Id() As Long
   Track_Id = mvarTrack_Id
End Property
Public Property Let Connstring(ByVal vData As String)
   mvarConnstring = vData
End Property
Public Property Get Connstring() As String
   Connstring = mvarConnstring
End Property
Public Function Save()
''On ERROR Resume Next

If mvarJType_Code = Empty Then Exit Function

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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarJType_Code)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 50, mvarJType_Nm)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 100, mvarDescription)
    cmd.Parameters.Append Param3
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Pkg_St_Company.St_JobType_IU(?, ?, ?)}"
    
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
    cmd.CommandText = "Delete from St_JbType where JType_Code='" + mvarJType_Code + "'"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
    Exit Function

End Function
Public Function GetX() As Recordset

'On ERROR Resume Next

If mvarJType_Code = Empty Then Exit Function

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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarJType_Code)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamOutput)
    cmd.Parameters.Append Param2
    '----------------------------------------------------------------------------------
    ' Enable PLSQLRSet property
    
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL PKG_St_Company.St_JobType_GetX(?,?)}"

    Set RS = cmd.Execute

    ' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
    
    Set GetX = RS
    
    If Not (RS.EOF And RS.BOF) Then
       mvarJType_Code = RS("JType_Code")
       mvarJType_Nm = RS("JType_Nm")
       mvarDescription = "" & RS("Description")
    Else
       mvarJType_Code = ""
       mvarJType_Nm = ""
       mvarDescription = ""
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
    cmd.CommandText = "Select * from St_JbType"
    cmd.Properties("IRowsetChange") = True
    cmd.Properties("Updatability") = 7
    
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
   
    If Not (RS.EOF Or RS.BOF) Then
        Set GetAll = RS
        Exit Function
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
