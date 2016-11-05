VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Company_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarConnstring As String
Private mvarCo_Nm As String
Private mvarCo_Type As String
Private mvarAddress As String
Private mvarPhone As String
Private mvarFax As String
Private mvarE_mail As String
Private mvarNotes As String
Private mvarLogo As String
Public Property Let Co_Nm(ByVal vData As String)
   mvarCo_Nm = vData
End Property
Public Property Get Co_Nm() As String
   Co_Nm = mvarCo_Nm
End Property
Public Property Let Co_Type(ByVal vData As String)
   mvarCo_Type = vData
End Property
Public Property Get Co_Type() As String
   Co_Type = mvarCo_Type
End Property
Public Property Let Address(ByVal vData As String)
   mvarAddress = vData
End Property
Public Property Get Address() As String
   Address = mvarAddress
End Property
Public Property Let Phone(ByVal vData As String)
   mvarPhone = vData
End Property
Public Property Get Phone() As String
   Phone = mvarPhone
End Property
Public Property Let Fax(ByVal vData As String)
   mvarFax = vData
End Property
Public Property Get Fax() As String
   Fax = mvarFax
End Property
Public Property Let E_mail(ByVal vData As String)
   mvarE_mail = vData
End Property
Public Property Get E_mail() As String
   E_mail = mvarE_mail
End Property
Public Property Let Notes(ByVal vData As String)
   mvarNotes = vData
End Property
Public Property Get Notes() As String
   Notes = mvarNotes
End Property
Public Property Let Logo(ByVal vData As String)
   mvarLogo = vData
End Property
Public Property Get Logo() As String
   Logo = mvarLogo
End Property
Public Property Let Connstring(ByVal vData As String)
   mvarConnstring = vData
End Property
Public Property Get Connstring() As String
   Connstring = mvarConnstring
End Property
Public Function Save() As String

'On ERROR Resume Next

If mvarCo_Nm = Empty Then Exit Function

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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, mvarCo_Nm)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 50, mvarCo_Type)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 200, mvarAddress)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 50, mvarPhone)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 50, mvarFax)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 50, mvarE_mail)
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 200, mvarNotes)
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 200, mvarLogo)
    cmd.Parameters.Append Param8
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Pkg_St_Company.St_Company_Info_IU(?, ?, ?, ?, ?, ?, ?, ?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False


End Function


Public Function Get_Company_Info() As Recordset

On Error Resume Next

    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset
        
    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn

    cmd.CommandType = adCmdText
    cmd.CommandText = "Select * from St_Company_Info"
    cmd.Properties("IRowsetChange") = True
    cmd.Properties("Updatability") = 7
    
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
   
    If Not (RS.EOF Or RS.BOF) Then
        Set Get_Company_Info = RS
    End If

    If Not (RS.EOF And RS.BOF) Then
        mvarCo_Nm = RS("Co_Nm")
        mvarCo_Type = RS("Co_Type")
        mvarAddress = RS("Address")
        mvarPhone = RS("Phone")
        mvarFax = RS("Fax")
        mvarE_mail = RS("E_mail")
        mvarNotes = RS("Notes")
        mvarLogo = RS("Logo")
       
    Else
        mvarCo_Nm = ""
        mvarCo_Type = ""
        mvarAddress = ""
        mvarPhone = ""
        mvarFax = ""
        mvarE_mail = ""
        mvarNotes = ""
        mvarLogo = ""
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
