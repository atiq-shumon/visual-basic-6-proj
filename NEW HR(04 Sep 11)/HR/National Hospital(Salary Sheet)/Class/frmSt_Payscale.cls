VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "St_Payscale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Private mvarConnstring As String
Private mvarScale_code As String
Private mvarYr_Ref As Integer
Private mvarStr_Basic As Integer
Private mvarIncr As Integer
Private mvarEnd_basic As Integer
Private mvarEb_incr As Integer
Private mvarEb_end As Integer
Private mvarHR As Integer
Private mvarMinHR As Integer
Private mvarMED As Integer
Private mvarCONV As Integer
Private mvarTFN As Integer

Public Property Let Scale_code(ByVal vData As String)
   mvarScale_code = vData
End Property
Public Property Get Scale_code() As String
   Scale_code = mvarScale_code
End Property
Public Property Let Yr_Ref(ByVal vData As Integer)
   mvarYr_Ref = vData
End Property
Public Property Get Yr_Ref() As Integer
   Yr_Ref = mvarYr_Ref
End Property
Public Property Let Str_Basic(ByVal vData As Integer)
   mvarStr_Basic = vData
End Property
Public Property Get Str_Basic() As Integer
   Str_Basic = mvarStr_Basic
End Property
Public Property Let Incr(ByVal vData As Integer)
   mvarIncr = vData
End Property
Public Property Get Incr() As Integer
   Incr = mvarIncr
End Property
Public Property Let End_basic(ByVal vData As Integer)
   mvarEnd_basic = vData
End Property
Public Property Get End_basic() As Integer
   End_basic = mvarEnd_basic
End Property
Public Property Let Eb_incr(ByVal vData As Integer)
   mvarEb_incr = vData
End Property
Public Property Get Eb_incr() As Integer
   Eb_incr = mvarEb_incr
End Property
Public Property Let Eb_end(ByVal vData As Integer)
   mvarEb_end = vData
End Property
Public Property Get Eb_end() As Integer
   Eb_end = mvarEb_end
End Property
Public Property Let HR(ByVal vData As Integer)
   mvarHR = vData
End Property
Public Property Get HR() As Integer
   HR = mvarHR
End Property
Public Property Let MinHR(ByVal vData As Integer)
   mvarMinHR = vData
End Property
Public Property Get MinHR() As Integer
   MinHR = mvarMinHR
End Property
Public Property Let MED(ByVal vData As Integer)
   mvarMED = vData
End Property
Public Property Get MED() As Integer
   MED = mvarMED
End Property
Public Property Let CONV(ByVal vData As Integer)
   mvarCONV = vData
End Property
Public Property Get CONV() As Integer
   CONV = mvarCONV
End Property
Public Property Let TFN(ByVal vData As Integer)
   mvarTFN = vData
End Property
Public Property Get TFN() As Integer
   TFN = mvarTFN
End Property
Public Property Let Connstring(ByVal vData As String)
   mvarConnstring = vData
End Property
Public Property Get Connstring() As String
   Connstring = mvarConnstring
End Property
Public Function Save()

   If mvarScale_code = Empty Then Exit Function

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
     
        
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarScale_code)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adInteger, adParamInput, 4, mvarYr_Ref)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adInteger, adParamInput, 4, mvarStr_Basic)
    cmd.Parameters.Append Param3
    Set Param4 = cmd.CreateParameter("param4", adInteger, adParamInput, 4, mvarIncr)
    cmd.Parameters.Append Param4

    Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 4, mvarEnd_basic)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adInteger, adParamInput, 4, mvarEb_incr)
    cmd.Parameters.Append Param6
    Set Param7 = cmd.CreateParameter("param7", adInteger, adParamInput, 4, mvarEb_end)
    cmd.Parameters.Append Param7

    Set Param8 = cmd.CreateParameter("param8", adInteger, adParamInput, 4, mvarHR)
    cmd.Parameters.Append Param8
    
    Set Param9 = cmd.CreateParameter("param9", adInteger, adParamInput, 4, mvarMinHR)
    cmd.Parameters.Append Param9
    Set Param10 = cmd.CreateParameter("param10", adInteger, adParamInput, 4, mvarMED)
    cmd.Parameters.Append Param10

    Set Param11 = cmd.CreateParameter("param11", adInteger, adParamInput, 4, mvarCONV)
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adInteger, adParamInput, 4, mvarTFN)
    cmd.Parameters.Append Param12
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Pkg_ST_Company.St_Payscale_IU(?,?,?, ?, ?,?,?, ?, ?,?,?, ?)}"
    
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
    cmd.CommandText = "Delete from St_Payscale where Scale_Code='" + mvarScale_code + "'"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
    Exit Function

End Function
Public Function GetX() As Recordset

  On Error Resume Next

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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarScale_code)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamOutput)
    cmd.Parameters.Append Param2
    '----------------------------------------------------------------------------------
    ' Enable PLSQLRSet property
    
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL Pkg_ST_Company.St_Payscale_GetX(?,?)}"

    Set RS = cmd.Execute

    ' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
    
    Set GetX = RS
    
    If Not (RS.EOF And RS.BOF) Then
      mvarScale_code = RS("Scale_code")
      mvarYr_Ref = RS("Yr_Ref")
      mvarStr_Basic = RS("Str_Basic")
      mvarIncr = RS("Incr")
      mvarEnd_basic = RS("End_Basic")
      mvarEb_incr = RS("Eb_incr")
      mvarEb_end = RS("Eb_end")
      mvarHR = RS("HR")
      mvarMinHR = RS("Min_HR")
      mvarMED = RS("Med")
      mvarCONV = RS("Conv")
      mvarTFN = RS("Tfn")
      
    Else
      
      mvarScale_code = ""
      mvarYr_Ref = 0
      mvarStr_Basic = 0
      mvarIncr = 0
      mvarEnd_basic = 0
      mvarEb_incr = 0
      mvarEb_end = 0
      mvarHR = 0
      mvarMinHR = 0
      mvarMED = 0
      mvarCONV = 0
      mvarTFN = 0
       
       
    End If
    Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Emp_info Table" + vbCr + vbCr + Err.Description
End Function

Public Function GetAll() As Recordset
   On Error Resume Next

    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset
    
    
    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn

    cmd.CommandType = adCmdText
    cmd.CommandText = "Select * from St_Payscale"
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
