VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarLEAVE_CODE As String 'local copy
Private mvarLeave_Name As String 'local copy
Private mvarDays As Integer 'local copy
Private mvarCarry_Arrear_Days As Integer 'local copy
Private mvarCarry_Max_Days As Integer 'local copy
Private mvarConnstring As String 'local copy
Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property
Public Function Save() As String
If mvarLEAVE_CODE = Empty Then Exit Function

    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim Param3 As New ADODB.Parameter
    Dim Param4 As New ADODB.Parameter
    Dim Param5 As New ADODB.Parameter
    
        
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarLEAVE_CODE)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 8, mvarLeave_Name)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adInteger, adParamInput, 2, mvarDays)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adInteger, adParamInput, 1, mvarCarry_Arrear_Days)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 3, mvarCarry_Max_Days)
    cmd.Parameters.Append Param5
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Pkg_St_Company.St_Leave_IU( ?, ?, ?, ?, ?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function
Public Function GetX() As Recordset
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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, mvarLEAVE_CODE)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamOutput)
    cmd.Parameters.Append Param2
    '----------------------------------------------------------------------------------
    ' Enable PLSQLRSet property
    
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL Pkg_St_Company.St_Leave_GetX(?,?)}"

    Set RS = cmd.Execute

    ' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
    
    Set GetX = RS
    
    If Not (RS.EOF And RS.BOF) Then
       mvarLEAVE_CODE = RS("Leave_Code")
       
       mvarLeave_Name = RS("Leave_Name")
       
       mvarDays = RS("Days")
       
       mvarCarry_Arrear_Days = RS("Carry_Arrear_Days")
       
       mvarCarry_Max_Days = RS("Carry_Max_Days")
       
       
       
    Else
        
        mvarLEAVE_CODE = ""
       
       mvarLeave_Name = ""
       
       mvarDays = ""
       
       mvarCarry_Arrear_Days = ""
       
       mvarCarry_Max_Days = ""
       
        
        End If
    Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Emp_Per_info Table" + vbCr + vbCr + Err.Description
End Function
Public Function GetAll() As Recordset
End Function
Public Function Delete() As Variant
End Function
Public Property Let Carry_Max_Days(ByVal vData As Integer)
    mvarCarry_Max_Days = vData
End Property
Public Property Get Carry_Max_Days() As Integer
    Carry_Max_Days = mvarCarry_Max_Days
End Property
Public Property Let Carry_Arrear_Days(ByVal vData As Integer)
    mvarCarry_Arrear_Days = vData
End Property
Public Property Get Carry_Arrear_Days() As Integer
    Carry_Arrear_Days = mvarCarry_Arrear_Days
End Property
Public Property Let Days(ByVal vData As Integer)
    mvarDays = vData
End Property
Public Property Get Days() As Integer
    Days = mvarDays
End Property
Public Property Let Leave_Name(ByVal vData As String)
    mvarLeave_Name = vData
End Property
Public Property Get Leave_Name() As String
    Leave_Name = mvarLeave_Name
End Property
Public Property Let LEAVE_CODE(ByVal vData As String)
    mvarLEAVE_CODE = vData
End Property
Public Property Get LEAVE_CODE() As String
    LEAVE_CODE = mvarLEAVE_CODE
End Property
