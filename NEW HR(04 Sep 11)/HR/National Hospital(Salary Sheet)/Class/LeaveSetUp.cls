VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LeaveSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarLEAVE_CODE As String 'local copy
Private mvarMax_Duration As String 'local copy
Private mvarArrey_Day As String 'local copy
Private mvarCarry_Arrear_Days As String 'local copy
Private mvarCarry_Max_Days As String 'local copy
Private mvarLeave_Name As String 'local copy
Private mvarConnstring As String 'local copy
Private mvarDays As String 'local copy
Public Property Let Days(ByVal vData As String)
    mvarDays = vData
End Property
Public Property Get Days() As String
    Days = mvarDays
End Property
Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property
Public Property Let Leave_Name(ByVal vData As String)
    mvarLeave_Name = vData
End Property
Public Property Get Leave_Name() As String
    Leave_Name = mvarLeave_Name
End Property
Public Property Let Carry_Max_Days(ByVal vData As String)
    mvarCarry_Max_Days = vData
End Property
Public Property Get Carry_Max_Days() As String
    Carry_Max_Days = mvarCarry_Max_Days
End Property
Public Property Let Carry_Arrear_Days(ByVal vData As String)
    mvarCarry_Arrear_Days = vData
End Property
Public Property Get Carry_Arrear_Days() As String
    Carry_Arrear_Days = mvarCarry_Arrear_Days
End Property
Public Property Let Arrey_Day(ByVal vData As String)
    mvarArrey_Day = vData
End Property
Public Property Get Arrey_Day() As String
    Arrey_Day = mvarArrey_Day
End Property
Public Property Let Max_Duration(ByVal vData As String)
    mvarMax_Duration = vData
End Property
Public Property Get Max_Duration() As String
    Max_Duration = mvarMax_Duration
End Property
Public Property Let LEAVE_CODE(ByVal vData As String)
    mvarLEAVE_CODE = vData
End Property
Public Property Get LEAVE_CODE() As String
    LEAVE_CODE = mvarLEAVE_CODE
End Property
Public Function Save()

   If mvarLEAVE_CODE = Empty Then Exit Function

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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarLEAVE_CODE)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adDouble, adParamInput, 2, mvarDays)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 2, mvarCarry_Arrear_Days)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 2, mvarCarry_Max_Days)
    cmd.Parameters.Append Param4

        
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 50, mvarLeave_Name)
    cmd.Parameters.Append Param5
    

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL leave_setup_save(?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    



End Function
