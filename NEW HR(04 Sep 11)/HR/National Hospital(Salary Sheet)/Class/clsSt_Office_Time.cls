VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSt_Office_Time"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarStart_Time As String 'local copy
Private mvarEnd_Time As String 'local copy
Private mvarRelax_Time As String 'local copy
Private mvarAbsent_Time As String 'local copy
Private mvarSpecial_Start_Time As String 'local copy
Private mvarSpecial_End_Time As String 'local copy
Private mvarSpecial_Day As String 'local copy
Private mvarEffect_Dt As Date 'local copy
Private mvarConnstring As String 'local copy
Private mvarTrack_Id As Double 'local copy
Public Property Let Track_Id(ByVal vData As Double)
    mvarTrack_Id = vData
End Property
Public Property Get Track_Id() As Double
    Track_Id = mvarTrack_Id
End Property
Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property

Public Property Let Effect_Dt(ByVal vData As Date)
    mvarEffect_Dt = vData
End Property
Public Property Get Effect_Dt() As Date
    Effect_Dt = mvarEffect_Dt
End Property
Public Property Let Special_Day(ByVal vData As String)
    mvarSpecial_Day = vData
End Property
Public Property Get Special_Day() As String
    Special_Day = mvarSpecial_Day
End Property
Public Property Let Special_End_Time(ByVal vData As String)
    mvarSpecial_End_Time = vData
End Property
Public Property Get Special_End_Time() As String
    Special_End_Time = mvarSpecial_End_Time
End Property
Public Property Let Special_Start_Time(ByVal vData As String)
    mvarSpecial_Start_Time = vData
End Property
Public Property Get Special_Start_Time() As String
    Special_Start_Time = mvarSpecial_Start_Time
End Property
Public Property Let Absent_Time(ByVal vData As String)
    mvarAbsent_Time = vData
End Property
Public Property Get Absent_Time() As String
    Absent_Time = mvarAbsent_Time
End Property
Public Property Let Relax_Time(ByVal vData As String)
    mvarRelax_Time = vData
End Property
Public Property Get Relax_Time() As String
    Relax_Time = mvarRelax_Time
End Property
Public Property Let End_Time(ByVal vData As String)
    mvarEnd_Time = vData
End Property
Public Property Get End_Time() As String
    End_Time = mvarEnd_Time
End Property
Public Property Let Start_Time(ByVal vData As String)
    mvarStart_Time = vData
End Property
Public Property Get Start_Time() As String
    Start_Time = mvarStart_Time
End Property
Public Function Delete() As Variant
End Function
Public Function GetAll() As Recordset
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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 8, mvarStart_Time)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamOutput)
    cmd.Parameters.Append Param2
    '----------------------------------------------------------------------------------
    ' Enable PLSQLRSet property
    
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL Pkg_St_Company.St_Office_Time_GetX(?,?)}"

    Set RS = cmd.Execute

    ' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
    
    Set GetX = RS
    
    If Not (RS.EOF And RS.BOF) Then
       mvarStart_Time = RS("Start_Time ")
       
       mvarEnd_Time = RS("End_Time")
       
       mvarRelax_Time = RS("Relax_Time")
       
       mvarAbsent_Time = RS("Absent_Time ")
       
       mvarSpecial_Start_Time = RS("Special_Start_Time")
       
       mvarSpecial_End_Time = RS("Special_End_Time")
       
       mvarSpecial_Day = RS("Special_Day  ")
       
       mvarEffect_Dt = RS("Effect_Dt")
       
    Else
        
       mvarStart_Time = ""
       mvarEnd_Time = ""
       mvarRelax_Time = ""
       mvarAbsent_Time = ""
       mvarSpecial_Start_Time = ""
       mvarSpecial_End_Time = ""
       mvarSpecial_Day = ""
       mvarEffect_Dt = ""
    End If
Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Emp_Per_info Table" + vbCr + vbCr + Err.Description
End Function
Public Function Save() As Variant
'If mvarStart_Time = Empty Then Exit Function

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
        
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 8, mvarStart_Time)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 8, mvarEnd_Time)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 5, mvarRelax_Time)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 8, mvarAbsent_Time)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 8, mvarSpecial_Start_Time)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 8, mvarSpecial_End_Time)
    cmd.Parameters.Append Param6

    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 10, mvarSpecial_Day)
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adDate, adParamInput, 10, Date)
    cmd.Parameters.Append Param8
    
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 10, 5)
    cmd.Parameters.Append Param9
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL OfficeTime_SetUp_save(?,?,?,?,?,?,?,?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
End Function
