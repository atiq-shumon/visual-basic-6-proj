VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSt_Holiday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarH_Type As Integer 'local copy
Private mvarTrack_Id As Integer 'local copy
Private mvarConnstring As String 'local copy
Private mvarHoliday_Name As String 'local copy
Private mvarYear_To_Show As String 'local copy
Private mvarFrom_Dt As Date 'local copy
Private mvarTo_Dt As Date 'local copy
Public Property Let To_Dt(ByVal vData As Date)
    mvarTo_Dt = vData
End Property
Public Property Get To_Dt() As Date
    To_Dt = mvarTo_Dt
End Property
Public Property Let From_Dt(ByVal vData As Date)
    mvarFrom_Dt = vData
End Property
Public Property Get From_Dt() As Date
    From_Dt = mvarFrom_Dt
End Property
Public Property Let Year_To_Show(ByVal vData As String)
    mvarYear_To_Show = vData
End Property
Public Property Get Year_To_Show() As String
    Year_To_Show = mvarYear_To_Show
End Property
Public Property Let Holiday_Name(ByVal vData As String)
    mvarHoliday_Name = vData
End Property
Public Property Get Holiday_Name() As String
    Holiday_Name = mvarHoliday_Name
End Property
Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property
Public Function Delete() As Variant
End Function
Public Function GetAll() As Recordset
End Function
'Public Property Let Track_Id(ByVal vData As Integer)
'    mvarTrack_Id = vData
'End Property
'Public Property Get Track_Id() As Integer
'    Set Track_Id = mvarTrack_Id
'End Property
Public Property Let H_Type(ByVal vData As Integer)
    mvarH_Type = vData
End Property
Public Property Get H_Type() As Integer
    H_Type = mvarH_Type
End Property
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
    Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, mvarTo_Dt)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamOutput)
    cmd.Parameters.Append Param2
    '----------------------------------------------------------------------------------
    ' Enable PLSQLRSet property
    
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL Pkg_St_Company.St_Holiday_GetX(?,?)}"

    Set RS = cmd.Execute

    ' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
    
    Set GetX = RS
    
    If Not (RS.EOF And RS.BOF) Then
       mvarHoliday_Name = RS("Holiday_Name ")
       mvarH_Type = RS("H_Type ")
       mvarFrom_Dt = RS("From_Dt ")
       mvarTo_Dt = RS("To_Dt ")
       mvarTrack_Id = RS("Track_Id ")
    Else
       mvarHoliday_Name = ""
       mvarH_Type = ""
       mvarFrom_Dt = ""
       mvarTo_Dt = ""
       mvarTrack_Id = ""
       
        
    End If
Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Emp_Per_info Table" + vbCr + vbCr + Err.Description
End Function
Public Function Save() As String
'If mvarTo_Dt = Empty Then Exit Function

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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 35, mvarHoliday_Name)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adInteger, adParamInput, 1, mvarH_Type)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, mvarFrom_Dt)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, mvarTo_Dt)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 4, mvarYear_To_Show)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adInteger, adParamInput, 10, mvarTrack_Id)
    cmd.Parameters.Append Param6
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Holiday_setup_save( ?, ?, ?, ?, ?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
End Function
