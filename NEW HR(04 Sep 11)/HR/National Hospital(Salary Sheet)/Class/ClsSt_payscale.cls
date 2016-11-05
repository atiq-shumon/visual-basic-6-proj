VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSt_payscale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarBEGIN_DATE As Date
Private mvarEND_DATE As Date
Private mvarENTRY_DATE As Date
Private mvarENTRY_BY As String
Private mvarConnstring As String
Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property
Public Property Let ENTRY_BY(ByVal vData As String)
    mvarENTRY_BY = vData
End Property
Public Property Get ENTRY_BY() As String
    ENTRY_BY = mvarENTRY_BY
End Property
Public Property Let ENTRY_DATE(ByVal vData As Date)
    mvarENTRY_DATE = vData
End Property
Public Property Get ENTRY_DATE() As Date
    ENTRY_DATE = mvarENTRY_DATE
End Property
Public Property Let END_DATE(ByVal vData As Date)
    mvarEND_DATE = vData
End Property
Public Property Get END_DATE() As Date
    END_DATE = mvarEND_DATE
End Property
Public Property Let BEGIN_DATE(ByVal vData As Date)
    mvarBEGIN_DATE = vData
End Property
Public Property Get BEGIN_DATE() As Date
    BEGIN_DATE = mvarBEGIN_DATE
End Property
Public Function Pay_Scale_SetUp_Save() As String

    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim Param3 As New ADODB.Parameter
    Dim Param4 As New ADODB.Parameter
        
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 12, mvarBEGIN_DATE)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 12, mvarEND_DATE)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 12, mvarENTRY_DATE)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 10, mvarENTRY_BY)
    cmd.Parameters.Append Param4
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Fiscal_Yr_setup_save(?, ?, ?, ? )}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function
Public Function GetAll() As Recordset

   On Error Resume Next


    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset
    
    
    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn

    cmd.CommandType = adCmdText
    cmd.CommandText = "Select BEGIN_DATE,END_DATE,ENTRY_DATE,ENTRY_BY from fISCAL_Year_SetUp order by TRACE_ID"
    cmd.Properties("IRowsetChange") = True
    cmd.Properties("Updatability") = 7
    
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
   
    If Not (RS.EOF Or RS.BOF) Then
        Set GetAll = RS
        Exit Function
    End If
    

End Function
Public Function Pay_Scale_SetUp_Delete() As String

    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim Param3 As New ADODB.Parameter
    Dim Param4 As New ADODB.Parameter
        
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 12, mvarBEGIN_DATE)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 12, mvarEND_DATE)
    cmd.Parameters.Append Param2
    
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Fiscal_Yr_setup_Delete(?, ?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function

