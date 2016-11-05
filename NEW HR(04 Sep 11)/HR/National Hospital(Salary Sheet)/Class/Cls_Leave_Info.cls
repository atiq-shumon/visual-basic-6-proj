VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_Leave_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarEmp_ID As String
Private mvarLEAVE_CODE As String
Private mvarNO_OF_DAYS_LEAVE As String
Private mvarDATE_OF_APPLIED As Date
Private mvarDATE_OF_JOINFROMLF As Date
Private mvarREMAIN_CASUALLEAVE As String
Private mvarConnstring As String
Private mvarUPDATE_DATE As Date 'local copy
Private mvarYEAR As String 'local copy
Private mvarEXTRA_LEAVE_REMAIN As String 'local copy
Private mvarTOTAL_EARN_LEAVE As String 'local copy
'local variable(s) to hold property value(s)
Private mvarLAST_RCREAT_DATE As Date 'local copy
Private mvarNEXT_RCREAT_DATE As Date 'local copy
Private mvarENJOED_RECRT As Long 'local copy
Private mvarENJOED_STATUS As Integer 'local copy
Private mvarRECRE_ST_DT As Date 'local copy
Private mvarRECRE_END_DT As Date 'local copy
Private mvarCreatedBy As String 'local copy
Private mvarCreatedDate As Date 'local copy
Private mvarUpdateBy As String 'local copy
Public Property Let UpdateBy(ByVal vData As String)
    mvarUpdateBy = vData
End Property
Public Property Get UpdateBy() As String
    UpdateBy = mvarUpdateBy
End Property
Public Property Let CreatedDate(ByVal vData As Date)
    mvarCreatedDate = vData
End Property
Public Property Get CreatedDate() As Date
    CreatedDate = mvarCreatedDate
End Property
Public Property Let CreatedBy(ByVal vData As String)
    mvarCreatedBy = vData
End Property
Public Property Get CreatedBy() As String
    CreatedBy = mvarCreatedBy
End Property
Public Property Let RECRE_END_DT(ByVal vData As Date)
    mvarRECRE_END_DT = vData
End Property
Public Property Get RECRE_END_DT() As Date
    RECRE_END_DT = mvarRECRE_END_DT
End Property
Public Property Let RECRE_ST_DT(ByVal vData As Date)
    mvarRECRE_ST_DT = vData
End Property
Public Property Get RECRE_ST_DT() As Date
    RECRE_ST_DT = mvarRECRE_ST_DT
End Property
Public Property Let ENJOED_STATUS(ByVal vData As Integer)
    mvarENJOED_STATUS = vData
End Property
Public Property Get ENJOED_STATUS() As Integer
    ENJOED_STATUS = mvarENJOED_STATUS
End Property
Public Property Let ENJOED_RECRT(ByVal vData As Long)
    mvarENJOED_RECRT = vData
End Property
Public Property Get ENJOED_RECRT() As Long
    ENJOED_RECRT = mvarENJOED_RECRT
End Property
Public Property Let NEXT_RCREAT_DATE(ByVal vData As Date)
    mvarNEXT_RCREAT_DATE = vData
End Property
Public Property Get NEXT_RCREAT_DATE() As Date
    NEXT_RCREAT_DATE = mvarNEXT_RCREAT_DATE
End Property
Public Property Let LAST_RCREAT_DATE(ByVal vData As Date)
    mvarLAST_RCREAT_DATE = vData
End Property
Public Property Get LAST_RCREAT_DATE() As Date
    LAST_RCREAT_DATE = mvarLAST_RCREAT_DATE
End Property
Public Property Let TOTAL_EARN_LEAVE(ByVal vData As String)
    mvarTOTAL_EARN_LEAVE = vData
End Property
Public Property Get TOTAL_EARN_LEAVE() As String
    TOTAL_EARN_LEAVE = mvarTOTAL_EARN_LEAVE
End Property
Public Property Let EXTRA_LEAVE_REMAIN(ByVal vData As String)
    mvarEXTRA_LEAVE_REMAIN = vData
End Property
Public Property Get EXTRA_LEAVE_REMAIN() As String
    EXTRA_LEAVE_REMAIN = mvarEXTRA_LEAVE_REMAIN
End Property
Public Property Let YEAR(ByVal vData As String)
    mvarYEAR = vData
End Property
Public Property Get YEAR() As String
    YEAR = mvarYEAR
End Property
Public Property Let UPDATE_DATE(ByVal vData As Date)
    mvarUPDATE_DATE = vData
End Property
Public Property Get UPDATE_DATE() As Date
    UPDATE_DATE = mvarUPDATE_DATE
End Property
Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property
Public Property Let REMAIN_CASUALLEAVE(ByVal vData As String)
    mvarREMAIN_CASUALLEAVE = vData
End Property
Public Property Get REMAIN_CASUALLEAVE() As String
    REMAIN_CASUALLEAVE = mvarREMAIN_CASUALLEAVE
End Property
Public Property Let DATE_OF_JOINFROMLF(ByVal vData As Date)
    mvarDATE_OF_JOINFROMLF = vData
End Property
Public Property Get DATE_OF_JOINFROMLF() As Date
    DATE_OF_JOINFROMLF = mvarDATE_OF_JOINFROMLF
End Property
Public Property Let DATE_OF_APPLIED(ByVal vData As Date)
    mvarDATE_OF_APPLIED = vData
End Property
Public Property Get DATE_OF_APPLIED() As Date
    DATE_OF_APPLIED = mvarDATE_OF_APPLIED
End Property
Public Property Let NO_OF_DAYS_LEAVE(ByVal vData As String)
    mvarNO_OF_DAYS_LEAVE = vData
End Property
Public Property Get NO_OF_DAYS_LEAVE() As String
    NO_OF_DAYS_LEAVE = mvarNO_OF_DAYS_LEAVE
End Property
Public Property Let LEAVE_CODE(ByVal vData As String)
    mvarLEAVE_CODE = vData
End Property
Public Property Get LEAVE_CODE() As String
    LEAVE_CODE = mvarLEAVE_CODE
End Property
Public Property Let Emp_ID(ByVal vData As String)
    mvarEmp_ID = vData
End Property
Public Property Get Emp_ID() As String
    Emp_ID = mvarEmp_ID
End Property
Public Function Save()
On Error GoTo Errdess
    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim Rs As New ADODB.Recordset

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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, mvarLEAVE_CODE)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, mvarNO_OF_DAYS_LEAVE)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 15, mvarDATE_OF_APPLIED)
    cmd.Parameters.Append Param4

    Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 15, mvarDATE_OF_JOINFROMLF)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 10, mvarREMAIN_CASUALLEAVE)
    cmd.Parameters.Append Param6
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Leave_Application_save(?,?,?,?,?,?)}"
    
    Set Rs = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
Exit Function
Errdess:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Function
'Leave_Application_Delete(
Public Sub Show_Message()
       
    Dim conn As New Connection
    Dim cmd As New Command
    Dim Rs As New Recordset
    
    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "Select * from Temp"
    cmd.Properties("IRowsetChange") = True
    cmd.Properties("Updatability") = 7
    
    Rs.CursorLocation = adUseClient
    Rs.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic

    MsgBox Rs.Fields(0), vbInformation + vbOKOnly, "Message"

End Sub
Public Function Delete()
    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim Rs As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
       
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, mvarDATE_OF_APPLIED)
    cmd.Parameters.Append Param2
       
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Leave_Application_Delete(?,?)}"
    
    Set Rs = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function
Public Function Earn_Leave_Save()

    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim Rs As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim Param3 As New ADODB.Parameter
    Dim Param4 As New ADODB.Parameter
    Dim Param5 As New ADODB.Parameter
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, mvarTOTAL_EARN_LEAVE)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, mvarEXTRA_LEAVE_REMAIN)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 15, mvarYEAR)
    cmd.Parameters.Append Param4

    Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 15, mvarUPDATE_DATE)
    cmd.Parameters.Append Param5
    
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Earn_Leave_save(?,?,?,?,?)}"
    
    Set Rs = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function
Public Function Creation_Leave_Save()
On Error GoTo Errdess
    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim Rs As New ADODB.Recordset

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
    
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, mvarLAST_RCREAT_DATE)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, NEXT_RCREAT_DATE)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 15, mvarENJOED_RECRT)
    cmd.Parameters.Append Param4

    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 2, mvarENJOED_STATUS)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, 10, mvarRECRE_ST_DT)
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adDate, adParamInput, 10, mvarRECRE_END_DT)
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adDate, adParamInput, 10, mvarCreatedDate)
    cmd.Parameters.Append Param8

    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 10, mvarCreatedBy)
    cmd.Parameters.Append Param9
    
    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 10, mvarUpdateBy)
    cmd.Parameters.Append Param10
    
     Set Param11 = cmd.CreateParameter("param11", adDate, adParamInput, 10, mvarUPDATE_DATE)
    cmd.Parameters.Append Param11
    
    
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Recreation_Save(?,?,?,?,?,?,?,?,?,?,?)}"
    
    Set Rs = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
Exit Function
Errdess:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Function

