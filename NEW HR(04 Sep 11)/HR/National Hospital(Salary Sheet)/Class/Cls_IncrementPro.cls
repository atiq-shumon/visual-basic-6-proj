VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_IncrementPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarConnstring As String 'local copy
Private mvarEmp_ID As String 'local copy
Private mvarAMOUNT As Double 'local copy
Private mvarLAST_DT_INCRE As Date 'local copy
Private mvarNEXT_DT_INCRE As Date 'local copy
Private mvarLEAVE_WITHOUTPAY As Double 'local copy
Private mvarLEAVE_HALFPAY As Double 'local copy
Private mvarUPDATE_DATE As Date 'local copy
Private mvarLAST_PROM_DATE As Date 'local copy
Private mvarP_PROMOTION_EFF_DT As Date 'local copy
Private mvarCURRENT_BASIC As Double 'local copy
Private mvarCURRENT_SCALE As String 'local copy
Private mvarENTRY_DATE As Date 'local copy
Private mvarENTRY_BY As String 'local copy
Private mvarBEGIN_YEAR As String 'local copy
Private mvarEND_YEAR As String 'local copy
Private mvarEMP_CONTRUBUTION As Double 'local copy
Private mvarEMPLOYEER_CONTRIBUTION As Double 'local copy
Private mvarINTEREST_AMOUNT As Double 'local copy
Private mvarOPENING_AMOUNT As Double 'local copy
Private mvarCLOSEING_AMOUNT As Double 'local copy
Private mvarFROMDEGINATION As String 'local copy
Private mvarTODESIGNATION As String 'local copy
Private mvarFROMDEPARTMENT As String 'local copy
Private mvarTODEPARTMENT As String 'local copy
Private mvarLastBasic As Double 'local copy
Private mvarLastScale As String 'local copy
Private mvarRegistrationPgNo As String 'local copy
Private mvarCommentsonIncre As String 'local copy
Private mvarEmpClass As String 'local copy
Public Property Let EmpClass(ByVal vData As String)
    mvarEmpClass = vData
End Property
Public Property Get EmpClass() As String
    EmpClass = mvarEmpClass
End Property
Public Property Let CommentsonIncre(ByVal vData As String)
    mvarCommentsonIncre = vData
End Property
Public Property Get CommentsonIncre() As String
    CommentsonIncre = mvarCommentsonIncre
End Property
Public Property Let RegistrationPgNo(ByVal vData As String)
    mvarRegistrationPgNo = vData
End Property
Public Property Get RegistrationPgNo() As String
    RegistrationPgNo = mvarRegistrationPgNo
End Property
Public Property Let LastScale(ByVal vData As String)
    mvarLastScale = vData
End Property
Public Property Get LastScale() As String
    LastScale = mvarLastScale
End Property
Public Property Let LastBasic(ByVal vData As Double)
    mvarLastBasic = vData
End Property
Public Property Get LastBasic() As Double
    LastBasic = mvarLastBasic
End Property
Public Property Let TODEPARTMENT(ByVal vData As String)
    mvarTODEPARTMENT = vData
End Property
Public Property Get TODEPARTMENT() As String
    TODEPARTMENT = mvarTODEPARTMENT
End Property
Public Property Let FROMDEPARTMENT(ByVal vData As String)
    mvarFROMDEPARTMENT = vData
End Property
Public Property Get FROMDEPARTMENT() As String
    FROMDEPARTMENT = mvarFROMDEPARTMENT
End Property
Public Property Let TODESIGNATION(ByVal vData As String)
    mvarTODESIGNATION = vData
End Property
Public Property Get TODESIGNATION() As String
    TODESIGNATION = mvarTODESIGNATION
End Property
Public Property Let FROMDEGINATION(ByVal vData As String)
    mvarFROMDEGINATION = vData
End Property
Public Property Get FROMDEGINATION() As String
    FROMDEGINATION = mvarFROMDEGINATION
End Property
Public Property Let CLOSEING_AMOUNT(ByVal vData As Double)
    mvarCLOSEING_AMOUNT = vData
End Property
Public Property Get CLOSEING_AMOUNT() As Double
    CLOSEING_AMOUNT = mvarCLOSEING_AMOUNT
End Property
Public Property Let OPENING_AMOUNT(ByVal vData As Double)
    mvarOPENING_AMOUNT = vData
End Property
Public Property Get OPENING_AMOUNT() As Double
    OPENING_AMOUNT = mvarOPENING_AMOUNT
End Property
Public Property Let INTEREST_AMOUNT(ByVal vData As Double)
    mvarINTEREST_AMOUNT = vData
End Property
Public Property Get INTEREST_AMOUNT() As Double
    INTEREST_AMOUNT = mvarINTEREST_AMOUNT
End Property
Public Property Let Employeer_Contribution(ByVal vData As Double)
    mvarEMPLOYEER_CONTRIBUTION = vData
End Property
Public Property Get Employeer_Contribution() As Double
    Employeer_Contribution = mvarEMPLOYEER_CONTRIBUTION
End Property
Public Property Let EMP_CONTRUBUTION(ByVal vData As Double)
    mvarEMP_CONTRUBUTION = vData
End Property
Public Property Get EMP_CONTRUBUTION() As Double
    EMP_CONTRUBUTION = mvarEMP_CONTRUBUTION
End Property
Public Property Let END_YEAR(ByVal vData As String)
    mvarEND_YEAR = vData
End Property
Public Property Get END_YEAR() As String
    END_YEAR = mvarEND_YEAR
End Property
Public Property Let BEGIN_YEAR(ByVal vData As String)
    mvarBEGIN_YEAR = vData
End Property
Public Property Get BEGIN_YEAR() As String
    BEGIN_YEAR = mvarBEGIN_YEAR
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
Public Property Let CURRENT_SCALE(ByVal vData As String)
    mvarCURRENT_SCALE = vData
End Property
Public Property Get CURRENT_SCALE() As String
    CURRENT_SCALE = mvarCURRENT_SCALE
End Property
Public Property Let CURRENT_BASIC(ByVal vData As Double)
   mvarCURRENT_BASIC = vData
End Property
Public Property Get CURRENT_BASIC() As Double
   CURRENT_BASIC = mvarCURRENT_BASIC
End Property
Public Property Let P_PROMOTION_EFF_DT(ByVal vData As Date)
    mvarP_PROMOTION_EFF_DT = vData
End Property
Public Property Get P_PROMOTION_EFF_DT() As Date
   P_PROMOTION_EFF_DT = mvarP_PROMOTION_EFF_DT
End Property
Public Property Let LAST_PROM_DATE(ByVal vData As Date)
    mvarLAST_PROM_DATE = vData
End Property
Public Property Get LAST_PROM_DATE() As Date
    LAST_PROM_DATE = mvarLAST_PROM_DATE
End Property
Public Property Let UPDATE_DATE(ByVal vData As Date)
    mvarUPDATE_DATE = vData
End Property
Public Property Get UPDATE_DATE() As Date
    UPDATE_DATE = mvarUPDATE_DATE
End Property
Public Property Let LEAVE_HALFPAY(ByVal vData As Double)
    mvarLEAVE_HALFPAY = vData
End Property
Public Property Get LEAVE_HALFPAY() As Double
    LEAVE_HALFPAY = mvarLEAVE_HALFPAY
End Property
Public Property Let LEAVE_WITHOUTPAY(ByVal vData As Double)
    mvarLEAVE_WITHOUTPAY = vData
End Property
Public Property Get LEAVE_WITHOUTPAY() As Double
    LEAVE_WITHOUTPAY = mvarLEAVE_WITHOUTPAY
End Property
Public Property Let NEXT_DT_INCRE(ByVal vData As Date)
    mvarNEXT_DT_INCRE = vData
End Property
Public Property Get NEXT_DT_INCRE() As Date
    NEXT_DT_INCRE = mvarNEXT_DT_INCRE
End Property
Public Property Let LAST_DT_INCRE(ByVal vData As Date)
    mvarLAST_DT_INCRE = vData
End Property
Public Property Get LAST_DT_INCRE() As Date
    LAST_DT_INCRE = mvarLAST_DT_INCRE
End Property
Public Property Let AMOUNT(ByVal vData As Double)
    mvarAMOUNT = vData
End Property
Public Property Get AMOUNT() As Double
    AMOUNT = mvarAMOUNT
End Property
Public Property Let Emp_ID(ByVal vData As String)
    mvarEmp_ID = vData
End Property
Public Property Get Emp_ID() As String
    Emp_ID = mvarEmp_ID
End Property
Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property
Public Function Increment_Save()
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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adDouble, adParamInput, 10, mvarAMOUNT)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, mvarLAST_DT_INCRE)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 15, mvarNEXT_DT_INCRE)
    cmd.Parameters.Append Param4

    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 15, mvarLEAVE_WITHOUTPAY)
    cmd.Parameters.Append Param5
    
    
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 15, mvarLEAVE_HALFPAY)
    cmd.Parameters.Append Param6

    Set Param7 = cmd.CreateParameter("param7", adDate, adParamInput, 15, mvarUPDATE_DATE)
    cmd.Parameters.Append Param7
    
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 25, mvarRegistrationPgNo)
    cmd.Parameters.Append Param8
    
    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 150, mvarCommentsonIncre)
    cmd.Parameters.Append Param9
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Increment_Information_save(?,?,?,?,?,?,?,?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
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
Public Function Increment_Delete()
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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, mvarLAST_DT_INCRE)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 15, mvarNEXT_DT_INCRE)
    cmd.Parameters.Append Param3

    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL INCREMENT_RECORD_Delete(?,?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function
Public Function Promotion_Save()
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
    Dim Param13 As New ADODB.Parameter
    Dim Param14 As New ADODB.Parameter
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, mvarLAST_PROM_DATE)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, mvarP_PROMOTION_EFF_DT)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 15, mvarCURRENT_BASIC)
    cmd.Parameters.Append Param4

    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 15, mvarCURRENT_SCALE)
    cmd.Parameters.Append Param5
    
    
    Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, 15, mvarENTRY_DATE)
    cmd.Parameters.Append Param6

    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 15, mvarENTRY_BY)
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 50, mvarFROMDEGINATION)
    cmd.Parameters.Append Param8

    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 50, mvarTODESIGNATION)
    cmd.Parameters.Append Param9
    
    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 50, mvarFROMDEPARTMENT)
    cmd.Parameters.Append Param10

    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 50, mvarTODEPARTMENT)
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adDouble, adParamInput, 15, mvarLastBasic)
    cmd.Parameters.Append Param12

    Set Param13 = cmd.CreateParameter("param13", adVarChar, adParamInput, 50, mvarLastScale)
    cmd.Parameters.Append Param13
    
    Set Param14 = cmd.CreateParameter("param14", adVarChar, adParamInput, 2, mvarEmpClass)
    cmd.Parameters.Append Param14
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL PromotionInformation_Save(?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    Set RS = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
End Function
Public Function PF_Closing_Save()
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
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, mvarBEGIN_YEAR)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 20, mvarEND_YEAR)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 15, mvarEMP_CONTRUBUTION)
    cmd.Parameters.Append Param4

    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 15, mvarEMPLOYEER_CONTRIBUTION)
    cmd.Parameters.Append Param5
    
    
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 15, mvarINTEREST_AMOUNT)
    cmd.Parameters.Append Param6

    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 15, mvarOPENING_AMOUNT)
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adDouble, adParamInput, 15, mvarCLOSEING_AMOUNT)
    cmd.Parameters.Append Param8
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Closing_PF_Balance_save(?,?,?,?,?,?,?,?)}"
    
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
    cmd.CommandText = "Select EMP_ID,BEGIN_YEAR,END_YEAR,EMP_CONTRUBUTION," + _
            " EMPLOYEER_CONTRIBUTION,INTEREST_AMOUNT,OPENING_AMOUNT,CLOSEING_AMOUNT " + _
            " from CLOSING_PF_BALANCE order by EMP_ID"
    cmd.Properties("IRowsetChange") = True
    cmd.Properties("Updatability") = 7
    
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
   
    If Not (RS.EOF Or RS.BOF) Then
        Set GetAll = RS
        Exit Function
    End If
    
End Function
Public Function PF_Closing_Delete()
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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, mvarBEGIN_YEAR)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 20, mvarEND_YEAR)
    cmd.Parameters.Append Param3
    
    
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Closing_PF_Balance_Delete(?,?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function

