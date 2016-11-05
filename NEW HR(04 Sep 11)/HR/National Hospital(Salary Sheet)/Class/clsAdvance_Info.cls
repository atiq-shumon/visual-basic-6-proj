VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdvance_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarAdv_Amt As Double 'local copy
Private mvarNum_Inst As String 'local copy
Private mvarNotes As String 'local copy
Private mvarREF_STATUS As String 'local copy
Private mvarEmp_ID As String 'local copy
Private mvarAdv_issue_dt As Date 'local copy
Private mvarAdvanceID As String 'local copy
Private mvarConnstring As String 'local copy
Private mvarPAID_INSTALLED As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarBalance As Integer 'local copy
Public Property Let Balance(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Balance = 5
    mvarBalance = vData
End Property


Public Property Get Balance() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Balance
    Balance = mvarBalance
End Property



Public Property Let PAID_INSTALLED(ByVal vData As Integer)
    mvarPAID_INSTALLED = vData
End Property
Public Property Get PAID_INSTALLED() As Integer
    PAID_INSTALLED = mvarPAID_INSTALLED
End Property
Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property
Public Property Let AdvanceID(ByVal vData As String)
    mvarAdvanceID = vData
End Property
Public Property Get AdvanceID() As String
    AdvanceID = mvarAdvanceID
End Property
Public Property Let Adv_issue_dt(ByVal vData As Date)
    mvarAdv_issue_dt = vData
End Property
Public Property Get Adv_issue_dt() As Date
    Adv_issue_dt = mvarAdv_issue_dt
End Property
Public Property Let Emp_ID(ByVal vData As String)
    mvarEmp_ID = vData
End Property
Public Property Get Emp_ID() As String
    Emp_ID = mvarEmp_ID
End Property
Public Property Let REF_STATUS(ByVal vData As String)
    mvarREF_STATUS = vData
End Property
Public Property Get REF_STATUS() As String
    REF_STATUS = mvarREF_STATUS
End Property
Public Property Let Notes(ByVal vData As String)
    mvarNotes = vData
End Property
Public Property Get Notes() As String
    Notes = mvarNotes
End Property
Public Property Let Num_Inst(ByVal vData As String)
    mvarNum_Inst = vData
End Property
Public Property Get Num_Inst() As String
    Num_Inst = mvarNum_Inst
End Property
Public Property Let Adv_Amt(ByVal vData As Double)
    mvarAdv_Amt = vData
End Property
Public Property Get Adv_Amt() As Double
    Adv_Amt = mvarAdv_Amt
End Property
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
    Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamOutput)
    cmd.Parameters.Append Param2
    '----------------------------------------------------------------------------------
    ' Enable PLSQLRSet property
    
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL Pkg_Advance_Info.Advance_Info_GetX(?,?)}"

    Set RS = cmd.Execute

    ' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
    
        
    Set GetX = RS

    
End Function
Public Function GetAll() As Recordset
   On Error Resume Next

    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset
    
    
    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn

    cmd.CommandType = adCmdText
    cmd.CommandText = "Select * from Advance_info where Emp_Id='" & mvarEmp_ID & "'"
    cmd.Properties("IRowsetChange") = True
    cmd.Properties("Updatability") = 7
    
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
   
    If Not (RS.EOF Or RS.BOF) Then
        Set GetAll = RS
        Exit Function
    End If
    
End Function
Public Function Delete() As Variant
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

Public Function Save() As String

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
    
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adNumeric, adParamInput, 5, mvarAdv_Id)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 20, mvarAdv_issue_dt)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adNumeric, adParamInput, 11, mvarAdv_Amt)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adNumeric, adParamInput, 2, mvarNum_Inst)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 150, mvarNotes)
    cmd.Parameters.Append Param6
     
     
    Set Param7 = cmd.CreateParameter("param7", adNumeric, adParamInput, 10, mvarPAID_INSTALLED)
    cmd.Parameters.Append Param7
     
    Set Param8 = cmd.CreateParameter("param8", adNumeric, adParamInput, 10, mvarBalance)
    cmd.Parameters.Append Param8
     
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL Advance_Inf_save( ?, ?, ?, ?, ?,?,?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False



End Function

