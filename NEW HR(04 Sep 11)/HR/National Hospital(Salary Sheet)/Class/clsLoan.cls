VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarEmp_ID As String 'local copy
Private mvarLn_Id As Integer 'local copy
Private mvarLn_Amt As Integer 'local copy
Private mvarIntsl_Amt As Integer 'local copy
Private mvarIntr_Amt As Integer 'local copy
Private mvarIntr_Rate As Integer 'local copy
Private mvarNum_Insl As Integer 'local copy
Private mvarNotes As String 'local copy
Private mvarREF_STATUS As Integer 'local copy
Private mvarConnstring As String 'local copy
'local variable(s) to hold property value(s)
Private mvarSalaryPreparationMonthTo As String 'local copy
Private mvarSalaryPreparationMonthFrom As String 'local copy
Private mvarSalaryPreparationYearTo As String 'local copy
Private mvarSalaryPreparationYearFrom As String 'local copy
Private mvarSalaryPreparationDate As String 'local copy
Public Property Let SalaryPreparationDate(ByVal vData As String)
    mvarSalaryPreparationDate = vData
End Property
Public Property Get SalaryPreparationDate() As String
    SalaryPreparationDate = mvarSalaryPreparationDate
End Property
Public Property Let SalaryPreparationYearFrom(ByVal vData As String)
    mvarSalaryPreparationYearFrom = vData
End Property
Public Property Get SalaryPreparationYearFrom() As String
    SalaryPreparationYearFrom = mvarSalaryPreparationYearFrom
End Property
Public Property Let SalaryPreparationYearTo(ByVal vData As String)
    mvarSalaryPreparationYearTo = vData
End Property
Public Property Get SalaryPreparationYearTo() As String
    SalaryPreparationYearTo = mvarSalaryPreparationYearTo
End Property
Public Property Let SalaryPreparationMonthFrom(ByVal vData As String)
    mvarSalaryPreparationMonthFrom = vData
End Property
Public Property Get SalaryPreparationMonthFrom() As String
    SalaryPreparationMonthFrom = mvarSalaryPreparationMonthFrom
End Property
Public Property Let SalaryPreparationMonthTo(ByVal vData As String)
    mvarSalaryPreparationMonthTo = vData
End Property
Public Property Get SalaryPreparationMonthTo() As String
    SalaryPreparationMonthTo = mvarSalaryPreparationMonthTo
End Property
Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property
'Public Function Save() As String
'
'If mvarLn_Id = Empty Then Exit Function
'
'    Dim Conn As New ADODB.Connection
'    Dim CMD As New ADODB.Command
'    Dim Rs As New ADODB.Recordset
'
'    Dim Param1 As New Parameter
'    Dim Param2 As New Parameter
'    Dim Param3 As New Parameter
'    Dim Param4 As New Parameter
'    Dim Param5 As New Parameter
'    Dim param6 As New Parameter
'    Dim Param7 As New Parameter
'    Dim Param8 As New Parameter
'    Dim param9 As New Parameter
'
'
'    Conn.Open mvarConnString
'
'    Set CMD.ActiveConnection = Conn
'    CMD.CommandType = adCmdText
'
'    '----------------------------------------------------------------------------------
'    Set Param1 = CMD.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_Id)
'    CMD.Parameters.Append Param1
'
'    Set Param2 = CMD.CreateParameter("param2", adInteger, adParamInput, 5, mvarLn_Id)
'    CMD.Parameters.Append Param2
'
'    Set Param3 = CMD.CreateParameter("param3", adInteger, adParamInput, 10, mvarLn_Amt)
'    CMD.Parameters.Append Param3
'
'    Set Param4 = CMD.CreateParameter("param4", adInteger, adParamInput, 8, mvarIntsl_Amt)
'    CMD.Parameters.Append Param4
'
'    Set Param5 = CMD.CreateParameter("param5", adInteger, adParamInput, 8, mvarIntr_Amt)
'    CMD.Parameters.Append Param5
'    Set param6 = CMD.CreateParameter("param6", adVarChar, adParamInput, 4, mvarNum_rate)
'    CMD.Parameters.Append param6
'
'    Set Param7 = CMD.CreateParameter("param7", adInteger, adParamInput, 2, mvarNum_Insl)
'    CMD.Parameters.Append Param7
'
'    Set Param8 = CMD.CreateParameter("param8", adInteger, adParamInput, 150, mvarNotes)
'    CMD.Parameters.Append Param8
'
'    Set param9 = CMD.CreateParameter("param9", adInteger, adParamInput, 1, mvarRef_Status)
'    CMD.Parameters.Append param9
'
'    '----------------------------------------------------------------------------------
'
'' Enable PLSQLRSet property
'    CMD.Properties("PLSQLRSet") = True
'
'    CMD.CommandText = "{CALL Pkg_loan_info.loan_info_IU( ?, ?, ?, ?, ?,?, ?, ?, ?)}"
'
'    Set Rs = CMD.Execute
'
'' Disable PLSQLRSet property
'CMD.Properties("PLSQLRSet") = False
'End Function
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
    Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, mvarLn_Id)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamOutput)
    cmd.Parameters.Append Param2
    '----------------------------------------------------------------------------------
    ' Enable PLSQLRSet property
    
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL Pkg_loan_Info.loan_Info_GetX(?,?)}"

    Set RS = cmd.Execute

    ' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
    
    Set GetX = RS
    
    If Not (RS.EOF And RS.BOF) Then
       mvarEmp_ID = RS("Emp_Id")
       
       mvarLn_Id = RS("Ln_Id")
       
       mvarLn_Amt = RS("Ln_Amt")
       
       mvarIntsl_Amt = RS("Intsl_Amt")
       
       mvarIntr_Amt = RS("Intr_Amt")
       
       mvarIntr_Rate = RS("Intr_Rate ")
       
       mvarNum_Insl = RS("Num_Insl")
       
       mvarNotes = RS("Notes")
       
       mvarREF_STATUS = RS("Ref_Status")
       
    Else
       mvarEmp_ID = ""
       mvarLn_Id = ""
       mvarLn_Amt = ""
       mvarIntsl_Amt = ""
       mvarIntr_Amt = ""
       mvarIntr_Rate = ""
       mvarNum_Insl = ""
       mvarNotes = ""
       mvarREF_STATUS = ""
    End If
    Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Emp_Per_info Table" + vbCr + vbCr + Err.Description
End Function
Public Function Salary_Preparation_For_the_NextMonth_Save()

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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 50, mvarSalaryPreparationYearFrom)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 50, mvarSalaryPreparationMonthFrom)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 50, mvarSalaryPreparationYearTo)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 50, mvarSalaryPreparationMonthTo)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 20, Format(mvarSalaryPreparationDate, "DD-MMM-YYYY"))
    cmd.Parameters.Append Param5
    
   'Format(mvarJobEndingDate, "DD-MMM-YYYY"))
   
   'Format(mvarSalaryPreparationDate, "DD-MMM-YYYY"))
   
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL salary_prepara_For_NeMonth(?,?,?,?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function




Public Function GetAll() As Recordset
End Function
Public Function Delete() As Variant
End Function
Public Property Let REF_STATUS(ByVal vData As Integer)
    mvarREF_STATUS = vData
End Property
Public Property Get REF_STATUS() As Integer
    REF_STATUS = mvarREF_STATUS
End Property
Public Property Let Notes(ByVal vData As String)
    mvarNotes = vData
End Property
Public Property Get Notes() As String
    Notes = mvarNotes
End Property
Public Property Let Num_Insl(ByVal vData As Integer)
    mvarNum_Insl = vData
End Property
Public Property Get Num_Insl() As Integer
    Num_Insl = mvarNum_Insl
End Property
Public Property Let Intr_Rate(ByVal vData As Integer)
    mvarIntr_Rate = vData
End Property
Public Property Get Intr_Rate() As Integer
    Intr_Rate = mvarIntr_Rate
End Property
Public Property Let Intr_Amt(ByVal vData As Integer)
    mvarIntr_Amt = vData
End Property
Public Property Get Intr_Amt() As Integer
    Intr_Amt = mvarIntr_Amt
End Property
Public Property Let Intsl_Amt(ByVal vData As Integer)
    mvarIntsl_Amt = vData
End Property
Public Property Get Intsl_Amt() As Integer
    Intsl_Amt = mvarIntsl_Amt
End Property
Public Property Let Ln_Amt(ByVal vData As Integer)
    mvarLn_Amt = vData
End Property
Public Property Get Ln_Amt() As Integer
    Ln_Amt = mvarLn_Amt
End Property
Public Property Let Ln_Id(ByVal vData As Integer)
    mvarLn_Id = vData
End Property
Public Property Get Ln_Id() As Integer
    Ln_Id = mvarLn_Id
End Property
Public Property Let Emp_ID(ByVal vData As String)
    mvarEmp_ID = vData
End Property
Public Property Get Emp_ID() As String
    Emp_ID = mvarEmp_ID
End Property



