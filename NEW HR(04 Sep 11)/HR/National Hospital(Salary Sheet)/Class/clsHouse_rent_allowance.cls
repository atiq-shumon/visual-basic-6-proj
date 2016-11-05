VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHouse_rent_allowance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarSlab_code As String 'local copy
Private mvarBasic_From As Integer 'local copy
Private mvarBasic_To As Integer 'local copy
Private mvarRate As Integer 'local copy
Private mvarMinimum As Integer 'local copy
Private mvarConnstring As String 'local copy
Public Function Save() As String

'On ERROR Resume Next

If mvarSlab_code = Empty Then Exit Function

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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, mvarSlab_code)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adInteger, adParamInput, 50, mvarBasic_From)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adInteger, adParamInput, 200, mvarBasic_To)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adInteger, adParamInput, 50, mvarRate)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 50, mvarMinimum)
    cmd.Parameters.Append Param5
    
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Pkg_St_Company.St_HouseRent_Allowance_IU( ?, ?, ?, ?, ?)}"
    
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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarSlab_code)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamOutput)
    cmd.Parameters.Append Param2
    '----------------------------------------------------------------------------------
    ' Enable PLSQLRSet property
    
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL Pkg_St_Company.St_HouseRent_Allowance_GetX(?,?)}"

    Set RS = cmd.Execute

    ' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
    
    Set GetX = RS
    
    If Not (RS.EOF And RS.BOF) Then
       mvarSlab_code = RS("Slab_code")
       
       mvarBasic_From = RS("Basic_From")
       
       mvarBasic_To = RS("Basic_To")
       
       mvarRate = RS("Rate")
       
       mvarMinimum = RS("Minimum")
       
       
       
       
    Else
       mvarSlab_code = 0
       
       mvarBasic_From = 0
       
       mvarBasic_To = 0
       
       mvarRate = 0
       
       mvarMinimum = 0
        End If
    Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Emp_Per_info Table" + vbCr + vbCr + Err.Description



End Function

Public Function GetAll() As Recordset

   On Error Resume Next


    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset
    
    
    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn

    cmd.CommandType = adCmdText
    cmd.CommandText = "Select * from St_Houserent_Allowance"
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

On Error GoTo DeleteError
Dim conn As New Connection
Dim cmd As New Command

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Emp_Per_info_Del"
    cmd(1) = mvarSlab_code
    
    cmd.Execute
    Exit Function
DeleteError:
    Err.Raise vbObjectError + 513, "Delete Error", "Problems in Deleting Data from the Emp_info Table" + vbCr + vbCr + Err.Description
End Function

Public Property Let Connstring(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ConnString = 5
    mvarConnstring = vData
End Property


Public Property Get Connstring() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ConnString
    Connstring = mvarConnstring
End Property



Public Property Let Minimum(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Minimum = 5
    mvarMinimum = vData
End Property


Public Property Get Minimum() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Minimum
    Minimum = mvarMinimum
End Property



Public Property Let Rate(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Rate = 5
    mvarRate = vData
End Property


Public Property Get Rate() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Rate
    Rate = mvarRate
End Property



Public Property Let Basic_To(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Basic_To = 5
    mvarBasic_To = vData
End Property


Public Property Get Basic_To() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Basic_To
    Basic_To = mvarBasic_To
End Property



Public Property Let Basic_From(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Basic_From = 5
    mvarBasic_From = vData
End Property


Public Property Get Basic_From() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Basic_From
    Basic_From = mvarBasic_From
End Property



Public Property Let Slab_code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Slab_code = 5
    mvarSlab_code = vData
End Property


Public Property Get Slab_code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Slab_code
    Slab_code = mvarSlab_code
End Property



