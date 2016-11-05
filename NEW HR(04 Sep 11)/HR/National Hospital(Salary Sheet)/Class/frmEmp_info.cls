VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Emp_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarConnstring As String
Private mvarEmp_ID As String
Private mvarEmp_Nm As String
Private mvarEmp_Fat_Nm As String
Private mvarEmp_Mat_Nm As String
Private mvarCode_No As String
Private mvarDOB As Date
Private mvarReligion As String
Private mvarAddress_Perm As String
Private mvarDistrict_Perm As String
Private mvarPS_Perm As String
Private mvarPost_Perm As String
Private mvarAddress_Pres As String
Private mvarDistrict_Pres As String
Private mvarPS_Pres As String
Private mvarPost_Pres As String
Private mvarMarital_Stat As String
Private mvarGender As String
Private mvarPhoto As String
Private mvarU_Id As String
Private mvarNationality As String 'local copy
Private mvarCountry_Pres As String 'local copy
Private mvarCountry_Perm As String 'local copy
Private mvarTELEPHONE As String 'local copy
Private mvarCellphone As String 'local copy
Private mvarE_mail As String 'local copy
Public Property Let E_mail(ByVal vData As String)
    mvarE_mail = vData
End Property
Public Property Get E_mail() As String
    E_mail = mvarE_mail
End Property
Public Property Let Cellphone(ByVal vData As String)
    mvarCellphone = vData
End Property
Public Property Get Cellphone() As String
    Cellphone = mvarCellphone
End Property
Public Property Let TELEPHONE(ByVal vData As String)
    mvarTELEPHONE = vData
End Property
Public Property Get TELEPHONE() As String
    TELEPHONE = mvarTELEPHONE
End Property
Public Property Let Country_Perm(ByVal vData As String)
    mvarCountry_Perm = vData
End Property
Public Property Get Country_Perm() As String
    Country_Perm = mvarCountry_Perm
End Property
Public Property Let Country_Pres(ByVal vData As String)
    mvarCountry_Pres = vData
End Property
Public Property Get Country_Pres() As String
    Country_Pres = mvarCountry_Pres
End Property



Public Property Let Nationality(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Nationality = 5
    mvarNationality = vData
End Property


Public Property Get Nationality() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Nationality
    Nationality = mvarNationality
End Property



Public Property Let Emp_ID(ByVal vData As String)
   mvarEmp_ID = vData
End Property
Public Property Get Emp_ID() As String
   Emp_ID = mvarEmp_ID
End Property
Public Property Let Emp_Nm(ByVal vData As String)
   mvarEmp_Nm = vData
End Property
Public Property Get Emp_Nm() As String
   Emp_Nm = mvarEmp_Nm
End Property
Public Property Let Emp_Fat_Nm(ByVal vData As String)
   mvarEmp_Fat_Nm = vData
End Property
Public Property Get Emp_Fat_Nm() As String
   Emp_Fat_Nm = mvarEmp_Fat_Nm
End Property
Public Property Let Emp_Mat_Nm(ByVal vData As String)
   mvarEmp_Mat_Nm = vData
End Property
Public Property Get Emp_Mat_Nm() As String
   Emp_Mat_Nm = mvarEmp_Mat_Nm
End Property
Public Property Let Code_No(ByVal vData As String)
   mvarCode_No = vData
End Property
Public Property Get Code_No() As String
   Code_No = mvarCode_No
End Property
Public Property Let DOB(ByVal vData As Date)
   mvarDOB = vData
End Property
Public Property Get DOB() As Date
   DOB = mvarDOB
End Property
Public Property Let Religion(ByVal vData As String)
   mvarReligion = vData
End Property
Public Property Get Religion() As String
   Religion = mvarReligion
End Property

Public Property Let Address_Perm(ByVal vData As String)
   mvarAddress_Perm = vData
End Property
Public Property Get Address_Perm() As String
   Address_Perm = mvarAddress_Perm
End Property
Public Property Let District_Perm(ByVal vData As String)
   mvarDistrict_Perm = vData
End Property
Public Property Get District_Perm() As String
   District_Perm = mvarDistrict_Perm
End Property
Public Property Let PS_Perm(ByVal vData As String)
   mvarPS_Perm = vData
End Property
Public Property Get PS_Perm() As String
   PS_Perm = mvarPS_Perm
End Property
Public Property Let Post_Perm(ByVal vData As String)
   mvarPost_Perm = vData
End Property
Public Property Get Post_Perm() As String
   Post_Perm = mvarPost_Perm
End Property
Public Property Let Address_Pres(ByVal vData As String)
   mvarAddress_Pres = vData
End Property
Public Property Get Address_Pres() As String
   Address_Pres = mvarAddress_Pres
End Property
Public Property Let District_Pres(ByVal vData As String)
   mvarDistrict_Pres = vData
End Property
Public Property Get District_Pres() As String
   District_Pres = mvarDistrict_Pres
End Property
Public Property Let PS_Pres(ByVal vData As String)
   mvarPS_Pres = vData
End Property
Public Property Get PS_Pres() As String
   PS_Pres = mvarPS_Pres
End Property
Public Property Let Post_Pres(ByVal vData As String)
   mvarPost_Pres = vData
End Property
Public Property Get Post_Pres() As String
   Post_Pres = mvarPost_Pres
End Property

Public Property Let Marital_Stat(ByVal vData As String)
   mvarMarital_Stat = vData
End Property
Public Property Get Marital_Stat() As String
   Marital_Stat = mvarMarital_Stat
End Property
Public Property Let Gender(ByVal vData As String)
   mvarGender = vData
End Property
Public Property Get Gender() As String
   Gender = mvarGender
End Property
Public Property Let Photo(ByVal vData As String)
   mvarPhoto = vData
End Property
Public Property Get Photo() As String
   Photo = mvarPhoto
End Property
Public Property Let U_Id(ByVal vData As String)
   mvarU_Id = vData
End Property
Public Property Get U_Id() As String
   U_Id = mvarU_Id
End Property

Public Property Let Connstring(ByVal vData As String)
   mvarConnstring = vData
End Property
Public Property Get Connstring() As String
   Connstring = mvarConnstring
End Property
Public Function Save()
   If mvarEmp_ID = Empty Then Exit Function

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
    Dim Param12 As New ADODB.Parameter
    Dim Param13 As New ADODB.Parameter
    Dim Param14 As New ADODB.Parameter
    Dim Param15 As New ADODB.Parameter
    Dim Param16 As New ADODB.Parameter
    Dim Param17 As New ADODB.Parameter
    Dim Param18 As New ADODB.Parameter
    Dim Param19 As New ADODB.Parameter
    Dim Param20 As New ADODB.Parameter
    Dim Param21 As New ADODB.Parameter
    Dim Param22 As New ADODB.Parameter
    Dim Param23 As New ADODB.Parameter
    
    
        
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 50, mvarEmp_Nm)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 50, mvarEmp_Fat_Nm)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 50, mvarEmp_Mat_Nm)
    cmd.Parameters.Append Param4

    Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, , mvarDOB)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 1, mvarReligion)
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 1, mvarGender)
    cmd.Parameters.Append Param7

    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 1, mvarNationality)
    cmd.Parameters.Append Param8
    
    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 100, mvarAddress_Perm)
    cmd.Parameters.Append Param9
    
    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 50, mvarPS_Perm)
    cmd.Parameters.Append Param10

    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 50, mvarPost_Perm)
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adVarChar, adParamInput, 100, mvarDistrict_Perm)
    cmd.Parameters.Append Param12
    
    Set Param13 = cmd.CreateParameter("param13", adVarChar, adParamInput, 50, mvarCountry_Perm)
    cmd.Parameters.Append Param13
    
    Set Param14 = cmd.CreateParameter("param14", adVarChar, adParamInput, 100, mvarAddress_Pres)
    cmd.Parameters.Append Param14
    
    Set Param15 = cmd.CreateParameter("param15", adVarChar, adParamInput, 50, mvarPS_Pres)
    cmd.Parameters.Append Param15
    
    Set Param16 = cmd.CreateParameter("param16", adVarChar, adParamInput, 100, mvarPost_Pres)
    cmd.Parameters.Append Param16
    
    Set Param17 = cmd.CreateParameter("param17", adVarChar, adParamInput, 50, mvarDistrict_Pres)
    cmd.Parameters.Append Param17

    Set Param18 = cmd.CreateParameter("param18", adVarChar, adParamInput, 50, mvarCountry_Pres)
    cmd.Parameters.Append Param18
    
    Set Param19 = cmd.CreateParameter("param19", adVarChar, adParamInput, 100, mvarTELEPHONE)
    cmd.Parameters.Append Param19
    
    Set Param20 = cmd.CreateParameter("param20", adVarChar, adParamInput, 25, mvarCellphone)
    cmd.Parameters.Append Param20

    Set Param21 = cmd.CreateParameter("param21", adVarChar, adParamInput, 50, mvarE_mail)
    cmd.Parameters.Append Param21
    
    Set Param22 = cmd.CreateParameter("param22", adVarChar, adParamInput, 25, mvarCode_No)
    cmd.Parameters.Append Param22
    
    Set Param23 = cmd.CreateParameter("param23", adVarChar, adParamInput, 1, mvarMarital_Stat)
    cmd.Parameters.Append Param23
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Pkg_Emp_Info.Emp_Info_IU(?, ?, ?,?,?, ?, ?,?,?, ?, ?,?,?, ?, ?,?,?, ?, ?,?,?, ?, ?)}"
    
    Set Rs = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function
Public Function Delete()
On Error GoTo DeleteError
Dim conn As New Connection
Dim cmd As New Command

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Emp_Per_info_Del"
    cmd(1) = mvarEmp_ID
    cmd.Execute
    Exit Function
DeleteError:
    Err.Raise vbObjectError + 513, "Delete Error", "Problems in Deleting Data from the Emp_info Table" + vbCr + vbCr + Err.Description
End Function

Public Function GetX() As Recordset

  On Error Resume Next

    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim Rs As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    
    ''Here Param1 is declared as Parameter type variable for taking input parameter
    '' and Param2 is for retrieving records fetched by the command text.
    
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_ID)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamOutput)
   cmd.Parameters.Append Param2
'    '----------------------------------------------------------------------------------
    ' Enable PLSQLRSet property
    
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL PKG_Emp_Info.Emp_Info_GetX(?,?)}"

    Set Rs = cmd.Execute

    ' Disable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = False
    
    Set GetX = Rs
    
    If Not (Rs.EOF And Rs.BOF) Then
       mvarEmp_ID = Rs("Emp_Id")
       mvarEmp_Nm = Rs("EMP_NM")
       mvarEmp_Fat_Nm = Rs("EMP_FAT_NM")
       mvarEmp_Mat_Nm = Rs("Emp_Mat_Nm")
       mvarDOB = Rs("DOB")
       mvarReligion = Rs("Relegion")
       mvarGender = Rs("Gender")
       mvarNationality = Rs("Nationality")
       mvarAddress_Perm = Rs("C_ADDRESS")
       mvarPS_Perm = Rs("P_PS")
       mvarPost_Perm = Rs("P_Post")
       mvarDistrict_Perm = Rs("P_Dist")
       mvarCountry_Perm = Rs("P_Country")
       mvarAddress_Pres = Rs("P_Address")
       mvarPS_Pres = Rs("C_PS")
       mvarPost_Pres = Rs("C_Post")
       mvarDistrict_Pres = Rs("C_Dist")
       mvarCountry_Pres = Rs("C_Country")

       mvarTELEPHONE = Rs("Telephone")
       mvarCellphone = Rs("Cellphone")
       mvarE_mail = Rs("Email")
       mvarMarital_Stat = Rs("Marital_Status")

            
    Else
       mvarEmp_ID = ""
       mvarEmp_Nm = ""
       mvarEmp_Fat_Nm = ""
       mvarEmp_Mat_Nm = ""
       mvarDOB = ""
       mvarReligion = ""
       mvarGender = ""
       mvarNationality = ""
       mvarAddress_Perm = ""
       mvarPS_Perm = ""
       mvarPost_Perm = ""
       mvarDistrict_Perm = ""
       mvarCountry_Perm = ""
       mvarAddress_Pres = ""
       mvarPS_Pres = ""
       mvarPost_Pres = ""
       mvarDistrict_Pres = ""
       mvarCountry_Pres = ""
       
       mvarTELEPHONE = ""
       mvarCellphone = ""
       mvarE_mail = ""
       mvarCode_No = ""
       mvarMarital_Stat = ""
       
       
       'mvarPhoto = ""
      ' mvarU_Id = ""
       
       
    End If
    Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Emp_info Table" + vbCr + vbCr + Err.Description
End Function
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
'Public Function GetX_For_RetiredmentInformation() As Recordset
'
'  On Error Resume Next
'
'    Dim Conn As New ADODB.Connection
'    Dim cmd As New ADODB.Command
'    Dim Rs As New ADODB.Recordset
'
'    Dim Param1 As New Parameter
'    Dim Param2 As New Parameter
'
'    ''Here Param1 is declared as Parameter type variable for taking input parameter
'    '' and Param2 is for retrieving records fetched by the command text.
'
'    Conn.Open mvarConnString
'
'    Set cmd.ActiveConnection = Conn
'    cmd.CommandType = adCmdText
'
'    '----------------------------------------------------------------------------------
'    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mvarEmp_Id)
'    cmd.Parameters.Append Param1
'
'    Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamOutput)
'    cmd.Parameters.Append Param2
'    '----------------------------------------------------------------------------------
'    ' Enable PLSQLRSet property
'
'    cmd.Properties("PLSQLRSet") = True
'    cmd.CommandText = "{CALL PKG_Emp_Info.Emp_Info_GetX(?,?)}"
'
'    Set Rs = cmd.Execute
'
'    ' Disable PLSQLRSet property
'    cmd.Properties("PLSQLRSet") = False
'
'    Set GetX = Rs
'
'    If Not (Rs.EOF And Rs.BOF) Then
'       mvarEmp_Id = Rs("Emp_Id")
'       mvarEmp_Nm = Rs("Emp_Nm")
'       mvarEmp_Fat_Nm = Rs("EMP_FAT_NM")
'       mvarEmp_Mat_Nm = Rs("Emp_Mat_Nm")
'       mvarDOB = Rs("DOB")
'       mvarReligion = Rs("Relegion")
'       mvarGender = Rs("Gender")
'       mvarNationality = Rs("Nationality")
'       mvarAddress_Perm = Rs("P_Address")
'       mvarPS_Perm = Rs("P_PS")
'       mvarPost_Perm = Rs("P_Post")
'       mvarDistrict_Perm = Rs("P_Dist")
'       mvarCountry_Perm = Rs("P_Country")
'       mvarAddress_Pres = Rs("P_Address")
'       mvarPS_Pres = Rs("C_PS")
'       mvarPost_Pres = Rs("C_Post")
'       mvarDistrict_Pres = Rs("C_Dist")
'       mvarCountry_Pres = Rs("C_Country")
'
'       mvarTELEPHONE = Rs("Telephone")
'       mvarCellphone = Rs("Cellphone")
'       mvarE_mail = Rs("Email")
'       mvarMarital_Stat = Rs("Marital_Status")
'
'
'    Else
'       mvarEmp_Id = ""
'       mvarEmp_Nm = ""
'       mvarEmp_Fat_Nm = ""
'       mvarEmp_Mat_Nm = ""
'       mvarDOB = ""
'       mvarReligion = ""
'       mvarGender = ""
'       mvarNationality = ""
'       mvarAddress_Perm = ""
'       mvarPS_Perm = ""
'       mvarPost_Perm = ""
'       mvarDistrict_Perm = ""
'       mvarCountry_Perm = ""
'       mvarAddress_Pres = ""
'       mvarPS_Pres = ""
'       mvarPost_Pres = ""
'       mvarDistrict_Pres = ""
'       mvarCountry_Pres = ""
'
'       mvarTELEPHONE = ""
'       mvarCellphone = ""
'       mvarE_mail = ""
'       mvarCode_No = ""
'       mvarMarital_Stat = ""
'
'
'    End If
'    Exit Function
'GetXError:
'    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Emp_info Table" + vbCr + vbCr + Err.Description
'End Function
'
'
'
