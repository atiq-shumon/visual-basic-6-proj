VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Clear_Screen()

'This Sub routine clears the text box and combo box exists on the calling form
'frm object variable refers to the calling form

On Error Resume Next

    Dim MyObj As Object

    For Each MyObj In Screen.ActiveForm.Controls

        If TypeOf MyObj Is TextBox Or _
                   TypeOf MyObj Is ComboBox Then
                        ''TypeOf MyObj Is MSForms.Label Then
            MyObj = ""
        End If
    Next

End Sub
Public Sub Screen_Position(frm As Form)

' set the calling form in a specified position

On Error Resume Next
    
    Dim Frm_ht As Long
    
    Frm_ht = frm.Height

    If Frm_ht > 6300 Then      'form2,form3
            frm.Move (Screen.Width - frm.Width) \ 2, (Screen.Height - frm.Height) / 1.4
    Else
        frm.Move (Screen.Width - frm.Width) \ 2, (Screen.Height - frm.Height) \ 2
    End If

End Sub
Public Function User_Shift_validation(UserId As String, user_type As String) As Boolean
   Dim Conn As New ADODB.Connection
   Dim RS As New ADODB.Recordset
   Dim cmd As New ADODB.Command
   Dim Server_time As Date

   If Conn.State = 0 Then
      Conn.Open strcn.Connection_String
   End If
      cmd.ActiveConnection = Conn
      cmd.CommandType = adCmdText
      
      Select Case user_type
            Case "Admin"
                 cmd.CommandText = "Select sysdate from dual"
            Case Else
                 cmd.CommandText = "Select shift_name,sysdate from working_schedule where to_number(user_id)=to_number('" & UserId & "') and to_char(dt,'dd/mon/yyyy')=to_char(sysdate,'dd/mon/yyyy')"
      End Select
      
      Set RS = cmd.Execute
      
     
       
      If Not RS.EOF Then
         If Format(RS!sysdate, "HH:MM:SS AM/PM") >= CDate("08:00:00 AM") And Format(RS!sysdate, "HH:MM:SS AM/PM") <= CDate("01:59:59 PM") Then
                 frmMAIN.lblShift.Caption = "Morning"
         ElseIf Format(RS!sysdate, "HH:MM:SS AM/PM") >= CDate("02:00:00 PM") And Format(RS!sysdate, "HH:MM:SS AM/PM") <= CDate("08:59:59 PM") Then
                 frmMAIN.lblShift.Caption = "Evening"
         ElseIf Format(RS!sysdate, "HH:MM:SS AM/PM") >= CDate("09:00:00 PM") And Format(RS!sysdate, "HH:MM:SS AM/PM") <= CDate("11:59:59 PM") Then
                 frmMAIN.lblShift.Caption = "Night"
         ElseIf Format(RS!sysdate, "HH:MM:SS AM/PM") >= CDate("12:00:00 AM") And Format(RS!sysdate, "HH:MM:SS AM/PM") <= CDate("07:59:59 AM") Then
                frmMAIN.lblShift.Caption = "Night"
         End If

    End If
    
    If user_type <> "Admin" Then
         If RS.EOF Then  ''''NO ROWS SELECTED IN THE PREVIOUS CMD COMMAND
            cmd.CommandText = "Select sysdate from dual"
            Set RS = cmd.Execute
         End If
        If Not RS.EOF Then ''for checking if time is after mid night
              If Format(RS!sysdate, "HH:MM:SS AM/PM") >= CDate("12:00:00 AM") And Format(RS!sysdate, "HH:MM:SS AM/PM") <= CDate("07:59:59 AM") Then
                 frmMAIN.lblShift.Caption = "Night"
                 cmd.CommandText = "Select shift_name,sysdate from working_schedule where to_number(user_id)=to_number('" & UserId & "') and to_char(dt,'dd/mon/yyyy')=to_char(sysdate-1,'dd/mon/yyyy')"
                Set RS = cmd.Execute
            End If
        End If
   End If '''end of user_type<>admin
          
   If user_type <> "Admin" Then
   
     User_Shift_validation = False
        If Not RS.EOF Then
            RS.MoveFirst
            Do Until RS.EOF
            If RS(0) = frmMAIN.lblShift Then
                User_Shift_validation = True
                Exit Do
            Else
                User_Shift_validation = False
            End If
            RS.MoveNext
          Loop
       End If
  End If
    
    If user_type = "Admin" Then
       User_Shift_validation = True
    End If
    
End Function
Public Function GetBedChargeForRow(totalRow As Integer, currentRow As Integer, admDate As Date, releaseDate As Date, bedCharge As Integer, admReleaseDateEqualinMoreRows As Boolean, isDepartmentTransferCase As Boolean) As Long
     Dim charge As Long
     charge = 0
If admReleaseDateEqualinMoreRows = False Then
     If totalRow = 1 Then '' no bed or dept transfer
              If admDate = releaseDate And isDepartmentTransferCase = True Then '' release on same date
                   charge = 0
              ElseIf admDate = releaseDate And isDepartmentTransferCase = False Then
                 charge = bedCharge
              Else '' release on other date
                  charge = CalculateDayCountCharge(admDate, releaseDate, bedCharge)
              End If ''' end of no bed or dept transfer
     ElseIf totalRow <> 1 And totalRow <> currentRow Then ''' i.e total row 3 and cur row 1
            If admDate = releaseDate Then
               charge = 0
            Else
               charge = CalculateDayCountCharge(admDate, releaseDate, bedCharge)
            End If
   ElseIf totalRow <> 1 And totalRow = currentRow Then '' bed or dept transfer occured and this is last row
            If admDate = releaseDate Then '' last row but release on same date
                 charge = 0
            Else '' release on other date
                  charge = CalculateDayCountCharge(admDate, releaseDate, bedCharge)
            End If ''' end of no bed or dept transfer
       
   End If
  Else '''admReleaseDateEqualinMoreRows=true
     charge = bedCharge
  End If
     GetBedChargeForRow = charge
End Function
Private Function CalculateDayCountCharge(admDate As Date, releaseDate As Date, bedCharge As Integer) As Long
    Dim dayCount As Integer
    Dim charge As Long
    dayCount = calculate_date(admDate, releaseDate)
    charge = CLng(bedCharge) * CLng(dayCount)
    CalculateDayCountCharge = charge
End Function


Public Function ChkForQuote(str As String)

' When single coat exits in a string (ie. Father's name, Coxe's Bazar )
' the string is trunketed and saved partially in the SQL Server Database
' like Father's name-->  Father ,Coxe's Bazar---> Coxe.
' It happens due to the difference of strig delimiter in VB and in SQL Server
' To overcome this problem a string needs to be processed before sending it to
' the database.

On Error Resume Next

Dim slposl As Integer, Ln As Integer
Dim str2 As String

str2 = ""
For Ln = 1 To Len(str)
    If Mid$(str, Ln, 1) = "'" Then
        str2 = str2 + "'" + Mid(str, Ln, 1)
    Else
        str2 = str2 + Mid$(str, Ln, 1)
    End If
Next
ChkForQuote = str2
End Function
Public Sub Close_Screen_Msg(frm As Form)               ''Close
    Dim Response As String
   Response = MsgBox("Do you really want to close it? ", vbQuestion + vbYesNo + vbDefaultButton2, "Close Screen")
    If Response = vbYes Then
        Unload frm
    Else
        Exit Sub
    End If
End Sub

Public Function Confirm_Delete_Msg(frm As Form) As Boolean
   Dim Response As String
   
   Response = MsgBox("Do you really want to delete it?   ", vbQuestion + vbYesNo, "Confirmation")
        
        If Response = vbNo Then
            
            Confirm_Delete_Msg = False
        
            Exit Function
        Else
            Confirm_Delete_Msg = True
        End If
End Function
Public Function START_END_VALIDATION(Start_date As Date, End_date As Date) As Boolean
  If (End_date - Start_date) < 0 Then
      START_END_VALIDATION = False
  Else
     START_END_VALIDATION = True
  End If
End Function
Public Function calculate_date(Start_date As Date, End_date As Date)
   Dim count_date As Integer
   count_date = (End_date - Start_date)
   calculate_date = count_date
End Function
Public Function Load_Service_Charge_Exceptional(BED As String, DEPARTMENT As String)
     If BED = "Free-Bed" Then
      Select Case DEPARTMENT
             Case "Ophth.", "Gynae-1", "Gynae-2", "Gynae-3", "ENT", "Surgery-1", "Surgery-2", "Surgery-3", "Ortho."
                   Load_Service_Charge_Exceptional = 250
             Case Else
                   Load_Service_Charge_Exceptional = 0
             End Select
       End If
   
End Function
Public Function GetCurrentBed(registrationNo As String, fiscalYear As String)
  Dim Conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim RS As New ADODB.Recordset
  
  If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
    End If
        cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        cmd.CommandText = "select BED_TYPE,Bed_type_no,BED_NO,extra_bed_flag ,SERIAL_NO From Indoor_pat_bed_info Where in_reg_no ='" & Trim(registrationNo) & "' AND YRCODE ='" & Trim(fiscalYear) & "'  AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM Indoor_pat_bed_info WHERE in_reg_no ='" & Trim(registrationNo) & "' AND YRCODE='" & Trim(fiscalYear) & "')"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        RS.CursorLocation = adUseClient

        RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False
          
       If RS.RecordCount > 0 Then
         GetCurrentBed = "" & RS!Bed_type & " -  " & RS!bed_TYPE_no & " -  " & RS!bed_no
      End If
       
       If Conn.State = 1 Then
        Conn.Close
        Set Conn = Nothing
        Set RS = Nothing
        Set cmd = Nothing
     End If
End Function
Public Function GetCurrentDepartment(registrationNo As String, fiscalYear As String)
  Dim Conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim RS As New ADODB.Recordset
  
  If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
  End If
  cmd.ActiveConnection = Conn
  cmd.CommandType = adCmdText
  cmd.CommandText = "SELECT doc_dept,SERIAL_NO FROM INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(registrationNo) & "' AND YRCODE='" & Trim(fiscalYear) & "' AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(registrationNo) & "' AND YRCODE='" & Trim(fiscalYear) & "')"
  
  cmd.Properties("iRowsetChange") = True
  cmd.Properties("updatability") = 7
  RS.CursorLocation = adUseClient

  RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
  cmd.Properties("iRowsetChange") = False

  If RS.RecordCount > 0 Then
     GetCurrentDepartment = "" & RS!doc_dept
     
  End If
  
  Conn.Close
  Set Conn = Nothing
  Set RS = Nothing
  Set cmd = Nothing

End Function
Public Function GetCCUBedInfo(registrationNo As String, fiscalYear As String) As ADODB.Recordset
   Dim Conn As New ADODB.Connection
   Dim cmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   totalCCUCharge = 0
   totalCCUDays = 0
   
   Conn.ConnectionString = strcn.Connection_String
   Conn.Open
   cmd.ActiveConnection = Conn
   cmd.CommandType = adCmdText
   cmd.CommandText = "select start_Date,end_date,bed_charge,DECODE(ROWNUM,1,DECODE((end_date-start_date),0,1,(end_date-start_date)),(end_date-start_date)) Days,bed_charge*DECODE(ROWNUM,1,DECODE((end_date-start_date),0,1,(end_date-start_date)),(end_date-start_date)) AS CHARGE from INDOOR_PAT_CCU_BED_INFO where in_reg_no='" & registrationNo & "' and yrcode='" & fiscalYear & "'"
   cmd.Properties("iRowsetChange") = True
   cmd.Properties("updatability") = 7
   RS.CursorLocation = adUseClient
   RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
   cmd.Properties("iRowsetChange") = False
   
   If RS.RecordCount > 0 Then
       Do Until RS.EOF
'      totalCCUCharge = totalCCUCharge + 0 & RS!charge
      
      If IsNull(RS!End_date) Then
        paramMode = 2
       Else
        paramMode = 1
       End If
       
      RS.MoveNext
      Loop
   Else
       paramMode = 1
   End If
  
   Set GetCCUBedInfo = RS
End Function
Public Function LOAD_STAFF(STAFF_ID As String)
  Dim Conn As New ADODB.Connection
  Dim RS As New ADODB.Recordset
  Dim cmd As New ADODB.Command
  Dim STAFF As String
  
  Conn.ConnectionString = strcn.Connection_String
  Conn.Open
  Set cmd.ActiveConnection = Conn
  cmd.CommandType = adCmdText
  cmd.CommandText = "select A.EMP_NM  AS EMP_name,(SELECT C.DESIGNATION " & _
   " FROM PAYROLL.ST_DESIG C WHERE TO_CHAR(C.DESIG_CODE)= " & _
   " TO_CHAR((SELECT MAX(B.DESIG) FROM PAYROLL.EMP_JOB_INFO B " & _
   " WHERE TO_CHAR(B.EMP_ID)=TO_CHAR('" & STAFF_ID & "') ))) AS DESIGNATION ," & _
   "(select min(emp_class) from  payroll.EMP_JOB_INFO WHERE UPPER(EMP_ID)=TO_CHAR('" & STAFF_ID & "')) class" & _
   " from PAYROLL.EMP_INFO A where A.emp_id='" & STAFF_ID & "'"

  Set RS = cmd.Execute
          
  If Not RS.EOF Then
    STAFF = RS(0) & "   -   " & RS(1)
      StaffClass = 0 & RS(2)
  Else
    STAFF = "0"
  End If
  
  LOAD_STAFF = STAFF
  
  
  Set Conn = Nothing
  Set cmd = Nothing
  Set RS = Nothing
  
End Function
Public Function GetPatientCurrentStatusInStringValue(PATIENT_ID As String, YEARCODE As String) As String
   
   Dim Conn As New ADODB.Connection
   Dim RS As New ADODB.Recordset
   Dim cmd As New ADODB.Command
   
   If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
    End If
    
    cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "select a.release_flag as release_flag,(SELECT MAX(to_char(ADMISSION_DATE,'dd/mm/yyyy')) ADMISSION_DATE FROM INDOOR_PAT_BED_INFO WHERE IN_REG_NO=A.IN_REG_NO AND YRCODE=A.YRCODE) AS CUR_BED_ADM_DATE,TO_CHAR(A.ADMISSION_DATE,'DD/MM/YYYY') AS FirstAdmission, A.cancellation_flag,A.PAT_NAME,A.RELEASE_DATE From in_door_pat_info_main a  Where  a.in_reg_no ='" & Trim(PATIENT_ID) & "'  AND A.YRCODE='" & Trim(YEARCODE) & "'"
      
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    RS.CursorLocation = adUseClient

    RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
    cmd.Properties("iRowsetChange") = False
    
   ''' 0 means admitted ''1 normal release,''2 fled ''3 ACCOUNT HOLD/BACKDATED,4 ADMISSION CANCELLATION
    
    If RS.RecordCount > 0 Then
       admissionDate = RS!CUR_BED_ADM_DATE
       Select Case RS!release_flag
              Case 0
                   GetPatientCurrentStatusInStringValue = "0"
                   PatientStatus = 0 '' admitted
              Case 1
                   GetPatientCurrentStatusInStringValue = "Patient Name :    " & UCase(RS!pat_name) & vbCrLf & vbCrLf & "Status            :  RELEASED" & vbCrLf & vbCrLf & "On date          : " & RS!release_date
                   PatientStatus = 1 ''released
              Case 2
                   GetPatientCurrentStatusInStringValue = "Patient Name :    " & UCase(RS!pat_name) & vbCrLf & vbCrLf & "Status            :  ABSCONDED" & vbCrLf & vbCrLf & "On date         : " & RS!release_date
                   PatientStatus = 2 ''''absconded
              Case 3
                   GetPatientCurrentStatusInStringValue = "Patient Name :    " & UCase(RS!pat_name) & vbCrLf & vbCrLf & "Status             :  HOLD/BACKED PATIENT" & vbCrLf & vbCrLf & "On date       : " & RS!release_date
                   PatientStatus = 3 '''hold
              Case 4
                   GetPatientCurrentStatusInStringValue = "Patient Name :    " & UCase(RS!pat_name) & vbCrLf & vbCrLf & "Status            :  ADMISSION CANCELLED" & vbCrLf & vbCrLf & "On date          : " & RS!release_date
                   PatientStatus = 4 '''cancelled
              
         End Select
   Else
         PatientStatus = -1
         GetPatientCurrentStatusInStringValue = "You have Entered Wrong Registration No"
        
   End If
         
    Set RS = Nothing
    If Conn.State = 1 Then
       Conn.Close
      Set Conn = Nothing
    End If
End Function
Public Function IsNum(ByVal KeyAscii As Integer)
 'it validates number during input

     Select Case KeyAscii
        Case 8
            IsNum = 8               ' if back space then back space
        Case 13
            IsNum = 13              ' if enter then enter
        Case 46
            IsNum = 46              ' if decimal point then decimal point
        Case Is < 48, Is > 57       ' if non numeral then nothing
            IsNum = 0
        Case Is >= 48, Is <= 57     ' if numeral then numeral
            IsNum = KeyAscii


    End Select

End Function
Public Sub SAVE_DELETE(MODE As Integer, deptCode As String, mCode As String, sCode As String, mNAME As String, sNAME As String, testType As String, charge As Integer, Booth As String)
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    Dim Param0 As New Parameter
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
   
   
    If Conn.State = 0 Then
        Conn.Open strcn.Connection_String
    End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    Set Param0 = cmd.CreateParameter("param0", adSmallInt, adParamInput, 5, MODE)
    cmd.Parameters.Append Param0 'mode
    
     Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, deptCode)
    cmd.Parameters.Append Param1 'DEPT_code
   
           
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 5, mCode)
    cmd.Parameters.Append Param2 'm_code
            
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 5, sCode)
    cmd.Parameters.Append Param3 's_code
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 60, mNAME)
    cmd.Parameters.Append Param4 'mName
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 160, sNAME)
    cmd.Parameters.Append Param5 's_name
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 10, testType)
     cmd.Parameters.Append Param6 'type
    
   Set Param7 = cmd.CreateParameter("param7", adInteger, adParamInput, 15, charge)
    cmd.Parameters.Append Param7 'charge
    
 
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 10, Booth)
    cmd.Parameters.Append Param8 'BOOTH
            
               
    cmd.Properties("PLSQLRSet") = True
            
   cmd.CommandText = "{CALL save_DELETE_temp_tab(?,?,?,?,?,?,?,?,?)}"
              
   Set RS = cmd.Execute
   cmd.Properties("PLSQLRSet") = False
   
   Set Conn = Nothing
   Set cmd = Nothing
   Set RS = Nothing
    
End Sub

Public Function GetFiscalYears() As String()
   Dim arr(0 To 5) As String
   arr(0) = "YR-1011"
   arr(1) = "YR-0910"
   arr(2) = "YR-0809"
   arr(3) = "YR-0708"
   arr(4) = "YR-0406"
   
'   arr(0) = "YR-1112"
'   arr(1) = "YR-1011"
'   arr(2) = "YR-0910"
'   arr(3) = "YR-0809"
'   arr(4) = "YR-0708"
'   arr(5) = "YR-0406"

   GetFiscalYears = arr
End Function

Public Function GetPatientStatus() As String()
   Dim arr(0 To 4) As String
   
   arr(0) = "Admitted"
   arr(1) = "Released"
   arr(2) = "Absconded"
   arr(3) = "Hold"
   arr(4) = "Cancelled"
   GetPatientStatus = arr
End Function
Public Function GetPatientStatusNo(statusString As String) As Integer
 Dim i As Integer
    Select Case statusString
           Case "Admitted"
                 i = 0
           Case "Released"
                 i = 1
           Case "Absconded"
                 i = 2
           Case "Hold"
                 i = 3
           Case "Cancelled"
                 i = 4
    End Select
    GetPatientStatusNo = i
End Function
Public Function GetBedType() As String()
  Dim arr(0 To 2) As String
  arr(0) = "Cabin"
  arr(1) = "Paying"
  arr(2) = "Free-Bed"
  GetBedType = arr
End Function
Public Function GetDischargeTypes() As String()
  Dim arr(0 To 3) As String
  arr(0) = "Normal"
  arr(1) = "DORB"
  arr(2) = "Death"
  arr(3) = "Refd.To"
  GetDischargeTypes = arr
End Function

Public Function GetDischargeTypeInCode(dischargeType As String) As String
  Select Case dischargeType
         Case "Normal"
               GetDischargeTypeInCode = "N"
          Case "DORB" ''request
               GetDischargeTypeInCode = "R"
          Case "Death"
              GetDischargeTypeInCode = "D"
           Case "Refd.To"
              GetDischargeTypeInCode = "T"
  End Select


End Function
Public Function IsAdmissionDateLess(admissionDate As Date, endDate As Date) As Boolean
      If admissionDate > endDate Then
         IsAdmissionDateLess = True
      Else
        IsAdmissionDateLess = False
      End If

End Function




Public Function GetDischargeTypeInString(dischargeType As String) As String
  Select Case dischargeType
         Case "N"
               GetDischargeTypeInString = "Normal"
          Case "R" ''' request
               GetDischargeTypeInString = "DORB"
          Case "D"
              GetDischargeTypeInString = "Death"
         Case "T"
              GetDischargeTypeInString = "Refd.To"
  End Select


End Function
Public Function IsEntitledToStayInCabin(StaffClass As Integer) As Boolean
       Select Case StaffClass
              Case 1
                  IsEntitledToStayInCabin = True
              Case Else
                  IsEntitledToStayInCabin = False
             End Select

End Function
