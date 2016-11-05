Attribute VB_Name = "Module1"
'Public rptMode As String
Public Con As New MyConnection
Dim Conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset
Global rptMode As Integer
Public strUid As String
Public strcn        As New MyConnection
Public Booth As String
Public boothN As String
Public IntOption As Integer
Public IntOption1 As Integer
Public Option_discount As Integer
Public optionbuttonval As Integer
Public user_name As Variant
Public PREVIEW_VAR As Long
Public empname As Double
Public Cur_reg_no As String
Public FLED_DATE As Date
Public LockingFlag As Boolean
Public cur_yr_code As String
Public OPD_OUT_INDICATION As String
Public formIndicator As Byte
Public paramMode As Byte
Public paramDate As Date
Public totalCCUCharge As Integer
Public totalCCUDays As Integer
Public canonPrinterName As String
Public PatientStatus As Integer
Public PatientStatusToBe As Integer
Public IRREGULAR_CASE As Integer
Public dischargeType As String
Public admissionDate As Date
Public StaffClass As Integer
Public UserRole As String
Public IsBedDiscountIrregularityOccured As Boolean

Public Function Get_Segment(MyStr As String, portion As Boolean)
    
    '' This function returns a segment of any given string
    '' delimited by a "-"
    '' Developed By Shameem Ferdous on 27 January 2004

    Dim RetStr As String
    Dim l As Integer
    
    MyStr = Trim(MyStr)
    
    If MyStr = Empty Then
        
        Get_Segment = ""
    Else
    
          l = InStr(1, MyStr, "-", vbBinaryCompare)
      
      
        Select Case portion
        
            Case 0  '' Left portion of the delimiter
                
                RetStr = Mid(MyStr, 1, l - 1)
        
            Case 1  '' Right portion of the delimiter
                
                Dim FL As Integer
                
                FL = Len(MyStr)
                
                RetStr = Mid(MyStr, l + 1, FL)
        
        End Select
                
        Get_Segment = RetStr
   
    End If

End Function


Public Sub Locate_Booth()


Dim Boothdata As String
Dim FileNumber As Integer
FileNumber = FreeFile
Open App.Path + "\WT.dat" For Input Access Read As #FileNumber
Input #FileNumber, Boothdata
Close #FileNumber

If Boothdata <> "" Then
    Booth = Boothdata
    boothN = Booth
    frmMAIN.txtBooth = boothN
    
Else
    MsgBox "Set Booth Name in Booth"
    Exit Sub
    End
End If
End Sub

Public Function Get_Code(SString As String) As String
    Get_Code = Trim(Mid(Trim(SString), InStr(Trim(SString), "~") + 1))
End Function
Public Function Get_Description(SString As String) As String
    Get_Description = Trim(Mid(Trim(SString), 1, InStr(Trim(SString), "~") - 1))
End Function
Public Sub BACKUP()
Dim Boothdata As String
Dim FileNumber As Integer
Dim FileNumber1 As Integer

FileNumber = FreeFile
Open App.Path + "C:\Documents and Settings\Administrator\Desktop\8-sep\BILL_98-for general\Project\BACKUP.bat" For Input Access Read As #FileNumber1
'''Input #FileNumber, Boothdata
'''Close #FileNumber

'If Boothdata <> "" Then
'    Booth = Boothdata
'    boothN = Booth
'    frmMAIN.txtBooth = boothN
'
'Else
'    MsgBox "Set Booth Name in Booth"
'    Exit Sub
'    End
'End If
End Sub
