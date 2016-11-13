Attribute VB_Name = "modOtherFunctions"
Public CompanyName, Address As String
Global Myrecord As New ADODB.Recordset

Public Function Check_ValidDate(InitialDate As String) As Boolean
    Dim Day1 As Integer
    Dim Month1 As Integer
    Dim Year1 As Integer
    Dim IDate As Integer
    
    If IsDate(InitialDate) = False Then
        MsgBox "Invalid Date.", vbInformation, "Daffodil Software"
        Check_ValidDate = False
        Exit Function
    End If
    
    Day1 = Mid(InitialDate, 1, 2)
    Month1 = Mid(InitialDate, 4, 2)
    Year1 = Mid(InitialDate, 7, 2)
    If Year1 < 50 Then
        IDate = "20" + Format(Year1, "00")
    Else
        IDate = "19" + Format(Year1, "00")
    End If
    
    Dim Month As Integer
    Dim Day As Integer
    Dim Year As Integer
    
    Month = Month1
    Day = Day1
    Year = IDate
    
    If Month = 4 Or Month = 6 Or Month = 9 Or Month = 11 Then
        If Day > 30 Then
            MsgBox "Invalid day format.", vbInformation, "Daffodil Software"
            Check_ValidDate = False
        Else
            Check_ValidDate = True
        End If
    ElseIf Month = 1 Or Month = 3 Or Month = 5 Or Month = 7 Or Month = 8 Or Month = 10 Or Month = 12 Then
        If Day > 31 Then
            MsgBox "Invalid day format.", vbInformation, "Daffodil Software"
            Check_ValidDate = False
        Else
            Check_ValidDate = True
        End If
    ElseIf Month = 2 Then
            If Year Mod 4 = 0 Then
                    If Day > 29 Then
                        MsgBox "Invalid day Format.", vbInformation, "Daffodil Software"
                        Check_ValidDate = False
                    Else
                        Check_ValidDate = True
                    End If
            Else
                If Day > 28 Then
                    MsgBox "Invalid Day Format.", vbInformation, "Daffodil Software"
                    Check_ValidDate = False
                Else
                    Check_ValidDate = True
                End If
            End If
    ElseIf Month > 12 Then
            MsgBox "Invalid Month Format.", vbInformation, "Daffodil Software"
            Check_ValidDate = False
    Else
            Check_ValidDate = True
    End If
End Function

Public Sub Company()

'Dim Myrecord As New adodb.Recordset
'Dim CompanyName, Address As String

Set Myrecord = getdata("SELECT CompanyName, CompanyAddress " + _
 "FROM dbo.CompanyInfo")
            
 If Not Myrecord.EOF Or Myrecord.BOF Then
    CompanyName = Myrecord.Fields(0)
    Address = Myrecord.Fields(1)
 Else
    CompanyName = "Not Avialable"
    Address = "Not Avialable"
 End If

End Sub

