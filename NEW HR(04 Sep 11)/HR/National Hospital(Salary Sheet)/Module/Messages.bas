Attribute VB_Name = "Msg"
Public Sub Close_Msg(Form_Num As Form)               ''Close
   yes_no = MsgBox("Do you really want to close it? ", vbQuestion + vbYesNo + vbDefaultButton2, "Close Screen")
    If yes_no = vbYes Then Unload Form_Num
    If yes_no = vbNo Then Exit Sub
End Sub

Public Sub Confirm_Save_Msg(Form_Num As Form)                                ''Exit
   
   yes_no = MsgBox("Do you want to save it?   ", vbQuestion + vbYesNo, "Confirmation")
        If yes_no = vbNo Then
            Form_Num.Command1(2).SetFocus
            Exit Sub
        End If
    If yes_no = vbYes Then
         Form_Num.Command1(0).SetFocus      ''Set Focus to Save Button
    End If
End Sub
Public Function ConvertX(nbrNumber As Double) As String
    Dim intDECI As Integer
    Dim intVAL As Integer
    Dim strNUM As String
    Dim strDECI As String
    Dim strHole As String
    Dim strWord As String
    
    intDECI = DecimalX(CStr(nbrNumber))
    
    If intDECI > 0 Then
       strNUM = Left$(nbrNumber, intDECI - 1)
       
        strDECI = Mid$(nbrNumber, intDECI + 1, 2)
        Dim L As String
        Dim R As String
        Dim strFlot As String
     '   Dim strFlot1 As String
        
        L = Left$(strDECI, 1)
        R = Mid$(strDECI, 2, 1)
             
        If L = "0" And R <> "0" Then
           strFlot = "  AND PAISA ZERO " & Word1(R)
           GoTo stp
        End If
        
        If L <> "0" And R = "0" Then
           strFlot = " AND PAISA " & Word1(L) & " ZERO "
           GoTo stp
        End If
        
        If L <> "" And R = "" Then
           strFlot = " AND PAISA " & Word1(L) & " ZERO "
        End If
        
        If L <> "" And R <> "" Then
           strFlot = " AND PAISA " & Word1(L) & "  " & Word1(R)
           GoTo stp
        End If
        
stp:
    Else
       strNUM = nbrNumber
    End If
    
    intVAL = Len(strNUM)
    
    Select Case intVAL
        Case 1
            strHole = Word1(strNUM)
        Case 2
            strHole = Word2(strNUM)
        Case 3
            strHole = Word3(strNUM)
        Case 4
            strHole = Word4(strNUM)
        Case 5
            strHole = Word5(strNUM)
        Case 6
            strHole = Word6(strNUM)
        Case 7
            strHole = Word7(strNUM)
        Case 8
            strHole = Word8(strNUM)
        Case 9
            strHole = Word9(strNUM)
        Case 10
            strHole = Word10(strNUM)
        Case Else
            MsgBox "Supports 10 digit with two decimal places", vbInformation
            Exit Function
    End Select
    
    If strFlot = "" Then
       ConvertX = strHole
    Else
       ConvertX = strHole & strFlot
    End If
    
    'MsgBox " " & nbrNumber.Text & " " & " In words( " & " " & strWord & " )"
End Function
Public Function Word1(intNum1 As String) As String
    If intNum1 <> "0" Then
       Select Case intNum1
       Case "1"
            Word1 = "ONE"
       Case "2"
            Word1 = "TWO"
       Case "3"
            Word1 = "THREE"
       Case "4"
            Word1 = "FOUR"
       Case "5"
            Word1 = "FIVE"
       Case "6"
            Word1 = "SIX"
       Case "7"
            Word1 = "SEVEN"
       Case "8"
            Word1 = "EIGHT"
       Case "9"
            Word1 = "NINE"
       End Select
    Else
        Word1 = ""
    End If
End Function
Public Function Word2(intNum2 As String) As String
    Dim f As String
    Dim L As String
    
    f = Left$(intNum2, 1)
    L = Right$(intNum2, 1)
    
    If f = "0" Then
       Word2 = Word1(L)
       Exit Function
    Else
       If Val(f) >= 2 Then
          Select Case f
          Case "2"
            Word2 = "TWENTY  " & Word1(L)
          Case "3"
            Word2 = "THIRTY  " & Word1(L)
          Case "4"
            Word2 = "FORTY  " & Word1(L)
          Case "5"
            Word2 = "FIFTY  " & Word1(L)
          Case "6"
            Word2 = "SIXTY  " & Word1(L)
          Case "7"
            Word2 = "SEVENTY  " & Word1(L)
          Case "8"
            Word2 = "EIGHTY  " & Word1(L)
          Case "9"
            Word2 = "NINTY  " & Word1(L)
          End Select
       Else
          Select Case intNum2
          Case "10"
            Word2 = "TEN"
          Case "11"
            Word2 = "ELEVEN"
          Case "12"
            Word2 = "TWELVE"
          Case "13"
            Word2 = "THIRTEEN"
          Case "14"
            Word2 = "FOURTEEN"
          Case "15"
            Word2 = "FIFTEEN"
          Case "16"
            Word2 = "SIXTEEN"
          Case "17"
            Word2 = "SEVENTEEN"
          Case "18"
            Word2 = "EIGHTEEN"
          Case "19"
            Word2 = "NINETEEN"
          End Select
       End If
    End If
End Function

Public Function Word3(intNUM3 As String) As String
    Dim f As String
    Dim L As String
    
    f = Left$(intNUM3, 1)
    L = Right$(intNUM3, 2)
    
    If f = "0" Then
        Word3 = Word2(L)
    Else
        Word3 = Word1(f) & " HUNDRED " & Word2(L)
    End If
End Function

Public Function Word4(intNUM4 As String) As String
    Dim f As String
    Dim L As String
    
    f = Left$(intNUM4, 1)
    L = Right$(intNUM4, 3)
    
    If f = "0" Then
        Word4 = Word3(L)
    Else
        Word4 = Word1(f) & " THOUSAND " & Word3(L)
    End If
End Function
Public Function Word5(intNUM5 As String) As String
    Dim LF As String
    Dim RT As String
    Dim LS As String
    Dim RF As String
    
    LF = Left$(intNUM5, 1)
    LS = Left$(intNUM5, 2)
    RT = Right$(intNUM5, 3)
    RF = Right$(intNUM5, 4)
    
    If LF = "0" Then
        Word5 = Word4(RF)
    Else
        Word5 = Word2(LS) & " THOUSAND " & Word3(RT)
    End If
End Function

Public Function Word6(intNUM6 As String) As String
    Dim LF As String
    Dim RT As String
    
    LF = Left$(intNUM6, 1)
    RT = Right$(intNUM6, 5)
    
    If LF = "0" Then
        Word6 = Word5(RT)
    Else
        Word6 = Word1(LF) & " LAC(S) " & Word5(RT)
    End If
End Function
Public Function Word7(intNUM7 As String) As String
    Dim LF As String
    Dim LS As String
    Dim RF As String
    Dim rs As String
    
    LF = Left$(intNUM7, 1)
    LS = Left$(intNUM7, 2)
    RF = Right$(intNUM7, 5)
    rs = Right$(intNUM7, 6)
    
    If LF = "0" Then
        Word7 = Word6(rs)
    Else
        Word7 = Word2(LS) & " LAC(S) " & Word5(RF)
    End If
End Function
Public Function Word8(intNUM8 As String) As String
    Dim LF As String
    Dim LS As String
    Dim RF As String
    Dim rs As String
    
    LF = Left$(intNUM8, 1)
    rs = Right$(intNUM8, 7)
    
    If LF = "0" Then
        Word8 = Word7(rs)
    Else
        Word8 = Word1(LF) & " CRORE " & Word7(rs)
    End If
End Function
Public Function Word9(intNUM9 As String) As String
    Dim LF As String
    Dim LS As String
    Dim RF As String
    Dim rs As String
    
    LF = Left$(intNUM9, 1)
    LS = Left$(intNUM9, 2)
    rs = Right$(intNUM9, 7)
    RF = Right$(intNUM9, 8)
    
    If LF = "0" Then
        Word9 = Word8(RF)
    Else
        Word9 = Word2(LS) & " CRORE " & Word7(rs)
    End If
End Function
Public Function Word10(intNUM10 As String) As String
    Dim LF As String
    Dim RF As String
    
    LF = Left$(intNUM10, 3)
    RF = Right$(intNUM10, 7)
    
    Word10 = Word3(LF) & " CROCE " & Word7(RF)
    
End Function

Public Function DecimalX(strValue As String) As Double
    Dim intCount As Integer
    Dim i As Integer
    Dim strDECI As String
    
    intCount = Len(Trim(strValue))
    For i = 1 To intCount
        strDECI = Mid$(Trim(strValue), i, 1)
        If strDECI = "." Then
           DecimalX = i
           Exit Function
        End If
    Next i
End Function





