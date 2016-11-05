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
