Attribute VB_Name = "Filemode"
Public Type File
modi1 As String * 180   'Moving Display 1 Length
modi2 As String * 180   'Moving Display 2 Length
modi3 As String * 180   'Moving Display 3 Length
modi4 As String * 180   'Moving Display 4 Length
modi5 As String * 180   'Moving Display 5 Length
End Type
Public Type Rg
Rg As Double
End Type
Public Rgg As Rg

Public Modfile As File

Sub main()
On Error Resume Next
MkDir "C:\windows"
MkDir "C:\windows\system"
If Dir("C:\windows\system\options.dll") = "" Then
Form4.Show
Exit Sub
End If
Close #1
Open "C:\windows\system\options.dll" For Random As #1 Len = Len(Rgg)
Get #1, 1, Rgg
Rggg = Rgg.Rg
If Rggg <> 9.99999999888889E+31 Then
End
Else
End If
Close #1
Form1.Show
End Sub








