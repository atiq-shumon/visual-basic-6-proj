VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "St_BnsFs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarConnstring As String
Private mvarBonus_Nm As String
Private mvarRate As Currency
Private mvarTrack_Id As Long
'local variable(s) to hold property value(s)
Private mvarBonus_Code As String 'local copy
Public Property Let Bonus_Code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Bonus_Code = 5
    mvarBonus_Code = vData
End Property
Public Property Get Bonus_Code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Bonus_Code
    Bonus_Code = mvarBonus_Code
End Property
Public Property Let Bonus_Nm(ByVal vData As String)
   mvarBonus_Nm = vData
End Property
Public Property Get Bonus_Nm() As String
   Bonus_Nm = mvarBonus_Nm
End Property
Public Property Let Rate(ByVal vData As Currency)
   mvarRate = vData
End Property
Public Property Get Rate() As Currency
   Rate = mvarRate
End Property
Public Property Let Track_Id(ByVal vData As Long)
   mvarTrack_Id = vData
End Property
Public Property Get Track_Id() As Long
   Track_Id = mvarTrack_Id
End Property
Public Property Let Connstring(ByVal vData As String)
   mvarConnstring = vData
End Property
Public Property Get Connstring() As String
   Connstring = mvarConnstring
End Property
Public Function Save()
On Error GoTo SaveError
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "St_BnsFs_IU"
    cmd(1) = mvarBonus_Code
    cmd(2) = mvarBonus_Nm
    cmd(3) = mvarRate
    cmd(4) = mvarTrack_Id
    
    Set RS = cmd.Execute
    MsgBox RS!Message, vbOKOnly + vbExclamation
    
    Exit Function
SaveError:
    Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the St_BnsFs Table" + vbCr + vbCr + Err.Description
End Function
Public Function Delete()
On Error GoTo DeleteError
Dim conn As New Connection
Dim cmd As New Command

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "St_BnsFs_Del"
    cmd.Execute
    Exit Function
DeleteError:
    Err.Raise vbObjectError + 513, "Delete Error", "Problems in Deleting Data from the St_BnsFs Table" + vbCr + vbCr + Err.Description
End Function
Public Function GetX() As Recordset
On Error GoTo GetXError
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "St_BnsFs_GetX"
    Set RS = cmd.Execute
    Set GetX = RS
    If Not (RS.EOF And RS.BOF) Then
       mvarBonus_Nm = RS("Bonus_Nm")
       mvarRate = RS("Rate")
       mvarTrack_Id = RS("Track_ID")
    Else
       mvarBonus_Nm = ""
       mvarRate = ""
       mvarTrack_Id = ""
    End If
    Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the St_BnsFs Table" + vbCr + vbCr + Err.Description
End Function
Public Function GetAll() As Recordset

On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New ADODB.Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "GetAllRecords 'BnsFs'"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
    If Not (RS.EOF Or RS.BOF) Then
        Set GetAll = RS
        Exit Function
    End If

End Function
