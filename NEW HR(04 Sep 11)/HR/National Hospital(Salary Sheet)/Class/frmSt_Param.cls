VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "St_Param"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Class Developed & Designed by : Shajeed Md Shafique
Option Explicit
Private mvarConnstring As String
Private mvarPolicy_No As Long
Private mvarPolicy_Desc As Long
Private mvarFlag As Long
Private mvarValue As Long
Public Property Let Policy_No(ByVal vData As Long)
   mvarPolicy_No = vData
End Property
Public Property Get Policy_No() As Long
   Policy_No = mvarPolicy_No
End Property
Public Property Let Policy_Desc(ByVal vData As Long)
   mvarPolicy_Desc = vData
End Property
Public Property Get Policy_Desc() As Long
   Policy_Desc = mvarPolicy_Desc
End Property
Public Property Let Flag(ByVal vData As Long)
   mvarFlag = vData
End Property
Public Property Get Flag() As Long
   Flag = mvarFlag
End Property
Public Property Let Value(ByVal vData As Long)
   mvarValue = vData
End Property
Public Property Get Value() As Long
   Value = mvarValue
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

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "St_Param_Save"
    cmd(1) = mvarPolicy_No
    cmd(2) = mvarPolicy_Desc
    cmd(3) = mvarFlag
    cmd(4) = mvarValue
    cmd.Execute
    Exit Function
SaveError:
    Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the St_Param Table" + vbCr + vbCr + Err.Description
End Function


Public Function GetAll() As Recordset
On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "St_Param_GetAll"
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    Set GetAll = RS
    Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the St_Param Table" + vbCr + vbCr + Err.Description
End Function
