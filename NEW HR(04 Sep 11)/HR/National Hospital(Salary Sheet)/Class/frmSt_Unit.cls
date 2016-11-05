VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "St_Unit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit
Private mvarConnString As String
Private mvarUnit_Code
Private mvarUnit_Name
Private mvarTrack_Id As Long
Public Property Let Unit_Code(ByVal vData As Long)
   mvarUnit_Code = vData
End Property
Public Property Get Unit_Code() As Long
   Unit_Code = mvarUnit_Code
End Property
Public Property Let Unit_Name(ByVal vData As Long)
   mvarUnit_Name = vData
End Property
Public Property Get Unit_Name() As Long
   Unit_Name = mvarUnit_Name
End Property
Public Property Let Track_Id(ByVal vData As Long)
   mvarTrack_Id = vData
End Property
Public Property Get Track_Id() As Long
   Track_Id = mvarTrack_Id
End Property
Public Property Let ConnString(ByVal vData As String)
   mvarConnString = vData
End Property
Public Property Get ConnString() As String
   ConnString = mvarConnString
End Property
Public Function Save()
On Error GoTo SaveError
Dim conn As New Connection
Dim cmd As New Command

    conn.Open mvarConnString
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "St_Unit_Save"
    cmd(1) = mvarUnit_Code
    cmd(2) = mvarUnit_Name
    cmd(3) = mvarTrack_Id
    cmd.Execute
    Exit Function
SaveError:
    Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the St_Unit Table" + vbCr + vbCr + Err.Description
End Function

Public Function GetAll() As Recordset
On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim Rs As New Recordset

    conn.Open mvarConnString
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "St_Unit_GetAll"
    Rs.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    Set GetAll = Rs
    Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the St_Unit Table" + vbCr + vbCr + Err.Description
End Function
