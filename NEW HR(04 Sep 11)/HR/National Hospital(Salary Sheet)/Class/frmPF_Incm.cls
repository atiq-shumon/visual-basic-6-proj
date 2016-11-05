VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PF_Incm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarConnstring As String
Private mvarEmp_ID
Private mvarPAY_MONTH
Private mvarPAY_YEAR
Private mvarFscl_Yr As String
Private mvarInc_won_part As Currency
Private mvarInc_Co_Part As String
Private mvarTrack_Id As Long
Public Property Let Emp_ID(ByVal vData As Long)
   mvarEmp_ID = vData
End Property
Public Property Get Emp_ID() As Long
   Emp_ID = mvarEmp_ID
End Property
Public Property Let PAY_MONTH(ByVal vData As Long)
   mvarPAY_MONTH = vData
End Property
Public Property Get PAY_MONTH() As Long
   PAY_MONTH = mvarPAY_MONTH
End Property
Public Property Let PAY_YEAR(ByVal vData As Long)
   mvarPAY_YEAR = vData
End Property
Public Property Get PAY_YEAR() As Long
   PAY_YEAR = mvarPAY_YEAR
End Property
Public Property Let Fscl_Yr(ByVal vData As String)
   mvarFscl_Yr = vData
End Property
Public Property Get Fscl_Yr() As String
   Fscl_Yr = mvarFscl_Yr
End Property
Public Property Let Inc_won_part(ByVal vData As Currency)
   mvarInc_won_part = vData
End Property
Public Property Get Inc_won_part() As Currency
   Inc_won_part = mvarInc_won_part
End Property
Public Property Let Inc_Co_Part(ByVal vData As String)
   mvarInc_Co_Part = vData
End Property
Public Property Get Inc_Co_Part() As String
   Inc_Co_Part = mvarInc_Co_Part
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

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "PF_Incm_Save"
    cmd(1) = mvarEmp_ID
    cmd(2) = mvarPAY_MONTH
    cmd(3) = mvarPAY_YEAR
    cmd(4) = mvarFscl_Yr
    cmd(5) = mvarInc_won_part
    cmd(6) = mvarInc_Co_Part
    cmd(7) = mvarTrack_Id
    cmd.Execute
    Exit Function
SaveError:
    Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the PF_Incm Table" + vbCr + vbCr + Err.Description
End Function

Public Function GetAll() As Recordset
On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "PF_Incm_GetAll"
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    Set GetAll = RS
    Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the PF_Incm Table" + vbCr + vbCr + Err.Description
End Function
