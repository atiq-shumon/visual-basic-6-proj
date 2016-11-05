VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bonus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarConnstring As String

Private mvarPAY_MONTH As String
Private mvarPAY_YEAR As String
Private mvarBonus_Type As Long

Private mvarU_Id As String
Private mvarFiscl_Year As String 'local copy
Private mvarUnit As String 'local copy
Private mvarCost As String 'local copy
Public Property Let Cost(ByVal vData As String)

    mvarCost = vData
End Property
Public Property Get Cost() As String
    Cost = mvarCost
End Property
Public Property Let Unit(ByVal vData As String)
    mvarUnit = vData
End Property
Public Property Get Unit() As String
   Unit = mvarUnit
End Property
Public Property Let Fiscl_Year(ByVal vData As String)
    mvarFiscl_Year = vData
End Property
Public Property Get Fiscl_Year() As String
    Fiscl_Year = mvarFiscl_Year
End Property
Public Property Let PAY_MONTH(ByVal vData As String)
   mvarPAY_MONTH = vData
End Property
Public Property Get PAY_MONTH() As String
   PAY_MONTH = mvarPAY_MONTH
End Property
Public Property Let PAY_YEAR(ByVal vData As String)
   mvarPAY_YEAR = vData
End Property
Public Property Get PAY_YEAR() As String
   PAY_YEAR = mvarPAY_YEAR
End Property
Public Property Let Bonus_Type(ByVal vData As Long)
   mvarBonus_Type = vData
End Property
Public Property Get Bonus_Type() As Long
   Bonus_Type = mvarBonus_Type
End Property

Public Property Let U_Id(ByVal vData As String)
   mvarU_Id = vData
End Property
Public Property Get U_Id() As String
   U_Id = mvarU_Id
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
Dim Rs As New Recordset
    
    conn.Open mvarConnstring
    conn.CommandTimeout = 0
   
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Bonus_IU"
    cmd(1) = mvarUnit
    cmd(2) = mvarCost
    cmd(3) = mvarPAY_MONTH
    cmd(4) = mvarPAY_YEAR
    cmd(5) = mvarFiscl_Year
    cmd(6) = mvarBonus_Type
    cmd(7) = mvarU_Id
    Set Rs = cmd.Execute
    MsgBox Rs!Message, vbOKOnly + vbExclamation
    
    Exit Function
    
SaveError:
    Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the Bonus Table" + vbCr + vbCr + Err.Description
End Function
Public Function Delete()
On Error GoTo DeleteError
Dim conn As New Connection
Dim cmd As New Command
    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Bonus_Del"
    cmd.Execute
    Exit Function
DeleteError:
    Err.Raise vbObjectError + 513, "Delete Error", "Problems in Deleting Data from the Bonus Table" + vbCr + vbCr + Err.Description
End Function
Public Function GetAll() As Recordset
On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim Rs As New ADODB.Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "GetAllRecords 'Bonus'," + CStr(mvarBonus_Type) + ",'" + mvarPAY_MONTH + "','" + mvarPAY_YEAR + "'"
    Rs.CursorLocation = adUseClient
    Rs.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
    'If Not (rs.EOF Or rs.BOF) Then
        Set GetAll = Rs
        Exit Function
    'End If

GetXError:
    
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Production Table" + vbCr + vbCr + Err.Description

End Function

