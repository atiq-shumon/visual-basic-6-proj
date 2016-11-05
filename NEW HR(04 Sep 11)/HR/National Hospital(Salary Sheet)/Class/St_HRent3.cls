VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "St_HRent3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarConnstring As String
Private mvarType_Code As String
Private mvarRange_From As Long
Private mvarRange_To As Long
Private mvarDed_Rate As Currency
Private mvarTrack_Id As Long
Public Property Let Type_Code(ByVal vData As String)
   mvarType_Code = vData
End Property
Public Property Get Type_Code() As String
   Type_Code = mvarType_Code
End Property
Public Property Let Range_From(ByVal vData As Currency)
   mvarRange_From = vData
End Property
Public Property Get Range_From() As Currency
   Range_From = mvarRange_From
End Property
Public Property Let Range_To(ByVal vData As Currency)
   mvarRange_To = vData
End Property
Public Property Get Range_To() As Currency
   Range_To = mvarRange_To
End Property
Public Property Let Ded_Rate(ByVal vData As Currency)
   mvarDed_Rate = vData
End Property
Public Property Get Ded_Rate() As Currency
   Ded_Rate = mvarDed_Rate
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
    cmd.CommandText = "St_HRent3_IU"
    cmd(1) = mvarType_Code
    cmd(2) = mvarRange_From
    cmd(3) = mvarRange_To
    cmd(4) = mvarDed_Rate
    cmd(5) = mvarTrack_Id
    Set RS = cmd.Execute
    
    MsgBox RS!Message, vbOKOnly + vbExclamation
    
    Exit Function
SaveError:
    Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the St_HRent2 Table" + vbCr + vbCr + Err.Description
End Function

Public Function GetAll() As Recordset
On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New ADODB.Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "GetAllRecords 'HR3'"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
    If Not (RS.EOF Or RS.BOF) Then
        Set GetAll = RS
        Exit Function
    End If

GetXError:
    
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Production Table" + vbCr + vbCr + Err.Description

End Function
