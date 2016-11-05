VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "St_HRent2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit
Private mvarConnstring As String
Private mvarSlab_code As String
Private mvarBasic_From As Currency
Private mvarBasic_To As Currency
Private mvarDed_Rate As Currency
Private mvarTrack_Id As Long
Public Property Let Slab_code(ByVal vData As String)
   mvarSlab_code = vData
End Property
Public Property Get Slab_code() As String
   Slab_code = mvarSlab_code
End Property
Public Property Let Basic_From(ByVal vData As Currency)
   mvarBasic_From = vData
End Property
Public Property Get Basic_From() As Currency
   Basic_From = mvarBasic_From
End Property
Public Property Let Basic_To(ByVal vData As Currency)
   mvarBasic_To = vData
End Property
Public Property Get Basic_To() As Currency
   Basic_To = mvarBasic_To
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
    cmd.CommandText = "St_HRent2_IU"
    cmd(1) = mvarSlab_code
    cmd(2) = mvarBasic_From
    cmd(3) = mvarBasic_To
    cmd(4) = mvarDed_Rate
    cmd(5) = mvarTrack_Id
    Set RS = cmd.Execute
    
    MsgBox RS!Message, vbOKOnly + vbExclamation
    
    Exit Function
SaveError:
    Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the St_HRent2 Table" + vbCr + vbCr + Err.Description
End Function
Public Function Delete()
On Error GoTo DeleteError
Dim conn As New Connection
Dim cmd As New Command

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "St_HRent2_Del"
    cmd(1) = mvarSlab_code
    cmd.Execute
    Exit Function
DeleteError:
    Err.Raise vbObjectError + 513, "Delete Error", "Problems in Deleting Data from the St_HRent2 Table" + vbCr + vbCr + Err.Description
End Function
Public Function GetX() As Recordset
On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "St_HRent2_GetX"
    cmd(1) = mvarSlab_code
    Set RS = cmd.Execute
    Set GetX = RS
    If Not (RS.EOF And RS.BOF) Then
       mvarBasic_From = RS("Basic_From")
       mvarBasic_To = RS("Basic_To")
       mvarDed_Rate = RS("Ded_Rate")
       mvarTrack_Id = RS("Track_ID")
    Else
       mvarBasic_From = ""
       mvarBasic_To = ""
       mvarDed_Rate = ""
       mvarTrack_Id = ""
    End If
    Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the St_HRent2 Table" + vbCr + vbCr + Err.Description
End Function
Public Function GetAll() As Recordset
On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New ADODB.Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "GetAllRecords 'HR2'"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
    If Not (RS.EOF Or RS.BOF) Then
        Set GetAll = RS
        Exit Function
    End If

End Function
