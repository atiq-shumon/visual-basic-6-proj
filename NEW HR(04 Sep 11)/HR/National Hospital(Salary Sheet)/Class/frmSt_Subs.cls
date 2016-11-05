VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "St_Subs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarConnstring As String
Private mvarSubs_Code As String
Private mvarCop As Currency
Private mvarClb As Currency
Private mvarUni As Currency
Private mvarWF As Currency
Private mvarBF As Currency
Private mvarMos As Currency
Private mvarTrack_Id As Long
Public Property Let Subs_Code(ByVal vData As String)
   mvarSubs_Code = vData
End Property
Public Property Get Subs_Code() As String
   Subs_Code = mvarSubs_Code
End Property
Public Property Let Cop(ByVal vData As Currency)
   mvarCop = vData
End Property
Public Property Get Cop() As Currency
   Cop = mvarCop
End Property
Public Property Let Clb(ByVal vData As Currency)
   mvarClb = vData
End Property
Public Property Get Clb() As Currency
   Clb = mvarClb
End Property
Public Property Let Uni(ByVal vData As Currency)
   mvarUni = vData
End Property
Public Property Get Uni() As Currency
   Uni = mvarUni
End Property
Public Property Let WF(ByVal vData As Currency)
   mvarWF = vData
End Property
Public Property Get WF() As Currency
   WF = mvarWF
End Property
Public Property Let BF(ByVal vData As Currency)
   mvarBF = vData
End Property
Public Property Get BF() As Currency
   BF = mvarBF
End Property
Public Property Let Mos(ByVal vData As Currency)
   mvarMos = vData
End Property
Public Property Get Mos() As Currency
   Mos = mvarMos
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
    cmd.CommandText = "St_Subs_IU"
    cmd(1) = mvarSubs_Code
    cmd(2) = mvarCop
    cmd(3) = mvarClb
    cmd(4) = mvarUni
    cmd(5) = mvarWF
    cmd(6) = mvarBF
    cmd(7) = mvarMos
    cmd(8) = mvarTrack_Id
    Set RS = cmd.Execute
    MsgBox RS!Message, vbOKOnly + vbExclamation
        
    Exit Function
SaveError:
    Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the St_Subs Table" + vbCr + vbCr + Err.Description
End Function
Public Function GetAll() As Recordset
On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New ADODB.Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "GetAllRecords 'Subs'"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
    If Not (RS.EOF Or RS.BOF) Then
        Set GetAll = RS
        Exit Function
    End If

GetXError:
    
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Production Table" + vbCr + vbCr + Err.Description

End Function
