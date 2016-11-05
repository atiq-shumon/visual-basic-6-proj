VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Inv_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarConnstring As String
Private mvarInv_ID As String
Private mvarInv_NM As String
Private mvarBank_Nm As String
Private mvarAmt As Currency
Private mvarInt_Rate As Currency
Private mvarInv_Dt As Date
Private mvarMat_Dt As Date
Private mvarNotes As String
Private mvarTrack_Id As Long
Public Property Let Inv_ID(ByVal vData As String)
   mvarInv_ID = vData
End Property
Public Property Get Inv_ID() As String
   Inv_ID = mvarInv_ID
End Property
Public Property Let Inv_NM(ByVal vData As String)
   mvarInv_NM = vData
End Property
Public Property Get Inv_NM() As String
   Inv_NM = mvarInv_NM
End Property
Public Property Let Bank_Nm(ByVal vData As String)
   mvarBank_Nm = vData
End Property
Public Property Get Bank_Nm() As String
   Bank_Nm = mvarBank_Nm
End Property
Public Property Let Amt(ByVal vData As Currency)
   mvarAmt = vData
End Property
Public Property Get Amt() As Currency
   Amt = mvarAmt
End Property
Public Property Let Int_Rate(ByVal vData As Currency)
   mvarInt_Rate = vData
End Property
Public Property Get Int_Rate() As Currency
   Int_Rate = mvarInt_Rate
End Property
Public Property Let Inv_Dt(ByVal vData As Date)
   mvarInv_Dt = vData
End Property
Public Property Get Inv_Dt() As Date
   Inv_Dt = mvarInv_Dt
End Property
Public Property Let Mat_Dt(ByVal vData As Date)
   mvarMat_Dt = vData
End Property
Public Property Get Mat_Dt() As Date
   Mat_Dt = mvarMat_Dt
End Property
Public Property Let Notes(ByVal vData As String)
   mvarNotes = vData
End Property
Public Property Get Notes() As String
   Notes = mvarNotes
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
    cmd.CommandText = "Inv_Info_IU"
    cmd(1) = mvarInv_ID
    cmd(2) = mvarInv_NM
    cmd(3) = mvarBank_Nm
    cmd(4) = mvarAmt
    cmd(5) = mvarInt_Rate
    cmd(6) = mvarInv_Dt
    cmd(7) = mvarMat_Dt
    cmd(8) = mvarNotes
    cmd(9) = mvarTrack_Id
    
    Set RS = cmd.Execute
    MsgBox RS!Message, vbOKOnly + vbExclamation
    Exit Function
SaveError:
    Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the Inv_Info Table" + vbCr + vbCr + Err.Description
End Function
   
Public Function GetAll() As Recordset
On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New ADODB.Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "GetAllRecords 'Invest'"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
    'If Not (rs.EOF Or rs.BOF) Then
        Set GetAll = RS
        Exit Function
   ' End If

GetXError:
    
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Production Table" + vbCr + vbCr + Err.Description

End Function

