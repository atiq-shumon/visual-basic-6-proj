VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "St_Loan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit
Private mvarConnstring As String
Private mvarLn_Code As String
Private mvarLn_Nm As String
Private mvarCeil_Amt As Currency
Private mvarAmt_Prcnt As Currency
Private mvarInst_No As Long
Private mvarInt_Rate As Long
Private mvarTrack_Id As Long
Public Property Let Ln_Code(ByVal vData As String)
   mvarLn_Code = vData
End Property
Public Property Get Ln_Code() As String
   Ln_Code = mvarLn_Code
End Property
Public Property Let Ln_Nm(ByVal vData As String)
   mvarLn_Nm = vData
End Property
Public Property Get Ln_Nm() As String
   Ln_Nm = mvarLn_Nm
End Property
Public Property Let Ceil_Amt(ByVal vData As Currency)
   mvarCeil_Amt = vData
End Property
Public Property Get Ceil_Amt() As Currency
   Ceil_Amt = mvarCeil_Amt
End Property
Public Property Let Amt_Prcnt(ByVal vData As Currency)
   mvarAmt_Prcnt = vData
End Property
Public Property Get Amt_Prcnt() As Currency
   Amt_Prcnt = mvarAmt_Prcnt
End Property
Public Property Let Inst_No(ByVal vData As Long)
   mvarInst_No = vData
End Property
Public Property Get Inst_No() As Long
   Inst_No = mvarInst_No
End Property
Public Property Let Int_Rate(ByVal vData As Long)
   mvarInt_Rate = vData
End Property
Public Property Get Int_Rate() As Long
   Int_Rate = mvarInt_Rate
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
    cmd.CommandText = "St_Loan_IU"
    cmd(1) = mvarLn_Code
    cmd(2) = mvarLn_Nm
    cmd(3) = mvarCeil_Amt
    cmd(4) = mvarAmt_Prcnt
    cmd(5) = mvarInst_No
    cmd(6) = mvarInt_Rate
    cmd(7) = mvarTrack_Id
    
    Set RS = cmd.Execute
    MsgBox RS!Message, vbOKOnly + vbExclamation
    
    Exit Function
SaveError:
    Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the St_Loan Table" + vbCr + vbCr + Err.Description
End Function

Public Function GetAll() As Recordset
On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New ADODB.Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "GetAllRecords 'Loan'"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
    If Not (RS.EOF Or RS.BOF) Then
        Set GetAll = RS
        Exit Function
    End If
End Function

