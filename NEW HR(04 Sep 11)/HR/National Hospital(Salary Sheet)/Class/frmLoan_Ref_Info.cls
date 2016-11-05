VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Loan_Ref_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Class Developed & Designed by : Shajeed Md Shafique
Option Explicit
Private mvarConnstring As String
Private mvarLn_Id
Private mvarPAY_MONTH
Private mvarPAY_YEAR
Private mvarInst_Amt As Currency
Private mvarInt_Rate As Currency
Private mvarTrack_Id As Long
Public Property Let Ln_Id(ByVal vData As Long)
   mvarLn_Id = vData
End Property
Public Property Get Ln_Id() As Long
   Ln_Id = mvarLn_Id
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
Public Property Let Inst_Amt(ByVal vData As Currency)
   mvarInst_Amt = vData
End Property
Public Property Get Inst_Amt() As Currency
   Inst_Amt = mvarInst_Amt
End Property
Public Property Let Int_Rate(ByVal vData As Currency)
   mvarInt_Rate = vData
End Property
Public Property Get Int_Rate() As Currency
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

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Loan_Ref_Info_Save"
    cmd(1) = mvarLn_Id
    cmd(2) = mvarPAY_MONTH
    cmd(3) = mvarPAY_YEAR
    cmd(4) = mvarInst_Amt
    cmd(5) = mvarInt_Rate
    cmd(6) = mvarTrack_Id
    cmd.Execute
    Exit Function
SaveError:
    Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the Loan_Ref_Info Table" + vbCr + vbCr + Err.Description
End Function

Public Function GetAll() As Recordset
On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Loan_Ref_Info_GetAll"
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    Set GetAll = RS
    Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Loan_Ref_Info Table" + vbCr + vbCr + Err.Description
End Function
