VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Loan_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarLoan_Nature As String

Private mvarConnstring As String
Private mvarEmp_ID As String
Private mvarLn_Id As String
Private mvarLn_Code As String
Private mvarLn_Amt As Currency
Private mvarInst_Amt As Currency
Private mvarNotes As String

Private mvarTrack_Id As Long

Private mvarInst_No As Integer 'local copy

Private mvarLn_Date As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarIntr_Amt As Currency 'local copy
'local variable(s) to hold property value(s)
Private mvarMONEY_TAKEN As String 'local copy
Public Property Let MONEY_TAKEN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MONEY_TAKEN = 5
    mvarMONEY_TAKEN = vData
End Property


Public Property Get MONEY_TAKEN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MONEY_TAKEN
    MONEY_TAKEN = mvarMONEY_TAKEN
End Property



Public Property Let Intr_Amt(ByVal vData As Currency)
    mvarIntr_Amt = vData
End Property


Public Property Get Intr_Amt() As Currency
    Intr_Amt = mvarIntr_Amt
End Property

Public Property Let Ln_Date(ByVal vData As Date)
    mvarLn_Date = vData
End Property
Public Property Get Ln_Date() As Date
    Ln_Date = mvarLn_Date
End Property

Public Property Let Inst_No(ByVal vData As Integer)
    mvarInst_No = vData
End Property
Public Property Get Inst_No() As Integer
    Inst_No = mvarInst_No
End Property
Public Property Let Emp_ID(ByVal vData As String)
   mvarEmp_ID = vData
End Property
Public Property Get Emp_ID() As String
   Emp_ID = mvarEmp_ID
End Property
Public Property Let Ln_Nature(ByVal vData As String)
   mvarLoan_Nature = vData
End Property
Public Property Let Ln_Id(ByVal vData As String)
   mvarLn_Id = vData
End Property
Public Property Get Ln_Id() As String
   Ln_Id = mvarLn_Id
End Property
Public Property Let Ln_Code(ByVal vData As String)
   mvarLn_Code = vData
End Property
Public Property Get Ln_Code() As String
   Ln_Code = mvarLn_Code
End Property
Public Property Let Ln_Amt(ByVal vData As Currency)
   mvarLn_Amt = vData
End Property
Public Property Get Ln_Amt() As Currency
   Ln_Amt = mvarLn_Amt
End Property
Public Property Let Inst_Amt(ByVal vData As Currency)
   mvarInst_Amt = vData
End Property
Public Property Get Inst_Amt() As Currency
   Inst_Amt = mvarInst_Amt
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
    
  '  If mvarEmp_Id = "" Then Exit Sub
    
    Dim conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Loan_Info_IU"
    cmd(1) = mvarEmp_ID
    cmd(2) = mvarLn_Id
    cmd(3) = mvarLn_Code
    cmd(4) = mvarLn_Amt
    cmd(5) = mvarInst_Amt
    cmd(6) = mvarInst_No
    cmd(7) = mvarIntr_Amt
    cmd(8) = mvarNotes
    cmd(9) = mvarLn_Date
    cmd(10) = mvarTrack_Id
    
    Set RS = cmd.Execute
    MsgBox RS!Message, vbOKOnly + vbExclamation
    
    Exit Function
SaveError:
    Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the Loan_Info Table" + vbCr + vbCr + Err.Description
End Function
Public Function GetAll() As Recordset
On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New ADODB.Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "GetAllRecords 'Loan_info','" + mvarEmp_ID + "','" + mvarLoan_Nature + "'"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
    'If Not (rs.EOF Or rs.BOF) Then
        Set GetAll = RS
        Exit Function
    'End If

GetXError:
    
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Production Table" + vbCr + vbCr + Err.Description

End Function
