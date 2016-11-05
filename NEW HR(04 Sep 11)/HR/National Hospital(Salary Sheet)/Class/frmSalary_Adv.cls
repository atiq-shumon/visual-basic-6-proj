VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Salary_Adv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarConnstring As String
Private mvarEmp_ID As String
Private mvarPAY_MONTH As String
Private mvarPAY_YEAR As String
Private mvarAMOUNT As Currency
Private mvarIssue_Dt As Date
Private mvarTrack_Id As Long
Private mvarU_Id  As String

Public Property Let Emp_ID(ByVal vData As String)
   mvarEmp_ID = vData
End Property
Public Property Get Emp_ID() As String
   Emp_ID = mvarEmp_ID
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
Public Property Let AMOUNT(ByVal vData As Currency)
   mvarAMOUNT = vData
End Property
Public Property Get AMOUNT() As Currency
   AMOUNT = mvarAMOUNT
End Property
Public Property Let Issue_Dt(ByVal vData As Date)
   mvarIssue_Dt = vData
End Property
Public Property Get Issue_Dt() As Date
   Issue_Dt = mvarIssue_Dt
End Property
Public Property Let Track_Id(ByVal vData As Long)
   mvarTrack_Id = vData
End Property
Public Property Get Track_Id() As Long
   Track_Id = mvarTrack_Id
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
Dim RS As New Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Salary_Adv_IU"
    cmd(1) = mvarEmp_ID
    cmd(2) = mvarPAY_MONTH
    cmd(3) = mvarPAY_YEAR
    cmd(4) = mvarAMOUNT
    cmd(5) = mvarIssue_Dt
    cmd(6) = mvarTrack_Id
    cmd(7) = mvarU_Id
    
     
    Set RS = cmd.Execute
    MsgBox RS!Message, vbOKOnly + vbExclamation
    
    Exit Function
SaveError:
    Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the Salary_Adv Table" + vbCr + vbCr + Err.Description
End Function
Public Function Delete()
On Error GoTo DeleteError
Dim conn As New Connection
Dim cmd As New Command

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Salary_Adv_Del"
    cmd.Execute
    Exit Function
DeleteError:
    Err.Raise vbObjectError + 513, "Delete Error", "Problems in Deleting Data from the Salary_Adv Table" + vbCr + vbCr + Err.Description
End Function
Public Function GetX() As Recordset
On Error GoTo GetXError
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Salary_Adv_GetX"
    Set RS = cmd.Execute
    Set GetX = RS
    If Not (RS.EOF And RS.BOF) Then
       mvarEmp_ID = RS("Emp_Id")
       mvarPAY_MONTH = RS("Pay_Month")
       mvarPAY_YEAR = RS("Pay_Year")
       mvarAMOUNT = RS("Amount")
       mvarIssue_Dt = RS("Issue_Dt")
       mvarTrack_Id = RS("Track_Id")
       mvarU_Id = RS("u_id")
       
    Else
       mvarEmp_ID = ""
       mvarPAY_MONTH = ""
       mvarPAY_YEAR = ""
       mvarAMOUNT = ""
       mvarIssue_Dt = ""
       mvarTrack_Id = ""
       mvarU_Id = ""
       
    End If
    Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Salary_Adv Table" + vbCr + vbCr + Err.Description
End Function
Public Function GetAll() As Recordset
On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "GetAllRecords 'Sal_Adv','" & mvarPAY_MONTH & "','" + mvarPAY_YEAR + "'"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
    If Not (RS.EOF Or RS.BOF) Then
        Set GetAll = RS
        Exit Function
    End If

    
End Function
