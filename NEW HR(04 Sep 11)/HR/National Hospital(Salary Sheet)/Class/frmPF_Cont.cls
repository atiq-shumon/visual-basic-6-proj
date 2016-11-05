VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PF_Cont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Class Developed & Designed by : Shajeed Md Shafique
Option Explicit
Private mvarConnstring As String
Private mvarEmp_ID
Private mvarPAY_MONTH
Private mvarPAY_YEAR
Private mvarWon_Cont As Currency
Private mvarCo_Cont As Currency
Private mvarWth_Status As Currency
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
Public Property Let Won_Cont(ByVal vData As Currency)
   mvarWon_Cont = vData
End Property
Public Property Get Won_Cont() As Currency
   Won_Cont = mvarWon_Cont
End Property
Public Property Let Co_Cont(ByVal vData As Currency)
   mvarCo_Cont = vData
End Property
Public Property Get Co_Cont() As Currency
   Co_Cont = mvarCo_Cont
End Property
Public Property Let Wth_Status(ByVal vData As Currency)
   mvarWth_Status = vData
End Property
Public Property Get Wth_Status() As Currency
   Wth_Status = mvarWth_Status
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
    cmd.CommandText = "PF_Cont_Save"
    cmd(1) = mvarEmp_ID
    cmd(2) = mvarPAY_MONTH
    cmd(3) = mvarPAY_YEAR
    cmd(4) = mvarWon_Cont
    cmd(5) = mvarCo_Cont
    cmd(6) = mvarWth_Status
    cmd(7) = mvarTrack_Id
    cmd.Execute
    Exit Function
SaveError:
    Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the PF_Cont Table" + vbCr + vbCr + Err.Description
End Function
Public Function Delete()
On Error GoTo DeleteError
Dim conn As New Connection
Dim cmd As New Command

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "PF_Cont_Del"
    cmd.Execute
    Exit Function
DeleteError:
    Err.Raise vbObjectError + 513, "Delete Error", "Problems in Deleting Data from the PF_Cont Table" + vbCr + vbCr + Err.Description
End Function
Public Function GetX() As Recordset
On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "PF_Cont_GetX"
    Set RS = cmd.Execute
    Set GetX = RS
    If Not (RS.EOF And RS.BOF) Then
       mvarEmp_ID = RS("Emp_ID")
       mvarPAY_MONTH = RS("Pay_Month")
       mvarPAY_YEAR = RS("Pay_Year")
       mvarWon_Cont = RS("Won_Cont")
       mvarCo_Cont = RS("Co_Cont")
       mvarWth_Status = RS("Wth_Status")
       mvarTrack_Id = RS("Track_Id")
    Else
       mvarEmp_ID = ""
       mvarPAY_MONTH = ""
       mvarPAY_YEAR = ""
       mvarWon_Cont = ""
       mvarCo_Cont = ""
       mvarWth_Status = ""
       mvarTrack_Id = ""
    End If
    Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the PF_Cont Table" + vbCr + vbCr + Err.Description
End Function
Public Function GetAll() As Recordset
On Error Resume Next
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "PF_Cont_GetAll"
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    Set GetAll = RS
    Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the PF_Cont Table" + vbCr + vbCr + Err.Description
End Function
