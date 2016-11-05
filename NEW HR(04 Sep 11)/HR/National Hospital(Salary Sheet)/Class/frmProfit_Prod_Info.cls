VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Profit_Prod_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarConnstring As String
Private mvarFscl_year As String
Private mvarProfit_Amt As Currency
Private mvarExp_Prod As Currency
Private mvarAct_Prod As Currency
Private mvarTrack_Id As Long
Public Property Let Fscl_year(ByVal vData As String)
   mvarFscl_year = vData
End Property
Public Property Get Fscl_year() As String
   Fscl_year = mvarFscl_year
End Property
Public Property Let Profit_Amt(ByVal vData As Currency)
   mvarProfit_Amt = vData
End Property
Public Property Get Profit_Amt() As Currency
   Profit_Amt = mvarProfit_Amt
End Property
Public Property Let Exp_Prod(ByVal vData As Currency)
   mvarExp_Prod = vData
End Property
Public Property Get Exp_Prod() As Currency
   Exp_Prod = mvarExp_Prod
End Property
Public Property Let Act_Prod(ByVal vData As Currency)
   mvarAct_Prod = vData
End Property
Public Property Get Act_Prod() As Currency
   Act_Prod = mvarAct_Prod
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
    cmd.CommandText = "Profit_Prod_Info_IU"
    cmd(1) = mvarFscl_year
    cmd(2) = mvarProfit_Amt
    cmd(3) = mvarExp_Prod
    cmd(4) = mvarAct_Prod
    cmd(5) = mvarTrack_Id
    
    Set RS = cmd.Execute
    MsgBox RS!Message, vbOKOnly + vbExclamation

    Exit Function
SaveError:
    Err.Raise vbObjectError + 513, "Save Error", "Problems in Saving Data to the Profit_Prod_Info Table" + vbCr + vbCr + Err.Description
End Function
Public Function Delete()
On Error GoTo DeleteError
Dim conn As New Connection
Dim cmd As New Command

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Profit_Prod_Info_Del"
    cmd.Execute
    Exit Function
DeleteError:
    Err.Raise vbObjectError + 513, "Delete Error", "Problems in Deleting Data from the Profit_Prod_Info Table" + vbCr + vbCr + Err.Description
End Function
Public Function GetX() As Recordset
On Error GoTo GetXError
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Profit_Prod_Info_GetX"
    Set RS = cmd.Execute
    Set GetX = RS
    If Not (RS.EOF And RS.BOF) Then
       mvarFscl_year = RS("Fscl_year")
       mvarProfit_Amt = RS("Profit_Amt")
       mvarExp_Prod = RS("Exp_Prod")
       mvarAct_Prod = RS("Act_Prod")
       mvarTrack_Id = RS("Track_ID")
    Else
       mvarFscl_year = ""
       mvarProfit_Amt = ""
       mvarExp_Prod = ""
       mvarAct_Prod = ""
       mvarTrack_Id = ""
    End If
    Exit Function
GetXError:
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Profit_Prod_Info Table" + vbCr + vbCr + Err.Description
End Function
Public Function GetAll() As Recordset
On Error GoTo GetXError
Dim conn As New Connection
Dim cmd As New Command
Dim RS As New ADODB.Recordset

    conn.Open mvarConnstring
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "GetAllRecords 'Prod'"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn, adOpenStatic, adLockOptimistic
    
    If Not (RS.EOF Or RS.BOF) Then
        Set GetAll = RS
        Exit Function
    End If

GetXError:
    
    Err.Raise vbObjectError + 513, "Error in Data Get", "Problems in Getting Data from the Production Table" + vbCr + vbCr + Err.Description

End Function

