VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function SalaryUpdateValidation(localMonth As String, localYear As String)
    Dim localConn As New ADODB.Connection
    Dim localCmd As New ADODB.Command
    Dim localRs As New ADODB.Recordset
    Dim msg As Integer
    localConn.ConnectionString = strCN.Connection_String
    localConn.Open
    localCmd.ActiveConnection = localConn
    localCmd.CommandType = adCmdText
    localCmd.CommandText = "Select sysdate from dual"
    Set localRs = localCmd.Execute
    
    If Not localRs.EOF Then
      If Val(localYear) < year(localRs(0)) Then
        msg = 1
      ElseIf Val(localYear) = year(localRs(0)) And Get_Month_No(localMonth) < month(localRs(0)) Then
        msg = 2
      End If
     
End If
   SalaryUpdateValidation = msg
End Function
