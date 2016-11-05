Attribute VB_Name = "Open"

Option Explicit
Public Sub Main()

U_Id = "it"


Form1.Show


End Sub
''
''
''Public Sub Set_Parameter(Policy_No As Integer, Flag As Integer, Value As String)
''
'''On ERROR Resume Next
''    con.ConnectionString = strCN.Connection_String
    
''    con.Open
''
''    Set Cmd.ActiveConnection = con
''        Cmd.CommandText = "exec  Set_Param " + Trim(CStr(Policy_No)) + "," _
''        + Trim(Flag) + ",'" + Trim(Value) + "'"
''
''    Cmd.Execute
''    con.Close
''
''End Sub
''
''Public Function Get_Param_Flag(Policy As Integer)
''    Dim rs As New Recordset
''
''    con.ConnectionString = strCN.Connection_String
    
''    con.Open
''    Cmd.CommandText = "exec POP_Param " + CStr(Policy)
''    Cmd.ActiveConnection = con
''        Set rs = Cmd.Execute
''        Get_Param_Flag = rs.Fields(1)
''    con.Close
''End Function
''
''Public Function Get_Param_Value(Policy As Integer)
''    Dim rs As New Recordset
''
''    con.ConnectionString = strCN.Connection_String
    
''    con.Open
''    Cmd.CommandText = "exec POP_Param " + CStr(Policy)
''    Cmd.ActiveConnection = con
''        Set rs = Cmd.Execute
''        Get_Param_Value = rs.Fields(2)
''    con.Close
''End Function
