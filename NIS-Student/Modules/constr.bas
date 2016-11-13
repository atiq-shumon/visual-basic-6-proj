Attribute VB_Name = "connectionstring"
Option Explicit
Public buy_src_mode As String
Public con As New ADODB.connection
Public cmd As New ADODB.Command
Public rs As New ADODB.Recordset
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public strcn As New Class1 '(adodb.connection)
Public rptMode As Integer
Public strBSISB As String
Public mode As String
Public VeiwMode As String
Public strUid As String
Public soft_user As String
Global GConnString As String
Public Param_mode As Integer
Global cmp As String

Public Sub Main()
    ''GConnString = "Provider=SQLOLEDB.1;Password=123;Persist Security Info=True;User ID=sa;Initial Catalog=NSMS;Data Source=."
    cmp = "Dhaka Soft Ltd."
   GConnString = strcn.connection
    Dim f As New frmMain
    f.Show
    soft_user = "00001"
End Sub
Public Function getdata(SQLString As String) As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = SQLString
 Set rs = cmd.Execute
Set getdata = rs
End Function
Public Function Get_Code(SString As String) As String
    Get_Code = Trim(Mid(Trim(SString), InStr(Trim(SString), "~") + 1))
End Function

Public Function Get_Description(SString As String) As String
    'Get_Description = Trim(Mid(Trim(SString), 1, InStr(Trim(SString), "~") + 1) - 1)
    Get_Description = Trim(Mid(Trim(SString), 1, InStr(Trim(SString), "~") - 1))
End Function

