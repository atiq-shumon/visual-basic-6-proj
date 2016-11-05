VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''ORACLE Database Connetivity Class
''-----------------------------------------------------------------------

'this class provides database connection string.
'it reads the connection parameters from the locate.dat file which is
'associated with the application.
'if the database is located somewhere in a LAN then the server name
'user and password should be changed to access it.

Private str_DataSource As String
Private str_User As String
Private str_Pass As String

Public Property Get Connection_String() As String

On Error GoTo Server_Unavailable

Dim snc As String
Dim FileNumber As Integer

FileNumber = FreeFile


Open App.Path + "\Locate_ORACLE.dat" For Input Access Read As #FileNumber

Input #FileNumber, str_DataSource
Input #FileNumber, str_User
Input #FileNumber, str_Pass

Close #FileNumber
    
snc = "Provider=OraOLEDB.Oracle.1;Persist Security Info=False"
''snc = "Provider=MSDAORA.1;Persist Security Info=False"

        If Not IsNull(str_Pass) Then
           snc = snc + ";Password=dn_ppbf"
'       + str_Pass
        End If

        If Not IsNull(str_User) Then
'            snc = snc + ";User ID=ppbf0809"
            snc = snc + ";User ID=ppbf"
        End If

        If Not IsNull(str_DataSource) Then
            snc = snc + ";Data Source=bank"
        End If

        Connection_String = snc

        Exit Property

Server_Unavailable:

    MsgBox "Client can not find the specified server. " _
    & "Please contact the system administrator." & Space(6) & Chr(10) _
    & "Network disconnected or a required file is missing.", vbCritical + vbOKOnly, "Server Unavailable"

End Property

