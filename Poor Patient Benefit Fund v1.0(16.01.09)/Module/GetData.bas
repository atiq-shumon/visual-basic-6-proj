Attribute VB_Name = "GetData"
Option Explicit

Public Sub GetControlCode(frm As Form, strAccName As String)
    On Error GoTo err_loop
'    frm.txtAccHead.Text = ""
'    Dim strUserAcc As String
'    'If con.State = adStateOpen Then con.Close
'
'    Con.ConnectionTimeout = 0
'    Con.Open strcn
'
'    RS.Open "select acc_code,user_acc from acct where acc_name='" & Trim(strAccName) & "'", Con
'    If RS.EOF = False Then
'       strUserAcc = RS!user_acc
'       frm.txtAccHead.Text = RS!acc_code
'    End If
'
'    RS.Close
'    Con.Close
'
'    frm.cboUserHead.Text = strUserAcc
    
    Exit Sub
    
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub
Public Sub GetControlName(frm As Form, strUserAcc As String)
    On Error GoTo err_loop
'    frm.txtAccHead.Text = ""
'
'    Con.ConnectionTimeout = 0
'    Con.Open strcn
'
'    RS.Open "select acc_code,acc_name from acct where user_acc='" & Trim(strUserAcc) & "'", Con
'    If RS.EOF = False Then
'       frm.cboHeadName.Text = RS!acc_name
'       frm.txtAccHead.Text = RS!acc_code
'    End If
'
'    RS.Close
'    Con.Close
'
    Exit Sub
    
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub
Public Sub GetUnitName(frm As Form)
    On Error GoTo err_loop
'
'    frm.cboUnit.Clear
'
'    Con.ConnectionTimeout = 0
'    Con.Open strcn
'
'    RS.Open "select prj_name from project order by prj_code ", Con
'    If RS.EOF = False Then
'       Do Until RS.EOF
'       frm.cboUnit.AddItem RS!prj_name
'       RS.MoveNext
'       Loop
'    End If
'
'    RS.Close
'    Con.Close
    
    Exit Sub
    
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub
Public Sub GetUserAcc(frm As Form)
    On Error GoTo err_loop
'
'    frm.cboUserAcc.Clear
'
'    Con.ConnectionTimeout = 0
'    Con.Open strcn
'
'    RS.Open "select user_acc from acct order by user_acc ", Con
'    If RS.EOF = False Then
'       Do Until RS.EOF
'       frm.cboUserAcc.AddItem RS!user_acc
'       RS.MoveNext
'       Loop
'    End If
'
'    RS.Close
'    Con.Close
    
    Exit Sub
   
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

Public Sub GetAccName(frm As Form, strUserAcc As String)
    Dim Conn As New Connection
    Dim cmd As New Command
    Dim RS As New Recordset

    Conn.Open strcn.Connection_String
    Set cmd.ActiveConnection = Conn
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "Select * from acct"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
    
    If Not (RS.EOF Or RS.BOF) Then
    
       
        ''--Set GetAll = RS
'        Exit Function
    End If


'    On Error GoTo err_loop
'
'    frm.txtacc_name.Text = ""
'
'    Con.ConnectionTimeout = 0
'    Con.Open strcn
'
'    RS.Open "select acc_name from acct where user_acc='" & Trim(strUserAcc) & "'", Con
'    If RS.EOF = False Then
'       frm.txtacc_name.Text = RS!acc_name
'    End If
'
'    RS.Close
'    Con.Close
'
'    Exit Sub
'
'err_loop:
'    MsgBox Err.Description, vbCritical
'    Resume Next
End Sub

Public Sub GetUnitCode(frm As Form, strUnitName As String)
    On Error GoTo err_loop
    
'    frm.txtUnitCode.Text = ""
'
'    Con.ConnectionTimeout = 0
'    Con.Open strcn
'
'    RS.Open "select prj_code from project where prj_name='" & Trim(strUnitName) & "'", Con
'    If RS.EOF = False Then
'       frm.txtUnitCode.Text = RS!prj_code
'    End If
'
'    RS.Close
'    Con.Close
    
    Exit Sub
   
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

