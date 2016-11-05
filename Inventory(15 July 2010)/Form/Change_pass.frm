VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Change_pass 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change Password"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H80000001&
      Height          =   1005
      Left            =   -30
      TabIndex        =   16
      Top             =   4200
      Width           =   6945
      Begin LVbuttons.LaVolpeButton cmdADD 
         Height          =   345
         Left            =   4470
         TabIndex        =   5
         ToolTipText     =   "Click to Add"
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "&New"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Change_pass.frx":0000
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmdExit 
         Height          =   345
         Left            =   5580
         TabIndex        =   6
         ToolTipText     =   "Click to Close"
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "&Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Change_pass.frx":001C
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmdSAVE 
         Height          =   345
         Left            =   3360
         TabIndex        =   4
         ToolTipText     =   "Click to Change Password"
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Update"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Change_pass.frx":0038
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   780
      Top             =   4350
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   780
      Top             =   4350
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Change_password 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   0
      TabIndex        =   9
      Top             =   2940
      Width           =   6855
      Begin VB.TextBox txtCpass 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2025
         TabIndex        =   3
         Top             =   780
         Width           =   4455
      End
      Begin VB.TextBox txtpass 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2010
         TabIndex        =   2
         Top             =   360
         Width           =   4485
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   203
         TabIndex        =   11
         Top             =   450
         Width           =   1635
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   225
         TabIndex        =   10
         Top             =   930
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   -30
      TabIndex        =   7
      Top             =   -30
      Width           =   6945
      Begin VB.Frame Frame2 
         BackColor       =   &H80000001&
         Height          =   1005
         Left            =   -60
         TabIndex        =   14
         Top             =   -90
         Width           =   6915
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Change Password Utility"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   705
            Left            =   1440
            TabIndex        =   15
            Top             =   210
            Width           =   4305
         End
      End
      Begin VB.TextBox txtpassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2070
         PasswordChar    =   "?"
         TabIndex        =   1
         Top             =   2250
         Width           =   4425
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   2070
         TabIndex        =   0
         Top             =   1530
         Width           =   4425
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   345
         Left            =   2070
         TabIndex        =   17
         Top             =   1890
         Width           =   4425
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Password:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   13
         Top             =   2310
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Id:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   12
         Top             =   1590
         Width           =   705
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " Name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   1950
         Width           =   705
      End
   End
End
Attribute VB_Name = "Change_pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public con As New MyConnection
Dim Conn2 As New Connection
Dim rs2 As New Recordset
Dim cmd As New Command


Private Sub cmdADD_Click()
txtpass = ""
txtCpass = ""
txtpassword = ""
Text1.SetFocus
' txtpass.SetFocus
End Sub

Private Sub cmdDelete_Click(Index As Integer)

End Sub

Private Sub cmdExit_Click()
'Dim reply As String
'    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
'    If reply = vbYes Then
        Unload Me
'    End If

End Sub
Private Sub cmdSAVE_Click()
   Dim Conn As New ADODB.Connection
   Dim cmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim Param1 As New Parameter

   If Text1 = "" Then
      MsgBox "User id Required", vbInformation, "IT Division, DNMIH"
      Text1.SetFocus
      Exit Sub
  End If
  If txtpassword = "" Then
     MsgBox "User Previous password Required", vbInformation, "IT Division, DNMIH"
     txtpassword.SetFocus
      Exit Sub
 End If
 If txtpass = "" Then
    MsgBox "User New password Required", vbInformation, "IT Division, DNMIH"
    txtpass.SetFocus
    Exit Sub
End If
                
If txtCpass = "" Then
   MsgBox "User New Confirm password Required", vbInformation, "IT Division, DNMIH"
   txtCpass.SetFocus
   Exit Sub
End If

   Call save_change_password
   cmdADD_Click

  If Conn.State = 1 Then
      Conn.Close
  End If
End Sub
Private Sub save_change_password()
  Dim con As New ADODB.Connection
  Dim cmd As New ADODB.Command
  con.ConnectionString = objmyCon.ConnectionString
                
   con.Open
   Set cmd.ActiveConnection = con
   cmd.CommandText = "update UserInfo set UserPass='" & Trim(txtpass.Text) & "'" & _
             " where to_number(userid)=to_number(('" & Trim(Text1.Text) & "'))"
            
                                  
  cmd.Execute
  con.Close
  MsgBox "Updated Successfully", vbInformation, cmp
End Sub

Private Sub Form_Load()
txtpass.BackColor = &H808080
txtCpass.BackColor = &H808080
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(Text1) <> 0 Then
            txtpassword.SetFocus
        Else
            MsgBox "Please enter any User Id.", vbInformation, "IT Division, DNMIH"
        End If
    End If
End Sub

Private Sub txtCpass_GotFocus()
txtCpass.BackColor = &HFFFFFF
End Sub

Private Sub txtCpass_LostFocus()
txtCpass.BackColor = &H808080
End Sub

Private Sub txtpass_GotFocus()
txtpass.BackColor = &HFFFFFF
End Sub

Private Sub txtpass_LostFocus()
txtpass.BackColor = &H808080


End Sub

Private Sub txtPassWord_KeyPress(KeyAscii As Integer)
If Text1 = "" Then
    MsgBox "User Id Required", vbCritical, "Warning"
    Text1.SetFocus
    Exit Sub
Else
    If KeyAscii = 13 Then
        Adodc1.ConnectionString = objmyCon.ConnectionString
        Adodc1.RecordSource = "Select UserID, UserName, UserPass From Userinfo Where (UserID = '" & Text1 & "')"
        Adodc1.Refresh

        
        If Adodc1.Recordset.EOF = True Then
            MsgBox "No such ID exists.", vbCritical, "Warning"
            Text1 = ""
            txtpassword = ""
            Text1.SetFocus
            Exit Sub
        Else
            If txtpassword = Adodc1.Recordset!UserPass Then
            Text3 = Adodc1.Recordset!userName
            txtpass.BackColor = &HFFFFFF
            
            txtCpass.BackColor = &HFFFFFF
            
            txtpass.Enabled = True
            txtCpass.Enabled = True
            txtpass.SetFocus
            Else
                MsgBox "Incorrect Password", vbCritical, "Warning"
                txtpassword = ""
                Exit Sub
            End If


        End If
    End If
End If

End Sub
