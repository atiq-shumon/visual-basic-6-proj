VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmExtraBedMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Extra Bed"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CBOYRCODE 
      Height          =   315
      ItemData        =   "frmExtraBedMain.frx":0000
      Left            =   3270
      List            =   "frmExtraBedMain.frx":000A
      TabIndex        =   22
      Text            =   "YR-0708"
      Top             =   3630
      Width           =   1755
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   570
      Picture         =   "frmExtraBedMain.frx":0020
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Preview"
      Top             =   3510
      Width           =   495
   End
   Begin VB.CommandButton cmdSAVE 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   60
      Picture         =   "frmExtraBedMain.frx":068A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Save"
      Top             =   3510
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1080
      Picture         =   "frmExtraBedMain.frx":0CF4
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Exit"
      Top             =   3510
      Width           =   495
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   -30
      TabIndex        =   17
      Top             =   -60
      Width           =   5385
      Begin VB.Image Image1 
         Height          =   480
         Left            =   720
         Picture         =   "frmExtraBedMain.frx":1612
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Extra Bed Registration"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   645
         Left            =   1350
         TabIndex        =   18
         Top             =   150
         Width           =   5445
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   5355
      Begin VB.Frame Frame2 
         Height          =   3435
         Left            =   0
         TabIndex        =   1
         Top             =   -60
         Width           =   5145
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   4110
            Top             =   3120
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
            Left            =   2790
            Top             =   3090
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
         Begin VB.TextBox txtReg_no_extra 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   390
            TabIndex        =   16
            Top             =   330
            Width           =   4125
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            Caption         =   "End  Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3120
            TabIndex        =   8
            Top             =   2100
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            Caption         =   "Start Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   390
            TabIndex        =   7
            Top             =   2100
            Width           =   1275
         End
         Begin VB.TextBox txtCharge 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2220
            MaxLength       =   17
            TabIndex        =   6
            Text            =   "100"
            Top             =   1680
            Width           =   555
         End
         Begin VB.ComboBox comDepartmentRelease 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "comDepartmentFree"
            Top             =   1680
            Width           =   1590
         End
         Begin VB.ComboBox comSexRelease 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmExtraBedMain.frx":1EDC
            Left            =   1140
            List            =   "frmExtraBedMain.frx":1EE6
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1680
            Width           =   945
         End
         Begin VB.TextBox txtNameRelease 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   390
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   990
            Width           =   4065
         End
         Begin VB.TextBox TxtAgeRelease 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   420
            Locked          =   -1  'True
            MaxLength       =   17
            TabIndex        =   2
            Top             =   1680
            Width           =   555
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   390
            TabIndex        =   9
            Top             =   2370
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   582
            _Version        =   393216
            Format          =   60882945
            CurrentDate     =   38049
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Charge"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2190
            TabIndex        =   15
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Department"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2940
            TabIndex        =   14
            Top             =   1470
            Width           =   990
         End
         Begin VB.Label Label88 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   390
            TabIndex        =   13
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Label89 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sex"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1380
            TabIndex        =   12
            Top             =   1440
            Width           =   330
         End
         Begin VB.Label Label91 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   540
            TabIndex        =   11
            Top             =   1440
            Width           =   345
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Reg. No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   390
            TabIndex        =   10
            Top             =   150
            Width           =   720
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FISCAL YEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   2010
      TabIndex        =   23
      Top             =   3690
      Width           =   1200
   End
   Begin VB.Shape Shape1 
      Height          =   585
      Left            =   30
      Top             =   3450
      Width           =   1605
   End
End
Attribute VB_Name = "frmExtraBedMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Con As New MyConnection
Dim Conn2 As New Connection
Dim rs2 As New Recordset
Dim cmd As New Command









Private Sub CMDEXIT_Click()

    Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdSave_Click()
Dim validation As Variant
              Adodc1.ConnectionString = strcn.Connection_String
              Adodc1.RecordSource = "Select user_id, user_name, user_password,user_type,shift_name From Security Where (User_Id = '" & frmMAIN.lbluser_id & "')"
              Adodc1.Refresh
              validation = Adodc1.Recordset!user_id
                
                Dim conn As New ADODB.Connection
                Dim cmd As New ADODB.Command
                Dim rs As New ADODB.Recordset

                        Dim Param1 As New Parameter
                    If conn.State = 0 Then
                        conn.Open strcn.Connection_String
                    End If
                    Set cmd.ActiveConnection = conn
                    cmd.CommandType = adCmdText
    
                   Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, validation)
                    cmd.Parameters.Append Param1 'validation
                    cmd.Properties("PLSQLRSet") = True
    
                     cmd.CommandText = "{CALL shift_validation(?)}"
    
                Debug.Print cmd.CommandText
    
                    Set rs = cmd.Execute
    

                cmd.Properties("PLSQLRSet") = False
             If conn.State = 1 Then
                conn.Close
             End If
                
          Adodc2.ConnectionString = strcn.Connection_String
          Adodc2.RecordSource = "Select * From user_validation"
          Adodc2.Refresh
        

        
             If Adodc2.Recordset!validation = 0 Then
             MsgBox "Your Working Time has been Expired", vbInformation, " IT, DNMIH."
             Exit Sub
             End If
             





If txtNameRelease = "" Then
Exit Sub
Unload Me
End If
If txtCharge = "" Then
MsgBox "Extra Bed charge Required"
txtCharge.SetFocus
Exit Sub
End If
Call saveExrta_bed_info
MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."

End Sub
Private Sub saveExrta_bed_info()
Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
If conn.State = 0 Then
    conn.Open strcn.Connection_String
End If
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 5, frmExtraBed.txtRegNoExtraBed.Text)
    cmd.Parameters.Append Param1 'in_reg_no
    
   
    Set Param2 = cmd.CreateParameter("param2", adDouble, adParamInput, 10, Trim(txtCharge.Text))
    cmd.Parameters.Append Param2 'Bed_charge
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 2, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param3 'U_id default Sumon

    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 12, DTPicker1.Value)
    cmd.Parameters.Append Param4 'START OR END DATE
    


   Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 5, frmMAIN.lblBooth)
    cmd.Parameters.Append Param5 'booth
    
     Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 10, frmExtraBed.CBOYRCODE)
    cmd.Parameters.Append Param6 'booth
  
    
    
       cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Save_extra_Bed_info_indoor(?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set rs2 = cmd.Execute
    

  cmd.Properties("PLSQLRSet") = False
    
    
End Sub

Private Sub Form_Load()
    txtReg_no_extra = frmExtraBed.txtRegNoExtraBed.Text
    DTPicker1.Value = Date
    If Conn2.State = 0 Then
       Conn2.ConnectionString = strcn.Connection_String
       Conn2.Open
    End If
        cmd.ActiveConnection = Conn2
        cmd.CommandType = adCmdText
        cmd.CommandText = "select pat_name,pat_guard_name,sex,age,doc_dept  From in_door_pat_info_main Where in_reg_no ='" & Trim(frmExtraBed.txtRegNoExtraBed.Text) & "' AND YRCODE='" & Trim(frmExtraBed.CBOYRCODE) & "'"
      
       cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        rs2.CursorLocation = adUseClient

        rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
        If rs2.RecordCount > 0 Then
         txtNameRelease = rs2!pat_name
         TxtAgeRelease = rs2!age
         comSexRelease.Text = rs2!sex
         comDepartmentRelease = rs2!doc_dept
         
       cmd.Properties("iRowsetChange") = False
         
'rs2.Close
If Conn2.State = 1 Then
    Conn2.Close
End If
Else
 MsgBox "Invalid Registration No", vbInformation, "Warning: IT, DNMIH"
 'rs2.Close
If Conn2.State = 1 Then
    Conn2.Close
End If
 Exit Sub
 Unload Me

End If

End Sub

