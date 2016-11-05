VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExtraBedMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Extra Bed"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4365
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5325
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
         Left            =   1110
         Picture         =   "frmReAdvance.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Preview"
         Top             =   3750
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
         Left            =   600
         Picture         =   "frmReAdvance.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Save"
         Top             =   3750
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
         Left            =   1620
         Picture         =   "frmReAdvance.frx":0CD4
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Exit"
         Top             =   3750
         Width           =   495
      End
      Begin VB.TextBox TxtAgeRelease 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   17
         TabIndex        =   15
         Top             =   2550
         Width           =   555
      End
      Begin VB.TextBox txtNameRelease 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   570
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1860
         Width           =   4065
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   570
         TabIndex        =   11
         Top             =   1230
         Width           =   4095
         Begin VB.TextBox txtInregExtraBed 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   2940
            Left            =   -330
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
            Top             =   450
            Width           =   405
         End
      End
      Begin VB.ComboBox comSexRelease 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmReAdvance.frx":15F2
         Left            =   1320
         List            =   "frmReAdvance.frx":15FC
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2550
         Width           =   945
      End
      Begin VB.ComboBox comDepartmentRelease 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "comDepartmentFree"
         Top             =   2550
         Width           =   1590
      End
      Begin VB.TextBox txtCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         MaxLength       =   17
         TabIndex        =   3
         Text            =   "100"
         Top             =   2550
         Width           =   555
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   570
         TabIndex        =   2
         Top             =   2970
         Width           =   1275
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   3300
         TabIndex        =   1
         Top             =   2970
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   570
         TabIndex        =   13
         Top             =   3240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   582
         _Version        =   393216
         Format          =   58392577
         CurrentDate     =   38049
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Extra Bed Registration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   540
         TabIndex        =   20
         Top             =   240
         Width           =   3975
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
         Left            =   570
         TabIndex        =   19
         Top             =   1020
         Width           =   720
      End
      Begin VB.Shape Shape1 
         Height          =   585
         Left            =   570
         Top             =   3690
         Width           =   1605
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
         Left            =   720
         TabIndex        =   10
         Top             =   2310
         Width           =   345
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
         Left            =   1560
         TabIndex        =   9
         Top             =   2310
         Width           =   330
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
         Left            =   570
         TabIndex        =   8
         Top             =   1650
         Width           =   495
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
         Left            =   3120
         TabIndex        =   7
         Top             =   2340
         Width           =   990
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
         Left            =   2370
         TabIndex        =   6
         Top             =   2310
         Width           =   615
      End
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









Private Sub cmdExit_Click()

    Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdSAVE_Click()
Dim validation As Variant
              Adodc1.ConnectionString = strcn.Connection_String
              Adodc1.RecordSource = "Select user_id, user_name, user_password,user_type,shift_name From Security Where (User_Id = '" & frmMAIN.lbluser_id & "')"
              Adodc1.Refresh
              validation = Adodc1.Recordset!user_id
                
                Dim Conn As New ADODB.Connection
                Dim cmd As New ADODB.Command
                Dim RS As New ADODB.Recordset

                        Dim Param1 As New Parameter
                        Conn.Open strcn.Connection_String
    
                    Set cmd.ActiveConnection = Conn
                    cmd.CommandType = adCmdText
    
                   Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, validation)
                    cmd.Parameters.Append Param1 'validation
                    cmd.Properties("PLSQLRSet") = True
    
                     cmd.CommandText = "{CALL shift_validation(?)}"
    
                Debug.Print cmd.CommandText
    
                    Set RS = cmd.Execute
    

                cmd.Properties("PLSQLRSet") = False
                
          Adodc2.ConnectionString = strcn.Connection_String
          Adodc2.RecordSource = "Select * From user_validation"
          Adodc2.Refresh
        

        
             If Adodc2.Recordset!validation = 0 Then
             MsgBox "Your Working Time has been Expired", vbInformation, "Daffodil Software Ltd."
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
Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
 
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 5, frmExtraBed.txtRegNoExtraBed.Text)
    cmd.Parameters.Append Param1 'in_reg_no
    
   
    Set Param2 = cmd.CreateParameter("param2", adDouble, adParamInput, 10, Trim(txtCharge.Text))
    cmd.Parameters.Append Param2 'Bed_charge
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 2, "na")
    cmd.Parameters.Append Param3 'U_id default Sumon

    Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 12, DTPicker1.Value)
    cmd.Parameters.Append Param4 'START OR END DATE
    


   Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 5, "bo")
    cmd.Parameters.Append Param5 'booth
    
    
'    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Save_extra_Bed_info_indoor(?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set rs2 = cmd.Execute
    

'    cmd.Properties("PLSQLRSet") = False
    
    
End Sub

Private Sub Form_Load()
 txtInregExtraBed = frmExtraBed.txtRegNoExtraBed.Text
 
Conn2.ConnectionString = strcn.Connection_String
        Conn2.Open
        cmd.ActiveConnection = Conn2
        cmd.CommandType = adCmdText
        cmd.CommandText = "select pat_name,pat_guard_name,sex,age,doc_dept  From in_door_pat_info_main Where in_reg_no ='" & Trim(frmExtraBed.txtRegNoExtraBed.Text) & "'"
      
''        cmd.Properties("iRowsetChange") = True
'        cmd.Properties("updatability") = 7
        rs2.CursorLocation = adUseClient

        rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
        If rs2.RecordCount > 0 Then
         txtNameRelease = rs2!pat_name
         TxtAgeRelease = rs2!age
         comSexRelease.Text = rs2!sex
         comDepartmentRelease = rs2!doc_dept
         
rs2.Close
Conn2.Close
Else
 MsgBox "Invalid Registration No", vbInformation, "Warning:Daffodil Software Ltd"
 rs2.Close
 Conn2.Close
 Exit Sub
 Unload Me

End If

End Sub

