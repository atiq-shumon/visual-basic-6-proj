VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " VB-Oracle"
   ClientHeight    =   5400
   ClientLeft      =   5130
   ClientTop       =   2310
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3210
      Left            =   360
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1260
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   5662
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      ColumnHeaders   =   0   'False
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   17
      RowDividerStyle =   6
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2654.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3195.213
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtAddress 
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   4500
      TabIndex        =   2
      Top             =   765
      Width           =   3345
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   465
      Left            =   2745
      TabIndex        =   4
      Top             =   4725
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   465
      Left            =   5535
      TabIndex        =   6
      Top             =   4725
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   465
      Left            =   4185
      TabIndex        =   5
      Top             =   4725
      Width           =   1185
   End
   Begin VB.TextBox txtEmp_Name 
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   1440
      TabIndex        =   1
      Top             =   765
      Width           =   3075
   End
   Begin VB.TextBox txtEmp_Id 
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   360
      TabIndex        =   0
      Top             =   765
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   465
      Left            =   1350
      TabIndex        =   3
      Top             =   4725
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   9
      Top             =   405
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   285
      Index           =   1
      Left            =   2025
      TabIndex        =   8
      Top             =   450
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "Emp Id"
      Height          =   285
      Index           =   0
      Left            =   405
      TabIndex        =   7
      Top             =   405
      Width           =   1050
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objEmp_Info As New clsEmp_Info
Private objUtility As New clsUtility

Private Sub cmdSAVE_Click()

    With objEmp_Info
        .Connection_String = Con.Connection_String
        .Emp_id = txtEmp_Id
        .Emp_Name = txtEmp_Name
        .Address = txtAddress
        .Save
        .Show_Message
    End With

     With objEmp_Info
        txtEmp_Id = .Emp_id
        txtEmp_Name = .Emp_Name
        txtAddress = .Address
    End With
    
    cmdClear_Click

    Show_All

End Sub

Private Sub cmdClose_Click()

    objUtility.Close_Screen_Msg Me

End Sub

Private Sub cmdClear_Click()

    objUtility.Clear_Screen

End Sub

Private Sub cmdDelete_Click()

    If objUtility.Confirm_Delete_Msg(Me) = True Then

            With objEmp_Info
                .Connection_String = Con.Connection_String
                .Emp_id = txtEmp_Id
                .Delete
            End With
            
            objUtility.Clear_Screen
            
            Show_All
    
    End If
    
End Sub

Private Sub DataGrid1_Click()

    With DataGrid1
        txtEmp_Id = .Columns(0)
        txtEmp_Name = .Columns(1)
        txtAddress = .Columns(2)
    End With

End Sub


Private Sub Form_Load()
   objUtility.Screen_Position Me
   Show_All
End Sub


Private Sub txtEmp_Id_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then   ''txtEmp_Id <> "" And
    
        With objEmp_Info
            .Connection_String = Con.Connection_String
            .Emp_id = txtEmp_Id
            .GetX
         txtEmp_Id = .Emp_id
         txtEmp_Name = .Emp_Name
         txtAddress = .Address
        End With
    End If
End Sub
Private Sub Show_All()
    Dim RS As ADODB.Recordset
    
    With objEmp_Info
            .Connection_String = Con.Connection_String
            Set RS = .GetAll
    End With
    
       Set DataGrid1.DataSource = RS
       
       With DataGrid1
            .Columns(0).Width = 800
            .Columns(1).Width = 3000
            .Columns(2).Width = 3400
       End With
End Sub


