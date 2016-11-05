VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form13 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form13"
   ClientHeight    =   5130
   ClientLeft      =   2160
   ClientTop       =   1830
   ClientWidth     =   7890
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form13"
   LockControls    =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   4050
      ScaleHeight     =   600
      ScaleWidth      =   3480
      TabIndex        =   17
      Top             =   225
      Width           =   3480
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   600
         Index           =   10
         Left            =   0
         Top             =   0
         Width           =   3435
      End
      Begin VB.Label lblJob_type 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F7E6EB&
         Caption         =   "Job type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   45
         TabIndex        =   23
         Top             =   45
         Width           =   1110
      End
      Begin VB.Label lblDesignation 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F7E6EB&
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1170
         TabIndex        =   22
         Top             =   45
         Width           =   1110
      End
      Begin VB.Label lblDepartment 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F7E6EB&
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2295
         TabIndex        =   21
         Top             =   45
         Width           =   1110
      End
      Begin VB.Label lblBranch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F7E6EB&
         Caption         =   "Branch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   45
         TabIndex        =   20
         Top             =   315
         Width           =   1110
      End
      Begin VB.Label lblSection 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F7E6EB&
         Caption         =   "Section"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1170
         TabIndex        =   19
         Top             =   315
         Width           =   1110
      End
      Begin VB.Label lblRank 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F7E6EB&
         Caption         =   "Rank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2295
         TabIndex        =   18
         Top             =   315
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   45
      Top             =   4095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      ScaleHeight     =   375
      ScaleWidth      =   7125
      TabIndex        =   13
      Top             =   4590
      Width           =   7125
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3420
         Picture         =   "frmEmp_Info.frx":0000
         ScaleHeight     =   330
         ScaleMode       =   0  'User
         ScaleWidth      =   3309.763
         TabIndex        =   15
         Top             =   45
         Width           =   3315
         Begin VB.CommandButton cmdClose 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   2490
            Picture         =   "frmEmp_Info.frx":0327
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   0
            Width           =   870
         End
         Begin VB.CommandButton cmdClear 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   1650
            Picture         =   "frmEmp_Info.frx":0EA9
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   0
            Width           =   870
         End
         Begin VB.CommandButton cmdEdit 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   810
            Picture         =   "frmEmp_Info.frx":1A2B
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   0
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.CommandButton cmdDel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   -30
            Picture         =   "frmEmp_Info.frx":25AD
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.CommandButton cmdSave 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   810
            Picture         =   "frmEmp_Info.frx":312F
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   0
            Width           =   870
         End
      End
   End
   Begin VB.PictureBox picJob_type 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   0
      Picture         =   "frmEmp_Info.frx":3CB1
      ScaleHeight     =   4515
      ScaleWidth      =   7890
      TabIndex        =   9
      Top             =   0
      Width           =   7890
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   6480
         TabIndex        =   1
         Top             =   1125
         Width           =   960
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1950
         Left            =   585
         TabIndex        =   10
         Top             =   2475
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3440
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   6
         FormatLocked    =   -1  'True
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
         ColumnCount     =   4
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
         BeginProperty Column03 
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
               DividerStyle    =   6
               Locked          =   -1  'True
               ColumnWidth     =   1695.118
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   6
               Locked          =   -1  'True
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column02 
               DividerStyle    =   6
               Locked          =   -1  'True
               ColumnWidth     =   4020.095
            EndProperty
            BeginProperty Column03 
               DividerStyle    =   6
               Locked          =   -1  'True
               ColumnWidth     =   1530.142
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtTelephone 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1755
         TabIndex        =   3
         Top             =   2070
         Visible         =   0   'False
         Width           =   5595
      End
      Begin VB.TextBox txtDes 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1755
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1530
         Width           =   5685
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1755
         TabIndex        =   0
         Top             =   1125
         Width           =   3750
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   405
         Top             =   3870
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
      Begin VB.Image Image2 
         Height          =   480
         Left            =   7515
         Picture         =   "frmEmp_Info.frx":18F7F
         ToolTipText     =   "  Close  "
         Top             =   -180
         Width           =   480
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   945
         TabIndex        =   24
         Top             =   90
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   2025
         Index           =   9
         Left            =   540
         Top             =   2445
         Width           =   6930
      End
      Begin VB.Label lblCode 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5760
         TabIndex        =   16
         Top             =   1170
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   8
         Left            =   6435
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   7
         Left            =   1710
         Top             =   2025
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   465
         Index           =   6
         Left            =   1710
         Top             =   1485
         Width           =   5775
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   11
         Left            =   1710
         Top             =   1080
         Width           =   3840
      End
      Begin VB.Label lblTelephone 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   585
         TabIndex        =   14
         Top             =   2115
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lblDescription 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   585
         TabIndex        =   12
         Top             =   1665
         Width           =   795
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   585
         TabIndex        =   11
         Top             =   1170
         Width           =   840
      End
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim st    As String
Dim Stored_Procedure
Dim POP_Table As String
Dim Prev_Title As String

Private Sub cmdClear_Click()
txtTitle = ""
txtDes = ""
txtTelephone = ""
txtCode = ""
Timer1.Enabled = True
End Sub

Private Sub cmdClose_Click()
yes_no = MsgBox("Do you really want to close it?", vbYesNo + vbQuestion, "Daffodil PMIS")
    If yes_no = vbYes Then
        Unload Me
    Else
        Exit Sub
    End If
End Sub
Private Sub cmdDel_Click()
opr = "D"

cmdSave_Click

End Sub



Private Sub cmdEdit_Click()
cmdSave_Click
End Sub

Private Sub Form_Load()

Grid_Click (False), Form13

lblJob_type_Click
'lblTitle.Caption = "Job type"

DataGrid1.Height = 2085
DataGrid1.Top = 2070

POP_Table = "Job_Type"
populate_grd


DataGrid1.Columns(0).Caption = "Job type"

Stored_Procedure = "exec Job_Type_I_U_D '"


End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub lblBranch_Click()

lblHeading.Caption = "Branch"
lblTitle.Caption = "Branch"
lblDescription.Caption = "Address"

DataGrid1.Columns(2).Visible = True
DataGrid1.Columns(3).Visible = True
DataGrid1.Columns(0).Caption = "Branch"
DataGrid1.Columns(1).Caption = "Code"
DataGrid1.Columns(2).Caption = "Address"
DataGrid1.Columns(3).Caption = "Telephone"

txtCode.Visible = True
lblCode.Visible = True
Shape1(8).Visible = True

DataGrid1.Height = 1640
DataGrid1.Top = 2520

DataGrid1.Columns(0).Width = 1120
DataGrid1.Columns(1).Width = 510
DataGrid1.Columns(2).Width = 3680
DataGrid1.Columns(3).Width = 1330

Shape1(9).Height = 1700
Shape1(9).Top = 2490

Shape1(7).Visible = True

POP_Table = "Branch_Info"
populate_grd

''****************************************

Stored_Procedure = "exec  Branch_Info_I_U_D '"

cmdClear_Click

End Sub

Private Sub cmdSave_Click()

If txtTitle = "" Then
        MsgBox "Blank field not acceptable"
        txtTitle.SetFocus
        Exit Sub
    End If

    If txtCode = "" Then
        MsgBox "Blank field not acceptable"
        txtCode.SetFocus
        Exit Sub
    End If

con.ConnectionString = strcn.Connection
    con.Open

If lblHeading = "Job Type" Or lblHeading = "Designation" _
        Or lblHeading = "Section" Or lblHeading = "Department" Then

    
    cmd.CommandText = Stored_Procedure + opr + "','" + txtCode + "','" _
    + txtTitle + "','" _
    + Prev_Title + "','" _
    + txtDes + "','" + u_id + "'"
    
    
    
ElseIf lblHeading = "Branch" Then
   
    cmd.CommandText = Stored_Procedure + opr + "','" + txtCode + "','" _
    + ChkForQuote(txtTitle) + "','" _
    + ChkForQuote(Prev_Title) + "','" _
    + ChkForQuote(txtDes) + "','" + ChkForQuote(txtTelephone) + "','" + u_id + "'"
    

Else  ''''********  lblHeading = "Rank"  ***********

''--------------------------------------------

    cmd.CommandText = Stored_Procedure + opr + "','" + txtTitle + "','" _
    + Prev_Title + "','" _
    + txtDes + "','" + u_id + "'"

End If

cmd.ActiveConnection = con
    cmd.Execute
    con.Close
    populate_grd
    
Grid_Click (False), Form13
cmdClear_Click

End Sub
Private Sub DataGrid1_DblClick()
On Error Resume Next
Prev_Title = Adodc1.Recordset!Title

    If lblHeading = "Job Type" Or lblHeading = "Designation" _
        Or lblHeading = "Section" Or lblHeading = "Department" Then
        

            If Adodc1.Recordset.EOF Then Exit Sub
            
                 txtTitle = Adodc1.Recordset!Title
                 txtDes = Adodc1.Recordset!Description
                 txtCode = Adodc1.Recordset!Code
                 
                 txtTitle.Refresh
                 txtDes.Refresh
                 txtTelephone.Refresh
             
        
    ElseIf lblHeading = "Branch" Then
            
            If Adodc1.Recordset.EOF Then Exit Sub
                
                txtTitle = Adodc1.Recordset!Title
                txtDes = Adodc1.Recordset!Address
                txtTelephone = Adodc1.Recordset!Telephone
                
                txtTitle.Refresh
                txtDes.Refresh
                txtTelephone.Refresh
            
            
   Else    ''''**********  Or lblHeading = "Rank"  ****************
        
            If Adodc1.Recordset.EOF Then Exit Sub
            
                 txtTitle = Adodc1.Recordset!Title
                 txtDes = Adodc1.Recordset!Description
                
                 
                 txtTitle.Refresh
                 txtDes.Refresh
                 txtTelephone.Refresh
           
   End If
    
            Grid_Click (True), Form13
    
 End Sub

Public Sub populate_grd()


If lblHeading = "Job Type" Or lblHeading = "Designation" _
        Or lblHeading = "Section" Or lblHeading = "Department" Then
 
        Adodc1.ConnectionString = strcn.Connection
        
        Adodc1.RecordSource = "exec POP_Job_Setup'" + POP_Table + "'"
        Adodc1.Refresh
            If Adodc1.Recordset.RecordCount <> 0 Then
                If Adodc1.Recordset.RecordCount > 0 Then
                    Adodc1.Recordset.MoveLast
                End If
                Set DataGrid1.DataSource = Adodc1
                 
                DataGrid1.Columns(0).DataField = "title"
                DataGrid1.Columns(1).DataField = "Code"
                DataGrid1.Columns(2).DataField = "Description"
               
                DataGrid1.ReBind
                DataGrid1.Refresh
            End If
    
    
    
ElseIf lblHeading = "Branch" Then


        Adodc1.ConnectionString = strcn.Connection
        
        Adodc1.RecordSource = "exec POP_Job_Setup'" + POP_Table + "'"
        Adodc1.Refresh
            If Adodc1.Recordset.RecordCount <> 0 Then
                If Adodc1.Recordset.RecordCount > 0 Then
                    Adodc1.Recordset.MoveLast
                End If
                Set DataGrid1.DataSource = Adodc1
                 
                DataGrid1.Columns(0).DataField = "title"
                DataGrid1.Columns(1).DataField = "Code"
                DataGrid1.Columns(2).DataField = "Address"
                DataGrid1.Columns(3).DataField = "Telephone"
                
                DataGrid1.ReBind
                DataGrid1.Refresh
    
            End If
    
Else            '''*************** lblHeading = "Rank"   ********************



Adodc1.ConnectionString = strcn.Connection
        
        Adodc1.RecordSource = "exec POP_Job_Setup'" + POP_Table + "'"
        Adodc1.Refresh
            If Adodc1.Recordset.RecordCount <> 0 Then
                If Adodc1.Recordset.RecordCount > 0 Then
                    Adodc1.Recordset.MoveLast
                End If
                Set DataGrid1.DataSource = Adodc1
                 
                DataGrid1.Columns(0).DataField = "title"
                DataGrid1.Columns(1).DataField = "Description"
               
                DataGrid1.ReBind
                DataGrid1.Refresh
            End If


End If


End Sub

Private Sub lblDepartment_Click()

lblHeading.Caption = "Department"
lblTitle.Caption = "Department"
lblDescription.Caption = "Description"

DataGrid1.Columns(2).Visible = True
DataGrid1.Columns(3).Visible = False
DataGrid1.Columns(0).Caption = "Department"
DataGrid1.Columns(1).Caption = "Code"
DataGrid1.Columns(2).Caption = "Description"


DataGrid1.Height = 2085
DataGrid1.Top = 2070

DataGrid1.Columns(0).Width = 2095
DataGrid1.Columns(1).Width = 510
DataGrid1.Columns(2).Width = 3690
'DataGrid1.Columns(3).Width = 1530

Shape1(9).Height = 2145
Shape1(9).Top = 2040

Shape1(7).Visible = False

txtCode.Visible = True
lblCode.Visible = True
Shape1(8).Visible = True

POP_Table = "Department_Info"
populate_grd

Stored_Procedure = "exec  Department_I_U_D '"

cmdClear_Click

End Sub

Private Sub lblDesignation_Click()

lblHeading.Caption = "Designation"
lblTitle.Caption = "Designation"
lblDescription.Caption = "Description"

DataGrid1.Columns(2).Visible = True
DataGrid1.Columns(3).Visible = False
DataGrid1.Columns(0).Caption = "Designation"
DataGrid1.Columns(1).Caption = "Code"
DataGrid1.Columns(2).Caption = "Description"

DataGrid1.Height = 2085
DataGrid1.Top = 2070

DataGrid1.Columns(0).Width = 2095
DataGrid1.Columns(1).Width = 510
DataGrid1.Columns(2).Width = 3690
'DataGrid1.Columns(3).Width = 1530


Shape1(9).Height = 2145
Shape1(9).Top = 2040

Shape1(7).Visible = False

txtCode.Visible = True
lblCode.Visible = True
Shape1(8).Visible = True

POP_Table = "Job_Title"
populate_grd

Stored_Procedure = "exec  Job_Title_I_U_D '"

 cmdClear_Click

End Sub

Private Sub lblHeading_Change()

If lblHeading.Caption = "Job Type" Then
    lblJob_type.ForeColor = &HC000C0
Else
    lblJob_type.ForeColor = &H800000
End If

If lblHeading.Caption = "Designation" Then
    lblDesignation.ForeColor = &HC000C0
Else
lblDesignation.ForeColor = &H800000
End If

If lblHeading.Caption = "Department" Then
    lblDepartment.ForeColor = &HC000C0
Else
lblDepartment.ForeColor = &H800000
End If

If lblHeading.Caption = "Branch" Then
    lblBranch.ForeColor = &HC000C0
    lblTelephone.Visible = True
    txtTelephone.Visible = True
Else
    lblBranch.ForeColor = &H800000
    lblTelephone.Visible = False
    txtTelephone.Visible = False
    
End If

If lblHeading.Caption = "Section" Then
    lblSection.ForeColor = &HC000C0
Else
    lblSection.ForeColor = &H800000
End If

If lblHeading.Caption = "Rank" Then
    lblRank.ForeColor = &HC000C0
Else
    lblRank.ForeColor = &H800000
End If

End Sub

Private Sub lblJob_type_Click()

lblHeading.Caption = "Job Type"
lblTitle.Caption = "Job type"

DataGrid1.Columns(2).Visible = True
DataGrid1.Columns(3).Visible = False
lblDescription.Caption = "Description"
DataGrid1.Columns(0).Caption = "Job type"
DataGrid1.Columns(1).Caption = "Code"
DataGrid1.Columns(2).Caption = "Description"


DataGrid1.Height = 2085
DataGrid1.Top = 2070
''------------------
DataGrid1.Columns(0).Width = 2095
DataGrid1.Columns(1).Width = 510
DataGrid1.Columns(2).Width = 3690
'DataGrid1.Columns(3).Width = 1530
''------------------

Shape1(9).Height = 2145
Shape1(9).Top = 2040

Shape1(7).Visible = False

txtCode.Visible = True
lblCode.Visible = True
Shape1(8).Visible = True

POP_Table = "Job_Type"
populate_grd

Stored_Procedure = "exec Job_Type_I_U_D '"

 cmdClear_Click
 

End Sub

Private Sub lblRank_Click()

lblHeading.Caption = "Rank"
lblTitle.Caption = "Rank"

DataGrid1.Columns(2).Visible = False
DataGrid1.Columns(3).Visible = False
lblDescription.Caption = "Description"
DataGrid1.Columns(0).Caption = "Rank"
DataGrid1.Columns(1).Caption = "Description"


DataGrid1.Height = 2085
DataGrid1.Top = 2070

DataGrid1.Columns(0).Width = 1120
DataGrid1.Columns(1).Width = 5180

Shape1(9).Height = 2145
Shape1(9).Top = 2040

Shape1(7).Visible = False

txtCode.Visible = False
lblCode.Visible = False
Shape1(8).Visible = False

POP_Table = "Rank"
populate_grd

Stored_Procedure = "exec Rank_I_U_D '"

cmdClear_Click

End Sub

Private Sub lblSection_Click()

lblHeading.Caption = "Section"
lblTitle.Caption = "Section"

lblDescription.Caption = "Description"

DataGrid1.Columns(2).Visible = True
DataGrid1.Columns(3).Visible = False
DataGrid1.Columns(0).Caption = "Section"
DataGrid1.Columns(1).Caption = "Code"
DataGrid1.Columns(2).Caption = "Description"

DataGrid1.Height = 2085
DataGrid1.Top = 2070

DataGrid1.Columns(0).Width = 2095
DataGrid1.Columns(1).Width = 510
DataGrid1.Columns(2).Width = 3980
'DataGrid1.Columns(3).Width = 1530

Shape1(9).Height = 2145
Shape1(9).Top = 2040

Shape1(7).Visible = False

txtCode.Visible = True
lblCode.Visible = True
Shape1(8).Visible = True

POP_Table = "Sec_Info"
populate_grd

Stored_Procedure = "exec Sec_info_I_U_D '"

cmdClear_Click

End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then Grid_Click (False), Form13
    
    txtTitle.SetFocus
        
    cmdClear_Click
        
End Sub
Private Sub Timer1_Timer()

    txtTitle.SetFocus
    Timer1.Enabled = False
End Sub
