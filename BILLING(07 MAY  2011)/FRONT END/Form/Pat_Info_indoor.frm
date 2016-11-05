VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Pat_Info_indoor 
   BackColor       =   &H00C9AD8F&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8625
   ClientLeft      =   -105
   ClientTop       =   435
   ClientWidth     =   11880
   FillColor       =   &H007DABD0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo5 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Pat_Info_indoor.frx":0000
      Left            =   5490
      List            =   "Pat_Info_indoor.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   4005
      Width           =   1440
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6930
      TabIndex        =   48
      Top             =   2160
      Width           =   765
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3465
      TabIndex        =   47
      Top             =   2160
      Width           =   360
   End
   Begin VB.TextBox txtReg 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3465
      TabIndex        =   46
      Top             =   1575
      Width           =   2715
   End
   Begin VB.TextBox nbrDisc_Per 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   9315
      TabIndex        =   45
      Top             =   4545
      Width           =   480
   End
   Begin VB.TextBox nbrDisc 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   7920
      TabIndex        =   44
      Top             =   4545
      Width           =   1200
   End
   Begin VB.TextBox nbrDue 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox nbrTotal 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5580
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1170
      TabIndex        =   40
      Top             =   3510
      Width           =   5715
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   6510
      MaxLength       =   10
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
      Top             =   2610
      Width           =   3300
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8595
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtAddr 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1170
      TabIndex        =   37
      Top             =   3060
      Width           =   8640
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4410
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   2160
      Width           =   630
   End
   Begin VB.TextBox nbrAdv 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4005
      Width           =   1905
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1170
      MaxLength       =   10
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
      Top             =   2565
      Width           =   4065
   End
   Begin VB.TextBox txtDummy_Pat_ID 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1170
      MaxLength       =   10
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
      Top             =   1575
      Width           =   1350
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2160
      Width           =   1350
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5805
      TabIndex        =   30
      Top             =   2160
      Width           =   360
   End
   Begin VB.TextBox txtAge 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1170
      MaxLength       =   17
      TabIndex        =   29
      Top             =   4005
      Width           =   555
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Pat_Info_indoor.frx":0042
      Left            =   2475
      List            =   "Pat_Info_indoor.frx":004C
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   4005
      Width           =   1260
   End
   Begin VB.TextBox txtPhone 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   7920
      TabIndex        =   27
      Top             =   3510
      Width           =   1890
   End
   Begin VB.CommandButton cmdDelete 
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
      Height          =   420
      Left            =   990
      Picture         =   "Pat_Info_indoor.frx":005E
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Delete"
      Top             =   6165
      Width           =   510
   End
   Begin VB.CommandButton cmdADD 
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
      Height          =   420
      Left            =   495
      Picture         =   "Pat_Info_indoor.frx":0B98
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "New"
      Top             =   6165
      Width           =   510
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
      Height          =   420
      Left            =   1500
      Picture         =   "Pat_Info_indoor.frx":1202
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Preview"
      Top             =   6165
      Width           =   510
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
      Height          =   420
      Left            =   0
      Picture         =   "Pat_Info_indoor.frx":186C
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Save"
      Top             =   6165
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
      Height          =   420
      Left            =   2010
      Picture         =   "Pat_Info_indoor.frx":1ED6
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exit"
      Top             =   6165
      Width           =   510
   End
   Begin MSComCtl2.DTPicker DT_TM 
      Height          =   330
      Left            =   8235
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Delevary Time"
      Top             =   1620
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "HH:MM:SS"
      Format          =   22675458
      UpDown          =   -1  'True
      CurrentDate     =   37163
   End
   Begin MSComCtl2.DTPicker Dt 
      Height          =   330
      Left            =   6930
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1620
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   582
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarForeColor=   16711680
      CalendarTitleBackColor=   16777215
      CalendarTitleForeColor=   49152
      Format          =   22675457
      CurrentDate     =   37114
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Religion"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4815
      TabIndex        =   51
      Top             =   4050
      Width           =   570
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Charge"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6345
      TabIndex        =   49
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type No"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2790
      TabIndex        =   26
      Top             =   2205
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6165
      TabIndex        =   20
      Top             =   3105
      Width           =   420
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7290
      TabIndex        =   19
      Top             =   3555
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   450
      TabIndex        =   18
      Top             =   3600
      Width           =   570
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guradian Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5355
      TabIndex        =   17
      Top             =   2655
      Width           =   1110
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7740
      TabIndex        =   16
      Top             =   2205
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   450
      TabIndex        =   15
      Top             =   1620
      Width           =   795
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Bed"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5085
      TabIndex        =   14
      Top             =   2205
      Width           =   690
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bed No."
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3870
      TabIndex        =   13
      Top             =   2205
      Width           =   675
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bed Type"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   405
      TabIndex        =   12
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reg No"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2835
      TabIndex        =   11
      Top             =   1620
      Width           =   600
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Indoor Patient Information "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   525
      Left            =   2610
      TabIndex        =   10
      Top             =   360
      Width           =   4920
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   450
      TabIndex        =   3
      Top             =   3150
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   450
      TabIndex        =   2
      Top             =   2610
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2070
      TabIndex        =   1
      Top             =   4050
      Width           =   315
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6480
      TabIndex        =   0
      Top             =   1665
      Width           =   435
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   450
      TabIndex        =   9
      Top             =   4050
      Width           =   285
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advance"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7155
      TabIndex        =   8
      Top             =   4050
      Width           =   645
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9135
      TabIndex        =   7
      Top             =   4635
      Width           =   120
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7110
      TabIndex        =   6
      Top             =   4590
      Width           =   630
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7245
      TabIndex        =   5
      Top             =   5130
      Width           =   300
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7335
      TabIndex        =   4
      Top             =   5670
      Width           =   420
   End
End
Attribute VB_Name = "Pat_Info_indoor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub clear()

    txtPat_ID1 = ""
    txtDummy_Pat_ID = ""
    txtPat_ID = ""
    txtPat_Name = ""
    ComSex = "Male"
    txtAge = ""
    txtRefer_Code = ""
    txtAddr = ""
    txtPhone = ""
    txtFax = ""
    txtEmail = ""
    Dt.Value = Now

    DT_TM.Value = Now
    txtDoc_Name = ""
    txtDegree = ""
    txtDoc_Addr = ""
    nbrAdv = 0
    nbrDisc = 0

    nbrTot_Disc = 0

    nbrDisc_Per = 0
    nbrDue = ""
    nbrNet_Amount = 0
    nbrTest_Rate = 0
    nbrTotal = ""
    nbrCollect_Fee = 0
    nbrTotCollect_Fee = 0
   
    nbrDisc.Locked = False
    nbrAdv_Pay = 0

    nbrTot_Test = ""


 


    txtPat_ID1 = ""
    txtPat_ID = ""
    txtDummy_Pat_ID = ""
    txtPat_ID1.Visible = True
   
    txtReg.SetFocus

End Sub

'Private Sub cmdPatOld_Click()
'
'    Temp_rst
''    txtPat_ID1 = ""
''    txtDummy_Pat_ID = ""
''    txtPat_ID = ""
'    txtPat_Name = ""
'    ComSex = "Male"
'    txtAge = ""
'    txtRefer_Code = ""
'    txtAddr = ""
'    txtPhone = ""
'    txtFax = ""
'    txtEmail = ""
'    Dt.Value = Now
'    Delv_Dt.Value = Now
'    DT_TM.Value = Now
'    txtDoc_Name = ""
'    txtDegree = ""
'    txtDoc_Addr = ""
'    nbrAdv = 0
'    nbrDisc = 0
'
'    nbrTot_Disc = 0
'
'    nbrDisc_Per = 0
'    nbrDue = ""
'    nbrNet_Amount = 0
'    nbrTest_Rate = 0
'    nbrTotal = ""
'    nbrCollect_Fee = 0
'    nbrTotCollect_Fee = 0
'    nbrAdv_Pay.Locked = False
'    nbrDisc.Locked = False
'    nbrAdv_Pay = 0
'    ChkPaid.Value = 0
'    nbrTot_Test = ""
'
'    Chkrefer_type.Value = 0
'
'    nbrCollect_Fee.Locked = False
'
'
'
'    txtPat_ID1 = ""
'    txtPat_ID = ""
'    txtDummy_Pat_ID = ""
'    txtPat_ID1.Visible = False
'    txtPat_ID.Visible = True
'    txtPat_ID.SetFocus
'
'End Sub
'
'Private Sub cmdPreview_Click()
'    If StPat_ID = "" And txtPat_ID = "" Then Exit Sub
'
'    CRViewer1_MODE = 14
'    Viewer.Show vbModal
'End Sub
'
'Private Sub CmdPreview_GotFocus()
'cmdPreview.BackColor = &HC0FFFF
'End Sub
'
'Private Sub CmdPreview_LostFocus()
'
'cmdPreview.BackColor = vbWhite
'
'End Sub
'
'Private Sub cmdPrint_Click()
'    If StPat_ID = "" And txtPat_ID = "" Then Exit Sub
'        '==========direct print==========================
'            If frmPatient_Info.txtPat_ID = "" Then
'            StrPat_ID = StPat_ID
'            Else
'            StrPat_ID = frmPatient_Info.txtPat_ID
'            End If
'
'            Dim Report14 As New Pat_Info1
'            Report14.DiscardSavedData
'            RS.Open "exec Rpt_pat_info '" + StrPat_ID + "'", strcn.Connection
'            Report14.Database.SetDataSource RS
'
'            Report14.PrintOut
'            RS.Close
'    '====================================
'End Sub
'
'Private Sub cmdPrint_GotFocus()
'
'cmdPrint.BackColor = &HC0FFFF
'
'End Sub
'
'Private Sub cmdPrint_LostFocus()
'
'cmdPrint.BackColor = vbWhite
'
'End Sub
'
'Private Sub cmdSAVE_Click()
'
'Strpat_id1 = "0"
'    'MsgBox BoothN
'
'    Dt.Value = Now
'    DT_TM.Value = Now
'
'    If Trim(txtPat_Name) = "" Then
'        MsgBox "Paitent Name Mandatory"
'        txtPat_Name.SetFocus
'        Exit Sub
'    End If
'
'    If Trim(txtDoc_Name) = "" Then
'        MsgBox "Doctor's name Mandatory"
'        txtRefer_Code = ""
'        txtRefer_Code.SetFocus
'        Exit Sub
'    End If
'
'
'    If Trim(nbrTotal_Amt) = "" Or Val(nbrTotal_Amt) = 0 Then
'        MsgBox "Test Mandatory"
'        txtM_Code.SetFocus
'        Exit Sub
'    End If
'
'
'    'temp_rst 'RECORDSET
'    Adodc1.ConnectionString = strcn.Connection
'    Adodc1.RecordSource = "select * from Pat_Info_main where pat_id='" & Trim(txtPat_ID.Text) & "'"
'    Adodc1.Refresh
'    If Adodc1.Recordset.RecordCount > 0 Then
'    'MsgBox u_id
'
'    If u_id <> "md" Then
'        MsgBox "If you want to any change you should contact to Managing Director.., Your ID saved..", vbCritical
'        Exit Sub
'    End If
'
'
'
'       StPat_ID = txtPat_ID 'TAKEN PAT_ID FOR PRINT
'
'
'       Strpat_id1 = DummyPat_ID1
'
'       If txtPat_ID1.Text <> "" Then
'             If StPat_Type1 <> Chkrefer_type.Value Then
'                 Make_Pat_ID1_U
''                 MsgBox Strpat_id1
'            End If
'
'       End If
'
'
'       'If Chkrefer_type = strRefer_Type1 Then
'        '    Strpat_id1 = DummyPat_ID1
'       'End If
'
'
'
'       UpdPat_Info_Main
'       Delete_Pat_Info_Sub1
'       InsPat_Info_Sub1_U 'after delete data then INSERT
'       InsPat_Info_Sub2_T1
'       nbrAdv_Pay.Locked = False
'       'UpdPat_Info_Sub3
'       InsPat_Info_Sub3
'
'       Search_Refer_Code 'search again refer_code for update refer_code/delete from doctor_info_new
'       Del_New_Doc
'
'    Else
'
'        Make_Pat_ID1
'
'        Dt.Value = Now
'        DT_TM.Value = Now
'
'        InsPat_Info_Main
'
'    '///////SEARCH PATIENT ID for insert another table//////////////////////
'        Adodc14.ConnectionString = strcn.Connection
'        Adodc14.RecordSource = "exec test_Info_SELECT 2,'" & BoothN & "'"
'        Adodc14.Refresh
'        If Adodc14.Recordset.RecordCount > 0 Then
'        StPat_ID = Adodc14.Recordset!pat_id
'        End If
'    '///////END////////////////////////////////////////////
'
'        InsPat_Info_Sub1
'       ''''to insert into PAT_INFO_SUB2'''''''''
'        If txtPat_ID = "" Then
'            InsPat_Info_Sub2_T
'            nbrAdv_Pay.Locked = False
'        End If
'    ''''''''''''''''end'''''''''''''''''''''''''
'        InsPat_Info_Sub3
'    ',,,,,,,,,for select,delete and insert into doctor_info_new,,,,,,,,,,,,,,,
'        InsDoc_info_new
'    ',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
'
'    End If
'
'    Temp_rst
'
'    txtDummy_Pat_ID = ""
'    txtPat_ID1.Text = ""
''    txtPat_ID1.Visible = False
'    txtPat_ID.Text = ""
'    txtPat_ID.Visible = False
'    txtPat_Name = ""
'    ComSex = "Male"
'    txtAge = ""
'    txtRefer_Code = ""
'    txtAddr = ""
'    txtPhone = ""
'    txtFax = ""
'    txtEmail = ""
'    Dt.Value = Now
'    'Delv_Dt.value = Now
'    DT_TM.Value = Now
'    txtDoc_Name = ""
'    txtDegree = ""
'    txtDoc_Addr = ""
'    nbrTot_Test = ""
'    nbrTotal = ""
'    nbrTotal_Amt = ""
'    nbrDisc = 0
'    nbrTot_Disc = 0
'    nbrDisc_Per = 0
'    nbrNet_Amount = 0
'    nbrNet_Amount = ""
'    nbrVAT_Amt = 0
'    nbrTotal_Amt = ""
'    nbrAdv.Text = 0
'    nbrAdv_Pay = 0
'    nbrTotCollect_Fee.Text = 0
'    nbrCollect_Fee.Text = 0
'    nbrDue = ""
'    ChkPaid.Value = 0
'    Chkrefer_type.Value = 0
'    '---------
'
'    nbrCollect_Fee.Locked = False
'    nbrDisc.Locked = False
'    cmdPrint.SetFocus
'
'End Sub
'Private Sub InsPat_Info_Main()
'
'
'    InsD_TM 'for insert current date and time
'    Sel_Refer_Type 'for select REFERENCE TYPE
'
'
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'    cmd.CommandText = "exec pro_PAT_INFO_MAIN 'I','" + ChkForQuote(txtPat_Name.Text) + "','" + ComSex.Text + "','" + txtAge.Text + _
'    "','" + txtRefer_Code.Text + "','" + ChkForQuote(txtAddr.Text) + "','" + txtPhone.Text + _
'    "','" + txtFax.Text + "','" + txtEmail.Text + "','" + u_id + _
'    "','" + CDate_TM + _
'    "'," + nbrVAT_Per.Text + _
'    "," + nbrVAT_Amt.Text + _
'    ",'" + BoothN + "','" + Format(Dt, "yyyy-mm-dd") + _
'    "','" + CDate_TM + _
'    "','" + StrRefer_Type + _
'    "','" + Strpat_id1 + _
'    "','" + Strpat_MY + "'"
'
'    cmd.Execute
''    MsgBox cmd.Execute
'    Con.Close
'
'End Sub
'Private Sub UpdPat_Info_Main()
'
''MsgBox Strpat_id1
''MsgBox Strpat_MY
'
'      InsD_TM ' for insert current date and time
'
'      Sel_Refer_Type 'for select REFERENCE TYPE
'
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'    cmd.CommandText = "exec pro_PAT_INFO_MAIN_UD 'U','" + txtPat_ID.Text + _
'    "','" + ChkForQuote(txtPat_Name) + "','" + ChkForQuote(ComSex) + "','" + ChkForQuote(txtAge) + "','" + txtRefer_Code + _
'    "','" + ChkForQuote(txtAddr) + "','" + txtPhone + "','" + txtFax + _
'    "','" + txtEmail + "','" + u_id + _
'    "','" + CDate_TM + _
'    "'," + nbrVAT_Per + "," + nbrVAT_Amt + ",'" + BoothN + _
'    "','" + Format(CDate_TM3, "yyyy-mm-dd hh:mm") + _
'    "','" + Format(CDate_TM6, "yyyy-mm-dd hh:mm") + _
'    "','" + StrRefer_Type + _
'    "','" + Strpat_id1 + _
'    "','" + Strpat_MY + "'"
'
'    cmd.Execute
'    Con.Close
'End Sub
'Private Sub InsPat_Info_Sub1()
'
'
'    Temp_Table.MoveFirst
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'
'    While Temp_Table.EOF = False
'          cmd.CommandText = "exec pro_PAT_INFO_SUB1 'I'," + StPat_ID + _
'              ",'" + Temp_Table!m_code + _
'              "','" + Temp_Table!s_code + _
'              "'," + CStr(Temp_Table!test_rate) + _
'              ",'" + Temp_Table!Delv_DTM + _
'              "','" + Temp_Table!Type + _
'              "','" + u_id + _
'              "','" + CDate_TM + _
'              "','" + Format(Dt, "yyyy-mm-dd") + _
'              "','" + CDate_TM + _
'              "','" + nbrUnique_id + "'"
''             Debug.Print cmd.CommandText
''             MsgBox cmd.CommandText
'              cmd.Execute
'              Temp_Table.MoveNext
'
'    Wend
'    Con.Close
'
'End Sub
'Private Sub InsPat_Info_Sub1_U()
'
'
'    If txtPat_ID = "" Then Exit Sub
'    Temp_Table.MoveFirst
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'
'    While Temp_Table.EOF = False
'          cmd.CommandText = "exec pro_PAT_INFO_SUB1 'I'," + txtPat_ID + _
'              ",'" + Temp_Table!m_code + _
'              "','" + Temp_Table!s_code + _
'              "'," + CStr(Temp_Table!test_rate) + _
'              ",'" + Format(Temp_Table!Delv_DTM, "yyyy-mm-dd hh:mm") + _
'              "','" + Temp_Table!Type + _
'              "','" + u_id + _
'              "','" + CDate_TM + _
'              "','" + Format(CDate_TM2, "yyyy-mm-dd hh:mm") + _
'              "','" + Format(CDate_TM7, "yyyy-mm-dd hh:mm") + _
'              "','" + nbrUnique_id + "'"
''             Debug.Print cmd.CommandText
'             'MsgBox cmd.CommandText
'              cmd.Execute
'              Temp_Table.MoveNext
'    Wend
'    Con.Close
'
'End Sub
'Private Sub Delete_Pat_Info_Sub1()
'    If txtPat_ID = "" Then Exit Sub
'
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'    cmd.CommandText = "exec Pat_Info_Sub1_Delete 1,'" + Trim(txtPat_ID) + "'"
'    cmd.Execute
'    Con.Close
'
'End Sub
'Private Sub DeletePat_Info_Sub1()
'
''    Temp_Table.MoveFirst
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'
''    While Temp_Table.EOF = False
'          cmd.CommandText = "exec Pat_Info_Sub1_Delete1 1,'" + Trim(nbrUnique_id) + "'"
'
'              cmd.Execute
'    Con.Close
'    txtM_Code = ""
'    txtS_Code = ""
'    txtS_Name = ""
'    nbrRate = 0
'    nbrUnique_id = ""
'End Sub
'Private Sub InsPat_Info_Sub3()
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'    cmd.CommandText = "exec pro_PAT_INFO_SUB3 'I'," + StPat_ID + _
'    "," + nbrDisc + "," + ChkPaidVal + ",'" + u_id + _
'    "','" + CDate_TM + _
'    "','" + Format(Dt, "yyyy-mm-dd") + _
'    "','" + CDate_TM + "'"
'    cmd.Execute
'    Con.Close
'End Sub
'Private Sub UpdPat_Info_Sub3()
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'    cmd.CommandText = "exec pro_PAT_INFO_SUB3 'U'," + txtPat_ID.Text + _
'    "," + nbrDisc + "," + ChkPaidVal + ",'" + u_id + _
'    "','" + CDate_TM + _
'    "','" + Format(CDate_TM4, "yyyy-mm-dd") + _
'    "','" + Format(CDate_TM9, "yyyy-mm-dd hh:mm") + "'"
'    cmd.Execute
'    Con.Close
'End Sub
'Private Sub cmdSave_GotFocus()
'    cmdSAVE.BackColor = &HC0FFFF
'End Sub
'Private Sub cmdSave_LostFocus()
'    cmdSAVE.BackColor = vbWhite
'End Sub
'
'Private Sub cmdSearch_Click()
'
'If u_id <> "md" Then Exit Sub
'
'
'    Dim StrMMS As String
'    StrMMS = MsgBox("Do you want Update New Patient ?", vbQuestion + vbYesNo)
'    If StrMMS = vbYes Then
'        cmdPatNew.Visible = True
'        cmdPatOld.Visible = False
'    Else
'        cmdPatNew.Visible = False
'        cmdPatOld.Visible = True
'    End If
'
'End Sub
'
'Private Sub ComSex_GotFocus()
'    ComSex.BackColor = &HFFFFC0
'
'End Sub
'
'Private Sub ComSex_LostFocus()
'ComSex.BackColor = vbWhite
'End Sub
'
'Private Sub Cr_DT_TM_Click()
'    Dt.Value = Now
'    DT_TM.Value = Now
'End Sub
'
'Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
''    Sum_Rate
'    nbrTot_Test = Rate_Tot
'End Sub
'Private Sub DataGrid1_DblClick()
'
'    If Temp_Table.EOF = True Then Exit Sub
'
'    txtM_Code = Temp_Table!m_code
'    txtS_Code = Temp_Table!s_code
'    txtS_Name = Temp_Table!s_name
'    nbrTest_Rate = Temp_Table!test_rate
''    nbrUnique_id = Temp_Table_Helper!unique_id
'    Select_Unique_ID
'End Sub
'Private Sub Delv_Dt_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = 13 Then
'        SendKeys Chr(9)
'    End If
'
'End Sub
'
'Private Sub Delv_TM_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = 13 Then
'        SendKeys Chr(9)
'    End If
'
'End Sub
'
'Private Sub Delv_TM_LostFocus()
'    If Len(Trim(txtM_Code.Text)) = 0 Then Exit Sub
'    If Len(Trim(txtS_Code.Text)) = 0 Then Exit Sub
'
''    Search_Type ' Search Type from table "Test_Info_Sub"
'
'    '----------------check--------
'    If Trim(nbrTest_Rate) = 0 Then
'        MsgBox "Rate mandatory"
'        nbrTest_Rate.SetFocus
'        Exit Sub
'    End If
'
'    If Trim(txtS_Name.Text) = "" Then
'        MsgBox "Test Name mandatory"
'        txtM_Code.Text = ""
'        txtM_Code.SetFocus
'        Exit Sub
'    End If
'
'Dim Check As Integer
'Check = 0
'If Temp_Table.RecordCount > 0 Then
'    Temp_Table.MoveFirst
'
'        While Temp_Table.EOF = False
'            If Temp_Table!m_code = txtM_Code And Temp_Table!s_code = txtS_Code Then
'                Check = 1
'            End If
'    Temp_Table.MoveNext
'        Wend
'
'    If Check = 1 Then
'        MsgBox "This Group Name and Test Name already exists"
'        Check = 0
'        txtS_Code.SetFocus
'        Exit Sub
'    End If
''    Temp_Table.MoveFirst
'End If
'
''--------------end check-----
'
''++++++for insert Delivary Date and Time++++++++++++++
'
'StrDATE = Trim(Format(Delv_Dt, "yyyy-mm-dd"))
'StrTIME = Trim(Format(Delv_TM, "hh:mm"))
'
'Date_TM = StrDATE + Space(1) + StrTIME
''MsgBox Date_TM
''++++++++++end+++++++++++++++++++++++++++++++++++++++
'
''+++to insert into TEMPORARY RECORDSET++
'
'        Temp_Table.AddNew
'        Temp_Table!m_code = txtM_Code
'        Temp_Table!s_code = txtS_Code
'        Temp_Table!s_name = txtS_Name
'        Temp_Table!test_rate = nbrTest_Rate
'        Temp_Table!Delv_DTM = Date_TM
'        Temp_Table!Type = txtType
'
'
'        'Search_Type ' Search Type from table "Test_Info_Sub"
'
'        DataGrid1.Refresh
''++++++++++for count TOTAL_RATE from Temp_Table+++++++++
'        Total_Rate = 0
'        Temp_Table.MoveFirst
'        While Temp_Table.EOF = False
'                Total_Rate = Total_Rate + Temp_Table!test_rate
'
'        Temp_Table.MoveNext
'        Wend
'        nbrTotal = Val(Total_Rate)
''++++++++++End count TOTAL_RATE from Temp_Table+++++++++
'
''=========count total test=============================
'        Total_Test = 0
'        Total_Test = Temp_Table.RecordCount
'        Me.nbrTot_Test = Total_Test
''======================================================
'
'
''END ++++++++++++++++++++++++++++++++
'        txtM_Code = ""
'        txtS_Code = ""
'        txtS_Name = ""
'        nbrTest_Rate = 0
'        txtType.Text = ""
'        txtM_Code.SetFocus
'
'    DataGrid1.Columns(0).Width = 450.1418
'    DataGrid1.Columns(1).Width = 810.1418
'    DataGrid1.Columns(2).Width = 3825.071
'    DataGrid1.Columns(3).Width = 1110.047
'    DataGrid1.Columns(4).Width = 1900
'    DataGrid1.Columns(5).Width = 600
'
'    ChkPaid.Value = 0
'
'    nbrVAT_Amt = Val(nbrTotal) * Val(nbrVAT_Per) / 100 'for show VAT Amount
'
'
'
'
'
'End Sub
'
'Private Sub Form_DblClick()
'    If CmdSearch.Visible = False Then
'        CmdSearch.Visible = True
'    End If
'End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'       SendKeys Chr(9)
'    End If
'    If KeyAscii = 27 Then
'        Unload Me
'    End If
'End Sub
'
'Private Sub Form_Load()
'    'Test_List_Mode = "frmPatient_Info_M" 'mode for show 'TEST NAME LIST'
'    Temp_rst
'    nbrAdv_Pay = 0
'    nbrDisc = 0
'    nbrTot_Disc.Text = 0
'    ChkPaidVal = 0
'    nbrTotal = 0
'    nbrTotCollect_Fee.Text = 0
'    nbrCollect_Fee.Text = 0
'    'Locate_Booth
'
'
'    Delv_TM = Now
'
''    Doc_List_MODE = "frmPatient_Info"
'
''    Val(StPat_ID) = Null
'    ChkPaid.Value = 0
'    Dt.Value = Now
'    Delv_Dt.Value = Now
'    DT_TM.Value = Now
'    ComSex = "Male"
'    temp_open = "0"
'    Flush_VAT_Per
'End Sub
'Private Sub Form_Unload(Cancel As Integer)
'    Set Temp_Table = Nothing
'End Sub
'
'Private Sub Label26_Click()
'
'End Sub
'
'Private Sub nbrAdv_Change()
''    nbrTot_Disc = Val(nbrTot_Disc) + Val(nbrDisc)
'
'    nbrDue = Val(nbrNet_Amount) - Val(nbrAdv)
'    '--for auto select PAID check box
'    If Val(nbrNet_Amount) = 0 Then Exit Sub
'    If Val(nbrAdv) = 0 Then Exit Sub
'    If Val(nbrNet_Amount) = Val(nbrAdv) Then
'        ChkPaid.Value = 1
'    Else
'       If Val(nbrTotal_Amt) + Val(nbrTotCollect_Fee) = Val(nbrTot_Disc) Then
'       ChkPaid.Value = 1
'       Else
'       ChkPaid.Value = 0
'       End If
'    End If
'End Sub
'
'Private Sub nbrAdv_GotFocus()
'
'nbrAdv.BackColor = &HFFFFC0
'
'End Sub
'
'Private Sub nbrAdv_LostFocus()
'    nbrAdv.BackColor = vbWhite
'End Sub
'
'Private Sub nbrAdv_Pay_Change()
'    If Not IsNumeric(nbrAdv_Pay.Text) Then
'        MsgBox "Only Numaric value allow"
'        nbrAdv_Pay = ""
'        nbrAdv_Pay.SelStart = 0
'        nbrAdv_Pay.SelLength = Len(nbrAdv_Pay)
'        nbrAdv_Pay.SetFocus
'    End If
'End Sub
'
'Private Sub nbrAdv_Pay_GotFocus()
'    nbrAdv_Pay.BackColor = &HFFFFC0
'
'    nbrAdv_Pay.SelStart = 0
'    nbrAdv_Pay.SelLength = Len(nbrAdv_Pay)
'End Sub
'
'Private Sub nbrAdv_Pay_LostFocus()
''    If txtPat_ID = "" And StPat_ID = "" Then Exit Sub
'
'    nbrAdv_Pay.BackColor = vbWhite
'
'    If Trim(nbrAdv_Pay.Text) = "" Or Val(nbrAdv_Pay) = 0 Then Exit Sub
'
'
'    If Val(nbrAdv_Pay) > Val(nbrDue) Then
'        MsgBox "It is Impossible to pay more then DUE", vbInformation
'        nbrAdv_Pay.Text = 0
'        nbrAdv_Pay.SetFocus
'        Exit Sub
'    End If
'
'    Dim Strmsg As String
'    Strmsg = MsgBox("The Paitent's present DUE is  " + nbrDue + " and PAID " + nbrAdv + "  Do you want to pay more  " + nbrAdv_Pay + "", vbQuestion + vbYesNo)
'        If Strmsg = vbYes Then
'          ' If txtPat_ID = "" Then
''           InsPat_Info_Sub2
'           nbrAdv_Pay.Locked = True
'           nbrAdv = Val(nbrAdv) + Val(nbrAdv_Pay)
'          ' End If
'          cmdSAVE.SetFocus
'        Else
'            nbrAdv_Pay.Text = "0"
'            Exit Sub
'        End If
'
'
'End Sub
'
'Private Sub nbrCollect_Fee_Change()
'    If Not IsNumeric(nbrCollect_Fee.Text) Then
'        MsgBox "Only Numaric value allow"
'        nbrCollect_Fee = 0
'        nbrCollect_Fee.SelStart = 0
'        nbrCollect_Fee.SelLength = Len(nbrCollect_Fee)
'        nbrCollect_Fee.SetFocus
'    End If
'
''    nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrDisc) + Val(nbrTotCollect_Fee.Text)
'    'nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrDisc) + Val(nbrTotCollect_Fee.Text)
'    'nbrNet_Amount = Val(nbrTotal_Amt) - DblDisc + Val(nbrTotCollect_Fee.Text)
'End Sub
'Private Sub nbrCollect_GotFocus()
'    nbrCollect_Fee.SelStart = 0
'    nbrCollect_Fee.SelLength = Len(nbrCollect_Fee.Text)
'End Sub
'
'Private Sub nbrCollect_Fee_GotFocus()
'    nbrCollect_Fee.BackColor = &HFFFFC0
'
'    nbrCollect_Fee.SelStart = 0
'    nbrCollect_Fee.SelLength = Len(nbrCollect_Fee)
'
'    'nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrDisc) + Val(nbrTotCollect_Fee.Text)
'    nbrDisc.Text = (Val(nbrDisc_Per) * Val(nbrTotal)) / 100 'for total discount
'    'nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrDisc) + Val(nbrTotCollect_Fee.Text)
'End Sub
'
'Private Sub nbrCollect_Fee_LostFocus()
'
'    nbrCollect_Fee.BackColor = vbWhite
'
'    If Trim(nbrCollect_Fee.Text) = "" Or Val(nbrCollect_Fee.Text) = 0 Then Exit Sub
'
'    Dim Strmsg As String
'    Strmsg = MsgBox("The Paitent collection fee PAID : " + nbrTotCollect_Fee + "  Now he is going to pay :  " + nbrCollect_Fee + "", vbQuestion + vbYesNo)
'
'        If Strmsg = vbYes Then
'           nbrCollect_Fee.Locked = True
'           nbrTotCollect_Fee = Val(nbrTotCollect_Fee) + Val(nbrCollect_Fee)
'           'nbrCollect_Fee.Text = "0"
'           nbrAdv_Pay.SetFocus
'        Else
'           nbrCollect_Fee.Text = "0"
'           Exit Sub
'        End If
'
'
'
'End Sub
'
'Private Sub nbrDisc_Change()
'    If Not IsNumeric(nbrDisc.Text) Then
'        MsgBox "Only Numaric value allow"
'        nbrDisc = 0
'        nbrDisc.SelStart = 0
'        nbrDisc.SelLength = Len(nbrDisc)
'        nbrDisc.SetFocus
'    End If
'
'    If Len(nbrTotal) = 0 Then Exit Sub
'
''    nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrDisc) + Val(nbrTotCollect_Fee.Text)
'
'    If Val(nbrTotal) = 0 Then Exit Sub
'    nbrDisc_Per.Text = Val(nbrTot_Disc) * 100 / Val(nbrTotal) ' for percentence
'
'
'    If Val(nbrTotal_Amt) + Val(nbrTotCollect_Fee) = Val(Me.nbrTot_Disc) Then
'            ChkPaid.Value = 1
'        Else
'            ChkPaid.Value = 0
'    End If
''
''    If Val(nbrDisc_Per.Text) = 0 Then
''       nbrDisc_Per.Text = ((Val(nbrDisc) * 100) / Val(nbrTotal.Text))
''    Else
''
''    End If
'
'
'    nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrDisc) + Val(nbrTotCollect_Fee.Text)
'
'End Sub
'Private Sub nbrDisc_GotFocus()
'
'    nbrDisc.BackColor = &HFFFFC0
'
'    nbrDisc.SelStart = 0
'    nbrDisc.SelLength = Len(nbrDisc)
'End Sub
'
'
'Private Sub nbrDisc_LostFocus()
'On Error Resume Next
'nbrDisc.BackColor = vbWhite
'
'If nbrDisc = "" Or nbrDisc = 0 Then Exit Sub
'
'
'    Strmsg1 = MsgBox("The Paitent's present Disscount " + nbrTot_Disc + "  Do you want to pay more  " + nbrDisc + "", vbQuestion + vbYesNo)
'        If Strmsg1 = vbYes Then
'            Dim StrNbrDisc As String
'            StrNbrDisc = nbrDisc.Text
'
'           nbrDisc.Locked = True
'           nbrTot_Disc = Val(nbrTot_Disc) + Val(nbrDisc)
'           nbrDisc_Per.Text = Val(nbrDisc) * 100 / Val(nbrTotal)
'           nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrDisc) + Val(nbrTotCollect_Fee.Text)
'
'            nbrDisc.Text = StrNbrDisc
'        Else
'
'            nbrDisc.Text = "0"
'
'            Exit Sub
'        End If
'End Sub
'
'Private Sub nbrDisc_Per_Change()
'    If Not IsNumeric(nbrDisc_Per.Text) Then
'        MsgBox "Only Numaric value allow"
'        nbrDisc_Per = 0
'        nbrDisc_Per.SelStart = 0
'        nbrDisc_Per.SelLength = Len(nbrDisc_Per)
'        nbrDisc_Per.SetFocus
'    End If
'
'    If Trim(nbrTotal) = 0 Then Exit Sub
'    If Trim(nbrDisc) = 0 Then Exit Sub
'
'    nbrDisc_Per.Text = Val(nbrDisc) * 100 / Val(nbrTotal) ' for percentence
'End Sub
'
'Private Sub nbrDisc_Per_GotFocus()
'    nbrDisc_Per.BackColor = &HFFFFC0
'
'
'    nbrDisc_Per.SelStart = 0
'    nbrDisc_Per.SelLength = Len(nbrDisc_Per)
'End Sub
'
'Private Sub nbrDisc_Per_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'
'    End If
'End Sub
'
'Private Sub nbrDisc_Per_LostFocus()
'
'If Me.nbrDisc = "0" Then
'    nbrDisc.Text = (Val(nbrDisc_Per) * Val(nbrTotal)) / 100 'for total discount
'
'    nbrTot_Disc = Val(nbrDisc.Text) + Val(nbrTot_Disc)
'End If
'
'
'
'    nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrTot_Disc) - Val(nbrDisc) + Val(nbrTotCollect_Fee.Text)
'
'    nbrDisc_Per.BackColor = vbWhite
'
'End Sub
'
'Private Sub nbrNet_Amount_Change()
'
'
'    nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrTot_Disc) + Val(nbrTotCollect_Fee.Text)
'
'
'    nbrDue = Val(nbrNet_Amount) - Val(nbrAdv)
'    If Val(nbrNet_Amount) = 0 Then Exit Sub
'    If Val(nbrAdv) = 0 Then Exit Sub
'    If Val(nbrNet_Amount) = Val(nbrAdv) Then
'    ChkPaid.Value = 1
'    Else
'        If Val(nbrTotal_Amt) + Val(nbrTotCollect_Fee) = Val(nbrTot_Disc) Then
'            ChkPaid.Value = 1
'        Else
'            ChkPaid.Value = 0
'        End If
'    End If
'
'End Sub
'
'Private Sub nbrTest_Rate_Change()
'    If Not IsNumeric(nbrTest_Rate.Text) Then
'        MsgBox "Only Numaric value allow"
'        nbrTest_Rate = 0
'        nbrTest_Rate.SelStart = 0
'        nbrTest_Rate.SelLength = Len(nbrTest_Rate)
'        nbrTest_Rate.SetFocus
'    End If
'End Sub
'
'Private Sub nbrTest_Rate_GotFocus()
'    nbrTest_Rate.BackColor = &HFFFFC0
'
'    nbrTest_Rate.SelStart = 0
'    nbrTest_Rate.SelLength = Len(nbrTest_Rate)
'End Sub
'
'Private Sub nbrTest_Rate_LostFocus()
'    nbrTest_Rate.BackColor = vbWhite
'End Sub
'
'Private Sub nbrTot_Disc_Change()
'nbrDisc.Text = (Val(nbrDisc_Per) * Val(nbrTotal)) / 100 'for total discount
'End Sub
'
'Private Sub nbrTot_Disc_GotFocus()
'nbrTot_Disc.BackColor = &HFFFFC0
'End Sub
'
'Private Sub nbrTot_Disc_LostFocus()
'    Me.nbrTot_Disc.BackColor = vbWhite
'End Sub
'
'Private Sub nbrTotal_Amt_Change()
'
'
'    nbrTotal_Amt = Val(nbrTotal) + Val(nbrVAT_Amt)
'    nbrNet_Amount = Val(nbrTotal_Amt) - (Val(nbrTot_Disc) + Val(nbrDisc)) + Val(nbrTotCollect_Fee.Text)
'
'    If Val(nbrTotal_Amt) + Val(nbrTotCollect_Fee) = Val(nbrTot_Disc) Then
'            ChkPaid.Value = 1
'        Else
'            ChkPaid.Value = 0
'    End If
'
'
'End Sub
'Private Sub nbrTotal_Change()
'
'    nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrTot_Disc) + Val(nbrTotCollect_Fee.Text)
'    nbrVAT_Amt = Val(nbrTotal) * Val(nbrVAT_Per) / 100
'    nbrTotal_Amt = Val(nbrTotal) + Val(nbrVAT_Amt)
'End Sub
'
'Private Sub nbrTotCollect_Fee_Change()
'
'    nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrTot_Disc) + Val(nbrTotCollect_Fee.Text)
'
'    If Val(nbrTotal_Amt) + Val(nbrTotCollect_Fee) = Val(nbrTot_Disc) Then
'            ChkPaid.Value = 1
'        Else
'            ChkPaid.Value = 0
'    End If
'
'End Sub
'
'Private Sub nbrVAT_Amt_Change()
'
'    nbrTotal_Amt = Val(nbrTotal) + Val(nbrVAT_Amt)
'    nbrVAT_Amt = Round(nbrVAT_Amt, 0)
'
'
'End Sub
'
'Private Sub txtAddr_GotFocus()
'txtAddr.BackColor = &HFFFFC0
'End Sub
'
'Private Sub txtAddr_LostFocus()
'
'txtAddr.BackColor = vbWhite
'
'End Sub
'
'Private Sub txtAge_GotFocus()
'txtAge.BackColor = &HFFFFC0
'End Sub
'
'Private Sub txtAge_LostFocus()
'    txtAge.BackColor = vbWhite
'End Sub
'
'Private Sub txtDoc_Addr_GotFocus()
'    txtDoc_Addr.BackColor = &HFFFFC0
'End Sub
'
'Private Sub txtDoc_Addr_LostFocus()
'txtDoc_Addr.BackColor = vbWhite
'End Sub
'
'Private Sub txtDoc_Name_GotFocus()
'    txtDoc_Name.BackColor = &HFFFFC0
'End Sub
'
'Private Sub txtDoc_Name_LostFocus()
'txtDoc_Name.BackColor = vbWhite
'End Sub
'
'Private Sub txtEmail_GotFocus()
'txtEmail.BackColor = &HFFFFC0
'End Sub
'
'Private Sub txtEmail_LostFocus()
'txtEmail.BackColor = vbWhite
'End Sub
'
'Private Sub txtFax_GotFocus()
'
'txtFax.BackColor = &HFFFFC0
'
'End Sub
'
'Private Sub txtFax_LostFocus()
'txtFax.BackColor = vbWhite
'End Sub
'
'Private Sub txtM_Code_GotFocus()
'txtM_Code.BackColor = &HFFFFC0
'End Sub
'
'Private Sub txtM_Code_LostFocus()
'
'    On Error GoTo err_sub
'
'    txtM_Code.BackColor = vbWhite
'
'    Test_List_Mode = "frmPatient_Info_M" 'mode for show 'TEST NAME LIST'
'
'
'    If Trim(txtM_Code.Text) = "" Then
'        If Val(nbrTotal) <> 0 Then
'            nbrDisc.SetFocus
'        End If
'        Exit Sub
'    End If
'
'    Adodc4.ConnectionString = strcn.Connection
'    Adodc4.RecordSource = "exec  sp_found '" + txtM_Code + "',''"
'    Adodc4.Refresh
'
'    If Adodc4.Recordset.Fields(0) = "N" Then
'     frmTest_List.Show vbModal 'show TEST NAME order by s_code
'     Exit Sub
'    End If
'    Exit Sub
'
'err_sub:
'    MsgBox Err.Description
'End Sub
'Private Sub txtPat_ID_Change()
'    If Trim(txtPat_ID) = "" Then Exit Sub
'    If Not IsNumeric(txtPat_ID.Text) Then
'        MsgBox "Invalid Patient ID, Please try again.......  "
'        txtPat_ID = ""
'        txtPat_ID.SelStart = 0
'        txtPat_ID.SelLength = Len(txtPat_ID)
'        txtPat_ID.SetFocus
'    End If
'
'End Sub
'Private Sub txtPat_ID_GotFocus()
'    txtPat_ID.SelStart = 0
'    txtPat_ID.SelLength = Len(txtPat_ID)
'End Sub
'
'Private Sub txtPat_ID_LostFocus()
''
'
'    ChkPaid.Value = 0
'    Temp_rst
'    StrAdv_sum = 0
'    nbrAdv.Text = ""
'   '-----------------------------------------------------------
'    DataGrid1.Columns(0).Width = 450.1418
'    DataGrid1.Columns(1).Width = 810.1418
'    DataGrid1.Columns(2).Width = 3825.071
'    DataGrid1.Columns(3).Width = 1110.047
'    DataGrid1.Columns(4).Width = 1900
'    DataGrid1.Columns(5).Width = 600
'    '-----------------------------------------------------------
'
'    If Len(Trim(txtPat_ID.Text)) = 0 Then Exit Sub
'
'       Adodc3.ConnectionString = strcn.Connection
'       Adodc3.RecordSource = "exec Pat_Info_SELECT 1," + txtPat_ID + ""
'       Adodc3.Refresh
'       If Adodc3.Recordset.RecordCount > 0 Then
'          txtPat_ID.Text = Adodc3.Recordset!pat_id
'          txtPat_Name = Adodc3.Recordset!pat_name
'          ComSex = Adodc3.Recordset!Sex
'          txtAge = Adodc3.Recordset!age
'
'          txtAddr = Adodc3.Recordset!addr
'          txtPhone = Adodc3.Recordset!phone
'          txtFax = Adodc3.Recordset!fax
'          txtEmail = Adodc3.Recordset!email
'
'          nbrVAT_Per = Adodc3.Recordset!vat_per
'          nbrVAT_Amt = Adodc3.Recordset!vat_amt
'
'            '`````````````to show date and time from pat_info_main``````
'           Adodc11.ConnectionString = strcn.Connection
'           Adodc11.RecordSource = "exec Pat_Info_SELECT 1,'" + txtPat_ID + "'"
'           Adodc11.Refresh
'
'            Dim StrCdt1 As String
'            Dim StrCtime1 As String
'            Dim CDate_TM1 As String
'
'           If Adodc11.Recordset.RecordCount > 0 Then
'            CDate_TM1 = Adodc11.Recordset!Dt
'            CDate_TM3 = Adodc11.Recordset!tmp_Dt
'            CDate_TM6 = Adodc11.Recordset!dt1
'
'            StrCdt1 = Mid(CDate_TM1, 1, 10)
'            StrCtime1 = Mid(CDate_TM1, 12, 12)
'            Dt = StrCdt1
'            DT_TM = StrCtime1
''
'            End If
'
'     '```````END````````````````````````````````````````````````
'
'     '`````````````to show date and time from pat_info_sub1``````
'           Adodc11.ConnectionString = strcn.Connection
'           Adodc11.RecordSource = "exec Pat_Info_SELECT 5,'" + txtPat_ID + "'"
'           Adodc11.Refresh
'
'           If Adodc11.Recordset.RecordCount > 0 Then
'            CDate_TM2 = Adodc11.Recordset!tmp_Dt
'            CDate_TM7 = Adodc11.Recordset!dt1
'           End If
'      '`````````````````END```````````````````````````
'
'      '`````````````to show date and time from pat_info_sub2``````
'           Adodc11.ConnectionString = strcn.Connection
'           Adodc11.RecordSource = "exec Pat_Info_SELECT 2,'" + txtPat_ID + "'"
'           Adodc11.Refresh
'
'           If Adodc11.Recordset.RecordCount > 0 Then
'            CDate_TM5 = Adodc11.Recordset!tmp_Dt
'            CDate_TM8 = Adodc11.Recordset!dt1
'           End If
'      '`````````````````END```````````````````````````
'
'      '`````````````to show date and time from pat_info_sub3``````
'           Adodc11.ConnectionString = strcn.Connection
'           Adodc11.RecordSource = "exec Pat_Info_SELECT 3,'" & txtPat_ID & "'"
'           Adodc11.Refresh
'          If Adodc11.Recordset.RecordCount > 0 Then
'            CDate_TM4 = Adodc11.Recordset!tmp_Dt
'            CDate_TM9 = Adodc11.Recordset!dt1
'           End If
'      '`````````````````END```````````````````````````
'
'
'
'           '--------flush into Temp_Tabel-------------------------------
'            Con.ConnectionString = strcn.Connection
'            Con.Open
'
'            Temp_Table_Helper.Open "select m_code,s_code,(select s_name from test_info_sub Where test_info_sub.s_code = pat_info_sub1.s_code and test_info_sub.m_code=pat_info_sub1.m_code and pat_id='" + txtPat_ID + "') as s_name,test_rate,delv_dt,type,unique_id from pat_info_sub1 where pat_id='" + txtPat_ID + "'", Con
'
'            'MsgBox Temp_Table_Helper.RecordCount
'              While Temp_Table_Helper.EOF = False
'                    Temp_Table.AddNew
'
'                    Temp_Table!m_code = Temp_Table_Helper!m_code
'                    Temp_Table!s_code = Temp_Table_Helper!s_code
'                    Temp_Table!s_name = Temp_Table_Helper!s_name
'                    Temp_Table!test_rate = Temp_Table_Helper!test_rate
'                    Temp_Table!Delv_DTM = Temp_Table_Helper!Delv_Dt
'                    Temp_Table!Type = Temp_Table_Helper!Type
'                    Temp_Table_Helper.MoveNext
'              Wend
'
'            DataGrid1.Refresh
'            Temp_Table_Helper.Close
'            Con.Close
'
'
'           '---------------------------------------------------------
'               '`````````````to show DISCOUNT from pat_info_sub3``````
'               Adodc6.ConnectionString = strcn.Connection
'               Adodc6.RecordSource = "exec Pat_Info_SELECT 11,'" & txtPat_ID.Text & "'"
'               Adodc6.Refresh
'
'               If Adodc6.Recordset.RecordCount > 0 Then
'               Dim strchkpaid As String
'                nbrDisc.Text = "0"
'
'                nbrTot_Disc = Adodc6.Recordset!disc
'                strchkpaid = Adodc6.Recordset!paid
'                'MsgBox strchkpaid
'                    If Trim(strchkpaid) = "True" Then
'                    ChkPaid.Value = 1
'                    ChkPaidVal = "1"
'                    Else
'                    ChkPaid.Value = 0
'                    ChkPaidVal = "0"
'                    End If
'               End If
'               '```````````````````````````````````````````````````````
'
'               '`````````````to show REFERENCE_TYPE from pat_info_MAIN``````
'               Adodc6.ConnectionString = strcn.Connection
'               Adodc6.RecordSource = "exec Pat_Info_SELECT 1,'" + txtPat_ID + "'"
'               Adodc6.Refresh
'
'               If Adodc6.Recordset.RecordCount > 0 Then
'               Dim strRefer_Type1 As String
'
'                strRefer_Type1 = Adodc6.Recordset!refer_type
'                    If strRefer_Type1 = 1 Then
'                    Chkrefer_type.Value = 1
'                    strRefer_Type1 = "1"
'                    Else
'                    Chkrefer_type.Value = 0
'                    strRefer_Type1 = "0"
'                    End If
'               End If
'               '``````````````````````````````````````````````````````
'
'               '*************for flush doctor ID and name ****************
'               Adodc12.ConnectionString = strcn.Connection
'               Adodc12.RecordSource = "exec Pat_Info_SELECT 7,'" + txtPat_ID + "'"
'
'               Adodc12.Refresh
'               If Adodc12.Recordset.RecordCount > 0 Then
'
'                   txtRefer_Code = Adodc12.Recordset!refer_code
'                'MsgBox txtRefer_Code
'
'               End If
'
'
''              '======DONTOR NAME FROM DOCTOR_INFO_NEW=============
'               Adodc13.ConnectionString = strcn.Connection
'               Adodc13.RecordSource = "exec Pat_Info_SELECT 6,'" + txtPat_ID + "'"
'
'               Adodc13.Refresh
'               If Adodc13.Recordset.RecordCount > 0 Then
'
'                  txtDoc_Name = Adodc13.Recordset!doc_name
'                  txtDoc_Addr = Adodc13.Recordset!addr
'               End If
'               '=====================END===========================
'               ',,,,,,,,,,,,,,for get registered doctor,,,,,,,,,,,
'               Dim My_Rst As New ADODB.Recordset
'               Con.ConnectionString = strcn.Connection
'               Con.Open
'               Set My_Rst.ActiveConnection = Con
'               My_Rst.Open "exec Pro_FLUSH1 1,'" & Trim(txtRefer_Code.Text) & "'", Con
'               If My_Rst.EOF = False Then
'
'                    txtDoc_Name.Text = My_Rst!doc_name
'                    txtDoc_Addr.Text = My_Rst!addr
'               Else
'                    txtDoc_Name.ForeColor = vbBlack
'                    txtDoc_Addr.ForeColor = vbBlack
'               End If
'               My_Rst.Close
'               Con.Close
'
'
'               ',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
'
'               '***************end****************************************
'
'         Else
'           txtPat_Name = ""
'           ComSex = "Male"
'           txtAge = ""
'           txtRefer_Code = ""
'           txtDegree = ""
'           txtAddr = ""
'           txtPhone = ""
'           txtFax = ""
'           txtEmail = ""
'           Dt.Value = Now
'           Delv_Dt.Value = Now
'           nbrVAT_Amt = 0
'           nbrAdv = 0
'           nbrDisc = 0
'           nbrTot_Disc = 0
'           nbrDisc_Per = 0
'           nbrDue = ""
'           nbrNet_Amount = ""
'
'           nbrTest_Rate = ""
'           nbrTotal = ""
'           ChkPaid.Value = 0
'           Delv_TM.Value = Now
'           Chkrefer_type.Value = 0
'        End If
'
''++++++++++for count TOTAL_RATE from Temp_Table+++++++++
'        If Temp_Table.RecordCount > 0 Then
'        Total_Rate = 0
'        Temp_Table.MoveFirst
'        While Temp_Table.EOF = False
'                Total_Rate = Total_Rate + Temp_Table!test_rate
'
'        Temp_Table.MoveNext
'        Wend
'        nbrTotal = Total_Rate
'        End If
''++++++++++End count TOTAL_RATE from Temp_Table+++++++++
'
''=========count total test=============================
'        Total_Test = 0
'        Total_Test = Temp_Table.RecordCount
'        Me.nbrTot_Test = Total_Test
''=========End count total test========================
'
''>>>>>>>>>>>>>>>>to show total advance>>>>>>>>>>>>>>>>>>>>>>
'    Adodc7.ConnectionString = strcn.Connection
'    Adodc7.RecordSource = "exec Pro_FLUSH 3,'" & Trim(txtPat_ID.Text) & "'"
'    Adodc7.Refresh
'    If Adodc7.Recordset.RecordCount > 0 Then
'        nbrAdv.Text = Adodc7.Recordset!adv_sum
'        nbrTotCollect_Fee.Text = Adodc7.Recordset!Coll_sum
'    End If
''<<<<<<<<<<<<End show total advance<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'
''nbrDisc_Per.Text = Val(nbrDisc) * 100 / Val(nbrTotal) ' for percentence
'
'    DataGrid1.Columns(0).Width = 450.1418
'    DataGrid1.Columns(1).Width = 810.1418
'    DataGrid1.Columns(2).Width = 3825.071
'    DataGrid1.Columns(3).Width = 1110.047
'    DataGrid1.Columns(4).Width = 1900
'    DataGrid1.Columns(5).Width = 600
'
'nbrAdv_Pay.SetFocus
'End Sub
'
'Private Sub txtPat_ID1_LostFocus()
'On Error Resume Next
'
'    If txtPat_ID1 = "" Then Exit Sub
'    If txtPat_ID1 <> "" Then
'        txtPat_ID.TabStop = False
'    End If
'    Search_Patient_Type
'
'    If StrRow_Count > "1" Then
'
'            Dim Patmsg As String
'            Patmsg = MsgBox("Do you want to update Inside Patient's information ? ", vbQuestion + vbYesNo)
'            If Patmsg = vbYes Then
'                StrPat_Type = "0"
'
'                Srch_Pat_ID
'            Else
'                StrPat_Type = "1"
'
'                Srch_Pat_ID
'
'            End If
'    Else
'            Srch_Pat_ID1
'    End If
'
'
'    txtPat_ID = IntPat_ID
'
'    txtDummy_Pat_ID.Text = IntPat_ID
'
'    If IntPat_ID = 0 Then
'        MsgBox "Invalid ID, Try again"
'        txtPat_ID = ""
'        txtPat_ID1 = ""
'        txtDummy_Pat_ID = ""
'        txtPat_ID1.SetFocus
'        Exit Sub
'    End If
'
'
'    '---------------------------------------------------------------
'
'    ChkPaid.Value = 0
'    Temp_rst
'    StrAdv_sum = 0
'    nbrAdv.Text = ""
'   '-----------------------------------------------------------
'    DataGrid1.Columns(0).Width = 450.1418
'    DataGrid1.Columns(1).Width = 810.1418
'    DataGrid1.Columns(2).Width = 3825.071
'    DataGrid1.Columns(3).Width = 1110.047
'    DataGrid1.Columns(4).Width = 1900
'    DataGrid1.Columns(5).Width = 600
'    '-----------------------------------------------------------
'
'    If Len(Trim(txtPat_ID.Text)) = 0 Then Exit Sub
'      'for flush patient information
'       Adodc3.ConnectionString = strcn.Connection
'       Adodc3.RecordSource = "exec Pat_Info_SELECT 1," & txtDummy_Pat_ID.Text & ""
'       Adodc3.Refresh
'       If Adodc3.Recordset.RecordCount > 0 Then
'          txtPat_ID.Text = Adodc3.Recordset!pat_id
'          txtPat_Name = Adodc3.Recordset!pat_name
'          ComSex = Adodc3.Recordset!Sex
'          txtAge = Adodc3.Recordset!age
'          txtAddr = Adodc3.Recordset!addr
'          txtPhone = Adodc3.Recordset!phone
'          txtFax = Adodc3.Recordset!fax
'          txtEmail = Adodc3.Recordset!email
'          nbrVAT_Per = Adodc3.Recordset!vat_per
'          nbrVAT_Amt = Adodc3.Recordset!vat_amt
'          StPat_Type1 = Adodc3.Recordset!refer_type
'          DummyPat_ID1 = Adodc3.Recordset!pat_id1
'          Strpat_MY = Adodc3.Recordset!pat_my
'
''          MsgBox DummyPat_ID1
''          MsgBox Strpat_MY
'
'            '`````````````to show date and time from pat_info_main``````
'           Adodc11.ConnectionString = strcn.Connection
'           Adodc11.RecordSource = "exec Pat_Info_SELECT 1,'" + txtDummy_Pat_ID.Text + "'"
'           Adodc11.Refresh
'
'            Dim StrCdt1 As String
'            Dim StrCtime1 As String
'            Dim CDate_TM1 As String
'
'           If Adodc11.Recordset.RecordCount > 0 Then
'            CDate_TM1 = Adodc11.Recordset!Dt
'            CDate_TM3 = Adodc11.Recordset!tmp_Dt
'            CDate_TM6 = Adodc11.Recordset!dt1
'
'            StrCdt1 = Mid(CDate_TM1, 1, 10)
'            StrCtime1 = Mid(CDate_TM1, 12, 12)
'            Dt = StrCdt1
'            DT_TM = StrCtime1
''
'            End If
'
'           '```````END````````````````````````````````````````````````
'
'     '`````````````to show date and time from pat_info_sub1``````
'           Adodc11.ConnectionString = strcn.Connection
'           Adodc11.RecordSource = "exec Pat_Info_SELECT 5,'" + txtDummy_Pat_ID.Text + "'"
'           Adodc11.Refresh
'
'           If Adodc11.Recordset.RecordCount > 0 Then
'            CDate_TM2 = Adodc11.Recordset!tmp_Dt
'            CDate_TM7 = Adodc11.Recordset!dt1
'           End If
'      '`````````````````END```````````````````````````
'
'      '`````````````to show date and time from pat_info_sub2``````
'           Adodc11.ConnectionString = strcn.Connection
'           Adodc11.RecordSource = "exec Pat_Info_SELECT 2,'" + txtDummy_Pat_ID.Text + "'"
'           Adodc11.Refresh
'
'           If Adodc11.Recordset.RecordCount > 0 Then
'            CDate_TM5 = Adodc11.Recordset!tmp_Dt
'            CDate_TM8 = Adodc11.Recordset!dt1
'           End If
'      '`````````````````END```````````````````````````
'
'      '`````````````to show date and time from pat_info_sub3``````
'           Adodc11.ConnectionString = strcn.Connection
'           Adodc11.RecordSource = "exec Pat_Info_SELECT 3,'" & txtDummy_Pat_ID.Text & "'"
'           Adodc11.Refresh
'          If Adodc11.Recordset.RecordCount > 0 Then
'            CDate_TM4 = Adodc11.Recordset!tmp_Dt
'            CDate_TM9 = Adodc11.Recordset!dt1
'           End If
'      '`````````````````END```````````````````````````
'
'
'
'           '--------flush into Temp_Tabel-------------------------------
'            Con.ConnectionString = strcn.Connection
'            Con.Open
'
'            Temp_Table_Helper.Open "select m_code,s_code,(select s_name=isnull(s_name,'') from test_info_sub Where test_info_sub.s_code = pat_info_sub1.s_code and test_info_sub.m_code=pat_info_sub1.m_code and pat_id='" + txtPat_ID + "') as s_name,test_rate,delv_dt,type,unique_id from pat_info_sub1 where pat_id='" + txtDummy_Pat_ID.Text + "'", Con
'
'            'MsgBox Temp_Table_Helper.RecordCount
'              While Temp_Table_Helper.EOF = False
'                    Temp_Table.AddNew
'
'                    Temp_Table!m_code = Temp_Table_Helper!m_code
'                    Temp_Table!s_code = Temp_Table_Helper!s_code
'                    Temp_Table!s_name = Temp_Table_Helper!s_name
'                    Temp_Table!test_rate = Temp_Table_Helper!test_rate
'                    Temp_Table!Delv_DTM = Temp_Table_Helper!Delv_Dt
'                    Temp_Table!Type = Temp_Table_Helper!Type
'                    Temp_Table_Helper.MoveNext
'              Wend
'
'            DataGrid1.Refresh
'            Temp_Table_Helper.Close
'            Con.Close
'
'
'           '---------------------------------------------------------
'
'               '`````````````to show DISCOUNT from pat_info_sub3``````
'               Adodc6.ConnectionString = strcn.Connection
'               Adodc6.RecordSource = "exec Pat_Info_SELECT 11,'" & txtDummy_Pat_ID.Text & "'"
'               Adodc6.Refresh
'
'               If Adodc6.Recordset.RecordCount > 0 Then
'               Dim strchkpaid As String
'                nbrDisc.Text = "0"
'
'                nbrTot_Disc = Adodc6.Recordset!disc
'                strchkpaid = Adodc6.Recordset!paid
'                'MsgBox strchkpaid
'                    If Trim(strchkpaid) = "True" Then
'                    ChkPaid.Value = 1
'                    ChkPaidVal = "1"
'                    Else
'                    ChkPaid.Value = 0
'                    ChkPaidVal = "0"
'                    End If
'               End If
'               '```````````````````````````````````````````````````````
'
'               '`````````````to show REFERENCE_TYPE from pat_info_MAIN``````
'               Adodc6.ConnectionString = strcn.Connection
'               Adodc6.RecordSource = "exec Pat_Info_SELECT 1,'" + txtDummy_Pat_ID.Text + "'"
'               Adodc6.Refresh
'
'               If Adodc6.Recordset.RecordCount > 0 Then
'               Dim strRefer_Type1 As String
'
'                strRefer_Type1 = Adodc6.Recordset!refer_type
'                    If strRefer_Type1 = 1 Then
'                    Chkrefer_type.Value = 1
'                    strRefer_Type1 = "1"
'                    Else
'                    Chkrefer_type.Value = 0
'                    strRefer_Type1 = "0"
'                    End If
'               End If
'               '``````````````````````````````````````````````````````
'
'               '*************for flush doctor ID and name ****************
'               Adodc12.ConnectionString = strcn.Connection
'               Adodc12.RecordSource = "exec Pat_Info_SELECT 7,'" + txtDummy_Pat_ID.Text + "'"
'
'               Adodc12.Refresh
'               If Adodc12.Recordset.RecordCount > 0 Then
'
'                   txtRefer_Code = Adodc12.Recordset!refer_code
'                'MsgBox txtRefer_Code
'
'               End If
'
'
''              '======DONTOR NAME FROM DOCTOR_INFO_NEW=============
'               Adodc13.ConnectionString = strcn.Connection
'               Adodc13.RecordSource = "exec Pat_Info_SELECT 6,'" + txtDummy_Pat_ID.Text + "'"
'
'               Adodc13.Refresh
'               If Adodc13.Recordset.RecordCount > 0 Then
'
'                  txtDoc_Name = Adodc13.Recordset!doc_name
'                  txtDoc_Addr = Adodc13.Recordset!addr
'               End If
'               '=====================END===========================
'               ',,,,,,,,,,,,,,for get registered doctor,,,,,,,,,,,
'               Dim My_Rst As New ADODB.Recordset
'               Con.ConnectionString = strcn.Connection
'               Con.Open
'               Set My_Rst.ActiveConnection = Con
'               My_Rst.Open "exec Pro_FLUSH1 1,'" & Trim(txtRefer_Code.Text) & "'", Con
'               If My_Rst.EOF = False Then
'
'                    txtDoc_Name.Text = My_Rst!doc_name
'                    txtDoc_Addr.Text = My_Rst!addr
'               Else
'                    txtDoc_Name.ForeColor = vbBlack
'                    txtDoc_Addr.ForeColor = vbBlack
'               End If
'               My_Rst.Close
'               Con.Close
'
'
'               ',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
'
'               '***************end****************************************
'
'         Else
'           txtPat_Name = ""
'           ComSex = "Male"
'           txtAge = ""
'           txtRefer_Code = ""
'           txtDegree = ""
'           txtAddr = ""
'           txtPhone = ""
'           txtFax = ""
'           txtEmail = ""
'           Dt.Value = Now
'           Delv_Dt.Value = Now
'           nbrVAT_Amt = 0
'           nbrAdv = 0
'           nbrDisc = 0
'           nbrTot_Disc = 0
'           nbrDisc_Per = 0
'           nbrDue = ""
'           nbrNet_Amount = ""
'
'           nbrTest_Rate = ""
'           nbrTotal = ""
'           ChkPaid.Value = 0
'           Delv_TM.Value = Now
'           Chkrefer_type.Value = 0
'        End If
'
''++++++++++for count TOTAL_RATE from Temp_Table+++++++++
'        If Temp_Table.RecordCount > 0 Then
'        Total_Rate = 0
'        Temp_Table.MoveFirst
'        While Temp_Table.EOF = False
'                Total_Rate = Total_Rate + Temp_Table!test_rate
'
'        Temp_Table.MoveNext
'        Wend
'        nbrTotal = Total_Rate
'        End If
''++++++++++End count TOTAL_RATE from Temp_Table+++++++++
'
''=========count total test=============================
'        Total_Test = 0
'        Total_Test = Temp_Table.RecordCount
'        Me.nbrTot_Test = Total_Test
''=========End count total test========================
'
''>>>>>>>>>>>>>>>>to show total advance>>>>>>>>>>>>>>>>>>>>>>
'    Adodc7.ConnectionString = strcn.Connection
'    Adodc7.RecordSource = "exec Pro_FLUSH 3,'" & txtDummy_Pat_ID.Text & "'"
'    Adodc7.Refresh
'    If Adodc7.Recordset.RecordCount > 0 Then
'        nbrAdv.Text = Adodc7.Recordset!adv_sum
'        nbrTotCollect_Fee.Text = Adodc7.Recordset!Coll_sum
'    End If
''<<<<<<<<<<<<End show total advance<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'
'    DataGrid1.Columns(0).Width = 450.1418
'    DataGrid1.Columns(1).Width = 810.1418
'    DataGrid1.Columns(2).Width = 3825.071
'    DataGrid1.Columns(3).Width = 1110.047
'    DataGrid1.Columns(4).Width = 1900
'    DataGrid1.Columns(5).Width = 600
'
'nbrAdv_Pay.SetFocus
'
'
'End Sub
'
'Private Sub txtPat_Name_GotFocus()
'txtPat_Name.BackColor = &HFFFFC0
'End Sub
'
'Private Sub txtPat_Name_LostFocus()
'    txtPat_Name.BackColor = vbWhite
'End Sub
'
'Private Sub txtPhone_GotFocus()
'
'txtPhone.BackColor = &HFFFFC0
'
'End Sub
'
'Private Sub txtPhone_LostFocus()
'txtPhone.BackColor = vbWhite
'End Sub
'
'Private Sub txtRefer_Code_GotFocus()
'txtRefer_Code.BackColor = &HFFFFC0
'End Sub
'
'Private Sub txtRefer_Code_LostFocus()
'On Error GoTo err_sub
'    txtRefer_Code.BackColor = vbWhite
'
'    txtM_Code.TabStop = True
'    If Trim(txtRefer_Code) = "" Then Exit Sub
'    'MsgBox "Patient1"
'    Doc_List_MODE = "frmPatient_Info"
'
'       If Trim(txtRefer_Code.Text) = "0" Then
'
'            If Trim(txtPat_ID.Text) <> "" Then
'
'                If u_id <> "md" Then
'                MsgBox "If you want to any change you should contact to Managing Director.., Your ID saved..", vbCritical
'                txtRefer_Code = ""
'                Exit Sub
'                End If
'                NdocMode = "0"
'                frmDoctor_Info_New.txtPat_ID = txtPat_ID
'            End If
'
'            If Trim(txtPat_ID.Text) = "" Then
'                NdocMode = "1"
'                frmDoctor_Info_New.txtPat_ID = "0"
'            End If
'
'       frmDoctor_Info_New.Show vbModal 'for new unknown doctor
'
'       Else
'               Adodc2.ConnectionString = strcn.Connection
'
'               Adodc2.RecordSource = "exec Pro_FLUSH1 1,'" & Trim(txtRefer_Code.Text) & "'"
'               Adodc2.Refresh
'
'
'                'MsgBox "Patient2"
'               If Adodc2.Recordset.RecordCount > 0 Then
'                   txtDoc_Name.Text = Adodc2.Recordset!doc_name
'                   txtDoc_Addr.Text = Adodc2.Recordset!addr
'                   txtM_Code.TabStop = True
'               Else
'
'                   'MsgBox "Patient3"
'                   txtM_Code.TabStop = False
'                   frmDoc_List.Show vbModal
'                   Exit Sub
'               End If
'       End If
'    Exit Sub
'
'err_sub:
'    MsgBox Err.Description
'
'End Sub
'
'Private Sub txtS_Code_GotFocus()
'txtS_Code.BackColor = &HFFFFC0
'End Sub
'
'Private Sub txtS_Code_LostFocus()
'
'On Error Resume Next
'
'    txtS_Code.BackColor = vbWhite
'
'    If Trim(txtS_Code) = "" Then Exit Sub
'
'    Adodc4.ConnectionString = strcn.Connection
'    Adodc4.RecordSource = "exec  sp_found '" + txtM_Code + "','" + txtS_Code + "'"
'    Adodc4.Refresh
'
'    If Adodc4.Recordset.Fields(0) = "N" Then
'        Test_List_Mode = "frmPatient_Info_S"
'        txtS_Name = ""
'        nbrTest_Rate = 0
'        txtType.Text = ""
'        txtS_Code = ""
'        frmTest_List.Show vbModal
'    Else
'        If Len(Trim(txtM_Code.Text)) = 0 Then
'            MsgBox "Group Code mandatory"
'            txtM_Code.SetFocus
'            Exit Sub
'        End If
'        txtS_Name = Adodc4.Recordset.Fields(0)
'        nbrTest_Rate = Adodc4.Recordset.Fields(1)
'        txtType.Text = Adodc4.Recordset.Fields(2)
'    End If
'
'End Sub
'Public Sub Temp_rst()
'    '--------------------------------------------
'    Set Temp_Table = New ADODB.Recordset
'    With Temp_Table
'        .Fields.Append "m_code", adVarChar, 2
'        .Fields.Append "s_code", adVarChar, 3
'        .Fields.Append "s_name", adVarChar, 60
'        .Fields.Append "test_rate", adDouble
'        .Fields.Append "Delv_DTM", adVarChar, 26
'        .Fields.Append "type", adVarChar, 10
'        .LockType = adLockOptimistic
'        .Open
'    End With
'
'    Set DataGrid1.DataSource = Temp_Table
'
'    DataGrid1.ReBind
'    DataGrid1.Refresh
'
'    DataGrid1.Columns(0).Width = 450.1418
'    DataGrid1.Columns(1).Width = 810.1418
'    DataGrid1.Columns(2).Width = 3825.071
'    DataGrid1.Columns(3).Width = 1110.047
'    DataGrid1.Columns(4).Width = 1900
'    DataGrid1.Columns(5).Width = 600
'
'
'End Sub
'Private Sub Select_Unique_ID()
'    If Len(Trim(txtPat_ID.Text)) = 0 Then Exit Sub
'    If Len(Trim(txtM_Code.Text)) = 0 Then Exit Sub
'    If Len(Trim(txtS_Code.Text)) = 0 Then Exit Sub
'
'    Adodc8.ConnectionString = strcn.Connection
'    Adodc8.RecordSource = "exec pro_flush_unique_id 1,'" + txtPat_ID + "','" + txtM_Code + "','" + txtS_Code + "'"
'
'    Adodc8.Refresh
'    If Adodc8.Recordset.RecordCount > 0 Then
'    nbrUnique_id = Adodc8.Recordset!unique_id
'    Else
'    nbrUnique_id = ""
'    End If
'End Sub
'Private Sub Auto_No()
'
'
'    Dim My_Rst As New ADODB.Recordset
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'    My_Rst.Open "select count(pat_id)+1 from pat_info_main", Con
'    If IsNull(My_Rst.Fields(0)) = False Then
'       txtPat_ID = BoothN + pad("l", 9, My_Rst.Fields(0), "0")
'    End If
'    My_Rst.Close
'    Con.Close
'
'
'End Sub
'Private Sub nbrVAT_Per_Change()
'    nbrVAT_Amt = Val(nbrTotal) * Val(nbrVAT_Per) / 100
'    nbrTotal_Amt = Val(nbrTotal) + Val(nbrVAT_Amt)
'End Sub
'Private Sub InsPat_Info_Sub2()
'    If Trim(brAdv_Pay) = 0 Or Trim(nbrAdv_Pay) = "" Then Exit Sub
'
'
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'    cmd.CommandText = "exec pro_PAT_INFO_SUB2 'I'," + "0" + _
'    "," + nbrAdv_Pay + ",'" + u_id + "','" + CDate_TM + "','" + "" + "'"
'    cmd.Execute
'    Con.Close
'End Sub
'Private Sub InsPat_Info_Sub2_T()
'    If Trim(nbrAdv_Pay) = "" Then
'        nbrAdv_Pay = "0"
'
'
'    End If
'
'
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'    cmd.CommandText = "exec pro_PAT_INFO_SUB2 'I'," + StPat_ID + _
'    "," + nbrAdv_Pay + ",'" + u_id + _
'    "','" + CDate_TM + _
'    "'," + nbrCollect_Fee.Text + _
'    "," + "ADV" + _
'    ",'" + Format(Dt, "yyyy-mm-dd") + _
'    "','" + CDate_TM + _
'    "','" + Format(CDate_TM, "yyyy-mm-dd") + _
'    "','" + "" + "'"
''    MsgBox cmd.CommandText
'    cmd.Execute
'    Con.Close
'
'End Sub
'Private Sub InsPat_Info_Sub2_T1()
'
'    If Trim(brAdv_Pay) = 0 Or Trim(nbrAdv_Pay) = "" Then Exit Sub
'
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'    cmd.CommandText = "exec pro_PAT_INFO_SUB2 'I'," + txtPat_ID + _
'    "," + nbrAdv_Pay + _
'    ",'" + u_id + _
'    "','" + CDate_TM + _
'    "'," + Trim(nbrCollect_Fee.Text) + _
'    "," + "DUE" + _
'    ",'" + Format(CDate_TM5, "yyyy-mm-dd") + _
'    "','" + Format(CDate_TM8, "yyyy-mm-dd hh:mm") + _
'    "','" + Format(CDate_TM10, "yyyy-mm-dd") + _
'    "','" + "" + "'"

Private Sub cmdADD_Click()
Call clear
End Sub

'    cmd.Execute
'    Con.Close
'
'End Sub
'Private Sub InsDoc_info_new()
'    Dim strRefer_Code As String
'    Dim StrDoc_Name As String
'    Dim strAddr As String
'    Dim strPhone As String
'    Dim strFax As String
'    Dim strEmail As String
'    Dim strUid As String
'    Dim strDoc_Date As String
'
'    Adodc15.ConnectionString = strcn.Connection
'
'    Adodc15.RecordSource = "exec New_Doc_Select 2,'','" & u_id & "'"
'    Adodc15.Refresh
'    If Adodc15.Recordset.RecordCount > 0 Then
'        strRefer_Code = Adodc15.Recordset!pat_id
'
'
'    '-------UPDATE DOCTOR ID into doctor_info_new------------------------
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'    cmd.CommandText = "exec pro_DOCTOR_INFO_NEW2 'U','" & StPat_ID & "','" & u_id & "'"
'
'    cmd.Execute
'    Con.Close
'    '-----------------------------------------------------------
'    '>>>>>>>>>>>>>>>>>>
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'    cmd.CommandText = "exec PAT_INFO_MAIN_U 'U'," & StPat_ID & ""
'
'    cmd.Execute
'    Con.Close
'    '>>>>>>>>>>>>>>>>>>>
'
'    End If
'End Sub
'Private Sub InsD_TM()
'
''DON'T DELETE
''    Dim My_Rst As New ADODB.Recordset
''    con.connectionstring = strcn.Connection
''    con.Open
''    Set cmd.ActiveConnection = con
''
''    My_Rst.Open "exec CR_Date", con
''    If My_Rst.EOF = False Then
''        Dt.value = My_Rst!crDate
''        DT_TM.value = My_Rst!crDate
''    End If
''    con.Close
'
'    '++++++for insert Current Date and Time++++++++++++++
'    Dim StrCdt As String
'    Dim StrCtime As String
'
'    StrCdt = Trim(Format(Dt, "yyyy-mm-dd"))
'    StrCtime = Trim(Format(DT_TM, "hh:mm"))
'    CDate_TM = StrCdt + Space(1) + StrCtime
'    CDate_TM10 = StrCdt
'   '++++++++++end+++++++++++++++++++++++++++++++++++++++
'End Sub
'Private Sub Sel_Refer_Type()
'
'    If Chkrefer_type.Value = 1 Then
'        StrRefer_Type = "1"
'    End If
'
'    If Chkrefer_type.Value = 0 Then
'        StrRefer_Type = "0"
'    End If
'End Sub
'Private Sub Search_Refer_Code() 'search again refer_code for update refer_code/delete from doctor_info_new
'    Dim My_Rst As New ADODB.Recordset
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'
'    My_Rst.Open "exec Doc_SELECT 4,'" + txtPat_ID.Text + "'", Con
'    If My_Rst.EOF = False Then
'        Del_Doc = My_Rst!refer_code
'
'    End If
'    Con.Close
'End Sub
'
'Private Sub Del_New_Doc()
'
'    If Del_Doc <> "" Then ''''delete from doctor_info_new
'       'MsgBox "del"
'        Con.ConnectionString = strcn.Connection
'        Con.Open
'        Set cmd.ActiveConnection = Con
'        cmd.CommandText = "exec delete_all 1," + txtPat_ID + ""
'        cmd.Execute
'        Con.Close
'
'       End If
'End Sub
'Private Sub Flush_VAT_Per()
'
'    Dim My_Rst As New ADODB.Recordset
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'
'    My_Rst.Open "exec pro_name_SELECT '19',''", Con
'    If My_Rst.EOF = False Then
'        nbrVAT_Per.Text = My_Rst!vat_per
'    End If
'
'    Con.Close
'
'
'
'
'End Sub
'Private Sub Make_Pat_ID1()
'
'    Dim My_Rst As New ADODB.Recordset
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'
'    My_Rst.Open "exec Make_Pat_ID1 '" & Chkrefer_type.Value & "'", Con
'    If My_Rst.EOF = False Then
'        Strpat_id1 = My_Rst!pat_id1
'        Strpat_MY = My_Rst!pat_my
''        MsgBox Strpat_id1
'    End If
'
'    Con.Close
'
'End Sub
'Private Sub Make_Pat_ID1_U()
'
'    Dim My_Rst As New ADODB.Recordset
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'
'    My_Rst.Open "exec Make_Pat_ID_U '" & Chkrefer_type.Value & "'", Con
'    If My_Rst.EOF = False Then
'        Strpat_id1 = My_Rst!pat_id1
'        Strpat_MY = My_Rst!pat_my
''        MsgBox Strpat_id1
'    End If
'
'    Con.Close
'
'End Sub
'
'
'
'Private Sub Search_Patient_Type()
'
'    StrRow_Count = "1"
'    Dim My_Rst As New ADODB.Recordset
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'
'    My_Rst.Open "exec Search_Pat_Type 1,'" & txtPat_ID1.Text & "'", Con
'    If My_Rst.EOF = False Then
'
'        StrRow_Count = My_Rst!Row_Count
'        'MsgBox StrRow_Count
'    End If
'
'    Con.Close
'
'End Sub
'Private Sub Srch_Pat_ID()
'
'    Dim My_Rst As New ADODB.Recordset
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'
'    My_Rst.Open "exec Search_Pat_ID 1,'" & txtPat_ID1.Text & "','" & StrPat_Type & "'", Con
'    If My_Rst.EOF = False Then
'        IntPat_ID = My_Rst!pat_id2
'  '      MsgBox IntPat_ID
'    End If
'    Con.Close
'
'End Sub
'Private Sub Srch_Pat_ID1()
'
'    Dim My_Rst As New ADODB.Recordset
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'
'    My_Rst.Open "exec Search_Pat_ID1 1,'" & txtPat_ID1.Text & "'", Con
'    If My_Rst.EOF = False Then
'        IntPat_ID = My_Rst!pat_id2
' '       MsgBox IntPat_ID
'    End If
'    Con.Close
'
'End Sub
'
'Private Sub Flush_Pat_ID()
'
'    Dim My_Rst As New ADODB.Recordset
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'
'    My_Rst.Open "exec Pat_Info_SELECT1 1,'" & txtPat_ID1.Text & "'", Con
'    If My_Rst.EOF = False Then
'        IntPat_ID = My_Rst!pat_id
''        MsgBox IntPat_ID
'    End If
'    Con.Close
'
'End Sub
'
'Private Sub GATE_DT()
'
'    Dim My_Rst As New ADODB.Recordset
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'
'    My_Rst.Open "exec CR_Date", Con
'    If My_Rst.EOF = False Then
'        Dt.Value = My_Rst!crDate
'        DT_TM.Value = My_Rst!crDate
''        MsgBox IntPat_ID
'    End If
'    Con.Close
'
'End Sub
'
'Private Sub Cal_Dis()
'
'
'    DblDisc = Val(nbrTotal_Amt) * Val(nbrDisc_Per) / 100
'
'End Sub
'Private Sub Del_False_New_Doc()
'
'    Dim My_Rst As New ADODB.Recordset
'    Con.ConnectionString = strcn.Connection
'    Con.Open
'    Set cmd.ActiveConnection = Con
'
'    My_Rst.Open "exec Del_Doc_New 1,'" & "0" & "','" & u_id & "'", Con
'
'    Con.Close
'
'End Sub
'
'Private Sub txtS_Name_GotFocus()
'txtS_Name.BackColor = &HFFFFC0
'End Sub
'
'Private Sub txtS_Name_LostFocus()
'
'    txtS_Name.BackColor = vbWhite
'
'End Sub
Private Sub cmdExit_Click()

Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If




End Sub

Private Sub cmdSave_Click()
Call save_msg
End Sub
Private Sub save_msg()
If txtname = "" Then
MsgBox "Patient Name Required"
Exit Sub
End If
If txtAddr = "" Then
MsgBox "Address Required"
Exit Sub
End If
If txtAge = "" Then
MsgBox "Age Required"
Exit Sub
End If



End Sub
Private Sub Form_Load()

Combo1.AddItem "Cabin"
Combo1.Text = Combo1.List(0)
Combo3.Text = Combo3.List(0)
Combo5.Text = Combo5.List(0)


End Sub
