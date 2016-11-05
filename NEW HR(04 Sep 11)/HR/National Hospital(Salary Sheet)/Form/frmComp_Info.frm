VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form16 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Company Information"
   ClientHeight    =   5055
   ClientLeft      =   2490
   ClientTop       =   1965
   ClientWidth     =   7800
   Icon            =   "frmComp_Info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7800
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4110
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   7620
      Begin VB.TextBox txtNotes 
         Height          =   330
         Left            =   1170
         TabIndex        =   19
         Top             =   3555
         Width           =   6180
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   1170
         TabIndex        =   18
         Top             =   3060
         Width           =   6180
      End
      Begin VB.TextBox txtFax 
         Height          =   330
         Left            =   1170
         TabIndex        =   17
         Top             =   2655
         Width           =   4110
      End
      Begin VB.TextBox txtPhone 
         Height          =   330
         Left            =   1170
         TabIndex        =   16
         Top             =   2205
         Width           =   4110
      End
      Begin VB.TextBox txtAddress 
         Height          =   735
         Left            =   1170
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   1305
         Width           =   4065
      End
      Begin VB.ComboBox cboOrg_Type 
         Height          =   315
         Left            =   1170
         TabIndex        =   14
         Top             =   810
         Width           =   4110
      End
      Begin VB.TextBox txtname 
         Height          =   420
         Left            =   1170
         TabIndex        =   13
         Top             =   270
         Width           =   6135
      End
      Begin VB.Image imgLogo 
         Height          =   1645
         Left            =   5535
         Picture         =   "frmComp_Info.frx":08CA
         Stretch         =   -1  'True
         ToolTipText     =   "   Click to load picture  "
         Top             =   1215
         Width           =   1770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   315
         TabIndex        =   12
         Top             =   1215
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   315
         TabIndex        =   11
         Top             =   765
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   315
         TabIndex        =   10
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   315
         TabIndex        =   9
         Top             =   3510
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   315
         TabIndex        =   8
         Top             =   3060
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   315
         TabIndex        =   7
         Top             =   2655
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   315
         TabIndex        =   6
         Top             =   2250
         Width           =   465
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Company Logo"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5580
         TabIndex        =   5
         Top             =   810
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   3855
      Picture         =   "frmComp_Info.frx":18B0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4410
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   1215
      Picture         =   "frmComp_Info.frx":349A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4410
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   2535
      Picture         =   "frmComp_Info.frx":4E2C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4410
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   5175
      Picture         =   "frmComp_Info.frx":67BE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4410
      Width           =   1185
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Comp_Info As New Company_Info
Dim Track_Id As Long
Dim Default_Pic_Path As String
Dim New_Pic_Path As String
Private Sub cmdClear_Click()
    Clear_Screen
End Sub

Private Sub cmdClose_Click()
    Close_Msg Me
End Sub


Private Sub cmdSave_Click()
    
    With Comp_Info
        .ConnString = strCN.Connection
        .Co_Nm = txtname
        .Co_Type = cboOrg_Type
        .Address = txtAddress
        .Phone = txtPhone
        .Fax = txtFax
        .E_mail = txtEmail
        .Notes = txtNotes
        .Logo = New_Pic_Path
        .Save
    End With
    
    Track_Id = 0
    Clear_Screen
    Show_Data
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

On Error Resume Next

Screen_Position Me
    
    With cboOrg_Type
        .AddItem "Private Company"
        .AddItem "Limited Company"
        .AddItem "Autonomous"
        .AddItem "Govt. Organization"
        .AddItem "Non Govt. Organization"
    End With
    
    Default_Pic_Path = App.Path + "\Default_Pic.bmp"
    
    Show_Data
  Exit Sub
    
End Sub
Public Sub Show_Data()
   ' On Error Resume Next
        
    With Comp_Info
            .ConnString = strCN.Connection
            .Get_Company_Info
        txtname = .Co_Nm
        cboOrg_Type = .Co_Type
        txtAddress = .Address
        txtPhone = .Phone
        txtFax = .Fax
        txtEmail = .E_mail
        txtNotes = .Notes
        If Not .Logo = Empty Then
            imgLogo.Picture = LoadPicture(.Logo)
            Default_Pic_Path = .Logo
            New_Pic_Path = .Logo
        Else
            imgLogo.Picture = LoadPicture(Default_Pic_Path)
        End If
        
        
    End With
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Destroy Me
End Sub
Public Sub Load_Photo(ComDiag As CommonDialog, Img As Image, Optional Photo_Path As String)

Dim resp As String


Start:  With ComDiag
            .Filter = "Photograph,*.bmp;*.jpg;*.gif|*.bmp;*.jpg;*.gif"
            .Action = 1
                If .filename = "" Then
                    Exit Sub
                Else
                    New_Pic_Path = .filename
                    Img.Picture = LoadPicture(New_Pic_Path)
                End If
        End With
        
'------------------------------------------------------

        resp = MsgBox("           Is it the right Picture ?", vbYesNoCancel + vbQuestion + vbDefaultButton2, "Message")
        
        
        If resp = vbCancel Then
            Img.Picture = LoadPicture(Default_Pic_Path)
            Exit Sub
        End If
        
        If resp = vbNo Then
            Img.Picture = LoadPicture(Default_Pic_Path)
            GoTo Start
            Exit Sub
        End If
        
        If resp = vbYes Then
                New_Pic_Path = ComDiag.filename
            Exit Sub
        
        End If
        

End Sub

Private Sub imgLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
    
        Call Load_Photo(CommonDialog1, imgLogo, Default_Pic_Path)
    End If
  
End Sub
