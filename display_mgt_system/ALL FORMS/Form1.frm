VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1 
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10800
   FillStyle       =   2  'Horizontal Line
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   4425
      Left            =   225
      ScaleHeight     =   4425
      ScaleWidth      =   9915
      TabIndex        =   64
      Top             =   1080
      Width           =   9915
      Begin VB.OptionButton Option6 
         Caption         =   "No"
         Height          =   195
         Left            =   5400
         TabIndex        =   81
         Top             =   325
         Width           =   780
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Yes"
         Height          =   195
         Left            =   4680
         TabIndex        =   80
         Top             =   325
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.TextBox T2 
         Height          =   285
         Index           =   5
         Left            =   2355
         MaxLength       =   15
         TabIndex        =   70
         Top             =   2610
         Width           =   2415
      End
      Begin VB.TextBox T2 
         Height          =   285
         Index           =   4
         Left            =   2355
         MaxLength       =   15
         TabIndex        =   69
         Top             =   2250
         Width           =   2415
      End
      Begin VB.TextBox T2 
         Height          =   285
         Index           =   3
         Left            =   2355
         MaxLength       =   15
         TabIndex        =   68
         Top             =   1890
         Width           =   2415
      End
      Begin VB.TextBox T2 
         Height          =   285
         Index           =   2
         Left            =   2355
         MaxLength       =   15
         TabIndex        =   67
         Top             =   1530
         Width           =   2415
      End
      Begin VB.TextBox T2 
         Height          =   285
         Index           =   1
         Left            =   2355
         MaxLength       =   15
         TabIndex        =   66
         Top             =   1170
         Width           =   2415
      End
      Begin VB.TextBox T2 
         Height          =   285
         Index           =   0
         Left            =   2355
         MaxLength       =   15
         TabIndex        =   65
         Top             =   810
         Width           =   2415
      End
      Begin MSMask.MaskEdBox NI 
         Height          =   255
         Left            =   2355
         TabIndex        =   71
         Top             =   450
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         ClipMode        =   1
         MaxLength       =   1
         Mask            =   "#"
         PromptChar      =   " "
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Display:"
         Height          =   195
         Left            =   3960
         TabIndex        =   79
         Top             =   325
         Width           =   555
      End
      Begin VB.Label IT 
         AutoSize        =   -1  'True
         Caption         =   "(6) Interest Type:"
         Height          =   195
         Index           =   5
         Left            =   315
         TabIndex        =   78
         Top             =   2610
         Width           =   1200
      End
      Begin VB.Label IT 
         AutoSize        =   -1  'True
         Caption         =   "(5) Interest Type:"
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   77
         Top             =   2250
         Width           =   1200
      End
      Begin VB.Label IT 
         AutoSize        =   -1  'True
         Caption         =   "(4) Interest Type:"
         Height          =   195
         Index           =   3
         Left            =   315
         TabIndex        =   76
         Top             =   1890
         Width           =   1200
      End
      Begin VB.Label IT 
         AutoSize        =   -1  'True
         Caption         =   "(3) Interest Type:"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   75
         Top             =   1530
         Width           =   1200
      End
      Begin VB.Label IT 
         AutoSize        =   -1  'True
         Caption         =   "(2) Interest Type:"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   74
         Top             =   1170
         Width           =   1200
      End
      Begin VB.Label IT 
         AutoSize        =   -1  'True
         Caption         =   "(1)  Interest Type:"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   73
         Top             =   810
         Width           =   1245
      End
      Begin VB.Label Label7 
         Caption         =   "Number of Interest"
         Height          =   375
         Index           =   1
         Left            =   315
         TabIndex        =   72
         Top             =   450
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   360
      ScaleHeight     =   4935
      ScaleWidth      =   10320
      TabIndex        =   10
      Top             =   600
      Width           =   10320
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   5055
         Left            =   0
         ScaleHeight     =   5055
         ScaleWidth      =   9975
         TabIndex        =   11
         Top             =   195
         Width           =   9975
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   19
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   49
            Top             =   3840
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   18
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   48
            Top             =   3480
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   17
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   47
            Top             =   3120
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   16
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   46
            Top             =   2760
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   15
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   45
            Top             =   2400
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   14
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   44
            Top             =   2040
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   13
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   43
            Top             =   1680
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   12
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   42
            Top             =   1320
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   11
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   41
            Top             =   960
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   10
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   40
            Top             =   600
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Yes"
            Height          =   195
            Left            =   4920
            TabIndex        =   39
            Top             =   360
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "No"
            Height          =   195
            Left            =   4320
            TabIndex        =   38
            Top             =   360
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   9
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   31
            Top             =   3960
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   8
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   30
            Top             =   3600
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   0
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   19
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   1
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   18
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   2
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   17
            Top             =   1440
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   3
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   16
            Top             =   1800
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   4
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   15
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   5
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   14
            Top             =   2520
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   6
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   13
            Top             =   2880
            Width           =   2415
         End
         Begin VB.TextBox T1 
            Height          =   285
            Index           =   7
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   12
            Top             =   3240
            Width           =   2415
         End
         Begin MSMask.MaskEdBox MB1 
            Height          =   255
            Left            =   2280
            TabIndex        =   20
            Top             =   360
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(18) Currency Name:"
            Height          =   195
            Index           =   19
            Left            =   5280
            TabIndex        =   59
            Top             =   3240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(17) Currency Name:"
            Height          =   195
            Index           =   18
            Left            =   5280
            TabIndex        =   58
            Top             =   2880
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(16) Currency Name:"
            Height          =   195
            Index           =   17
            Left            =   5280
            TabIndex        =   57
            Top             =   2520
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(15) Currency Name:"
            Height          =   195
            Index           =   16
            Left            =   5280
            TabIndex        =   56
            Top             =   2160
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(14) Currency Name:"
            Height          =   195
            Index           =   15
            Left            =   5280
            TabIndex        =   55
            Top             =   1800
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(13) Currency Name:"
            Height          =   195
            Index           =   14
            Left            =   5280
            TabIndex        =   54
            Top             =   1440
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(12) Currency Name:"
            Height          =   195
            Index           =   13
            Left            =   5280
            TabIndex        =   53
            Top             =   1080
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(11) Currency Name:"
            Height          =   195
            Index           =   12
            Left            =   5280
            TabIndex        =   52
            Top             =   720
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(19) Currency Name:"
            Height          =   195
            Index           =   11
            Left            =   5280
            TabIndex        =   51
            Top             =   3600
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(20) Currency Name:"
            Height          =   195
            Index           =   10
            Left            =   5280
            TabIndex        =   50
            Top             =   3960
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Cheque:"
            Height          =   195
            Left            =   3600
            TabIndex        =   37
            Top             =   360
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(10) Currency Name:"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   33
            Top             =   3960
            Width           =   1455
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(9) Currency Name:"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   32
            Top             =   3600
            Width           =   1365
         End
         Begin VB.Label Label7 
            Caption         =   "Number of Currency:"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(1) Currency Name:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   28
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(2) Currency Name:"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   1080
            Width           =   1365
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(3) Currency Name:"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   26
            Top             =   1440
            Width           =   1365
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(4) Currency Name:"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   25
            Top             =   1800
            Width           =   1365
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(5) Currency Name:"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   24
            Top             =   2160
            Width           =   1365
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(6) Currency Name:"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   23
            Top             =   2520
            Width           =   1365
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(7) Currency Name:"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   22
            Top             =   2880
            Width           =   1365
         End
         Begin VB.Label CN 
            AutoSize        =   -1  'True
            Caption         =   "(8) Currency Name:"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   21
            Top             =   3240
            Width           =   1365
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   360
      ScaleHeight     =   4935
      ScaleWidth      =   10005
      TabIndex        =   1
      Top             =   810
      Width           =   10005
      Begin MSComDlg.CommonDialog cd5 
         Left            =   1800
         Top             =   3600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   3
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1320
         Width           =   4455
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   3360
         Top             =   3120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CD2 
         Left            =   3120
         Top             =   3120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CD3 
         Left            =   3600
         Top             =   2640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CD4 
         Left            =   5280
         Top             =   3120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cd6 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   6840
         MouseIcon       =   "Form1.frx":0442
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":074C
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6840
         MouseIcon       =   "Form1.frx":0E8E
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":1198
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6360
         TabIndex        =   9
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6360
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Company's Name:"
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Slogan/Address:"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   1320
         Width           =   1185
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6240
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6240
         TabIndex        =   4
         Top             =   1320
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   570
      ScaleHeight     =   4935
      ScaleWidth      =   9735
      TabIndex        =   60
      Top             =   720
      Width           =   9735
      Begin VB.OptionButton Option4 
         Caption         =   "No"
         Height          =   495
         Left            =   2040
         TabIndex        =   63
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Yes"
         Height          =   495
         Left            =   1320
         TabIndex        =   62
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Display Date:"
         Height          =   195
         Left            =   240
         TabIndex        =   61
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.PictureBox Picture4 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   10740
      TabIndex        =   34
      Top             =   6255
      Width           =   10800
      Begin VB.CommandButton Command2 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   1200
         TabIndex        =   36
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   975
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6375
      Left            =   135
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   11245
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Company Setup"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Currency Setup"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Option"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Interest Setup"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c1, c2, c3, c4

Private Sub Command1_Click()

On Error Resume Next
Picture1.Visible = True
Picture2.Visible = True

If Not IsNumeric(MB1.Text) Then
MB1.SetFocus
Exit Sub
End If

If MB1.Text > 10 Then
MsgBox "Invalid Input", vbOKOnly + vbInformation, "ERDS"
Exit Sub
End If

If Len(Trim(Text1.Text)) = 0 Then
Text1.SetFocus
Exit Sub
End If

If Len(Trim(Text2.Text)) = 0 Then
Text2.SetFocus
Exit Sub
End If

If Not IsNumeric(MB1.Text) Then
MB1.SetFocus
Exit Sub
End If

If cd5.FontName = "" Then
Exit Sub
End If

If cd6.FontName = "" Then
Exit Sub
End If

If Option5.Value = True Then
dcr = 1
Else
dcr = 0
End If

With Opnrs
.AddNew
.Fields(0) = Text1.Text
.Fields(1) = Text2.Text

If Len(Trim(c1)) > 0 Then
.Fields(2) = c1
Else
.Fields(2) = &H80000007
End If

If Len(Trim(c2)) > 0 Then
.Fields(3) = c2
Else
.Fields(3) = &H80000009
End If

If Len(Trim(c3)) > 0 Then
.Fields(4) = c3
Else
.Fields(4) = &H80000007
End If

If Len(Trim(c4)) > 0 Then
.Fields(5) = c4
Else
.Fields(5) = &H80000009
End If

.Fields(6) = MB1.Text

If Len(Trim(T1(0).Text)) > 0 Then
.Fields(7) = T1(0).Text
Else
.Fields(7) = " "
End If

If Len(Trim(T1(1).Text)) > 0 Then
.Fields(8) = T1(1).Text
Else
.Fields(8) = " "
End If

If Len(Trim(T1(2).Text)) > 0 Then
.Fields(9) = T1(2).Text
Else
.Fields(9) = " "
End If

If Len(Trim(T1(3).Text)) > 0 Then
.Fields(10) = T1(3).Text
Else
.Fields(10) = " "
End If

If Len(Trim(T1(4).Text)) > 0 Then
.Fields(11) = T1(4).Text
Else
.Fields(11) = " "
End If


If Len(Trim(T1(5).Text)) > 0 Then
.Fields(12) = T1(5).Text
Else
.Fields(12) = " "
End If

If Len(Trim(T1(6).Text)) > 0 Then
.Fields(13) = T1(6).Text
Else
.Fields(13) = " "
End If

If Len(Trim(T1(7).Text)) > 0 Then
.Fields(14) = T1(7).Text
Else
.Fields(14) = " "
End If

If Len(Trim(T1(8).Text)) > 0 Then
.Fields(15) = T1(8).Text
Else
.Fields(15) = " "
End If

If Len(Trim(T1(9).Text)) > 0 Then
.Fields(16) = T1(9).Text
Else
.Fields(16) = " "
End If

If Len(Trim(T1(10).Text)) > 0 Then
.Fields(17) = T1(10).Text
Else
.Fields(17) = " "
End If

If Len(Trim(T1(11).Text)) > 0 Then
.Fields(18) = T1(11).Text
Else
.Fields(18) = " "
End If

If Len(Trim(T1(12).Text)) > 0 Then
.Fields(19) = T1(12).Text
Else
.Fields(19) = " "
End If

If Len(Trim(T1(13).Text)) > 0 Then
.Fields(20) = T1(13).Text
Else
.Fields(20) = " "
End If

If Len(Trim(T1(14).Text)) > 0 Then
.Fields(21) = T1(14).Text
Else
.Fields(21) = " "
End If

If Len(Trim(T1(15).Text)) > 0 Then
.Fields(22) = T1(15).Text
Else
.Fields(22) = " "
End If

If Len(Trim(T1(16).Text)) > 0 Then
.Fields(23) = T1(16).Text
Else
.Fields(23) = " "
End If

If Len(Trim(T1(17).Text)) > 0 Then
.Fields(24) = T1(17).Text
Else
.Fields(24) = " "
End If


If Len(Trim(T1(18).Text)) > 0 Then
.Fields(25) = T1(18).Text
Else
.Fields(25) = " "
End If


If Len(Trim(T1(19).Text)) > 0 Then
.Fields(26) = T1(19).Text
Else
.Fields(26) = " "
End If

If Option1.Value = True Then
.Fields(27) = 1
Else
.Fields(27) = 0
End If

If Option2.Value = True Then
.Fields(28) = 1
Else
.Fields(28) = 0
End If

.Fields("dcr") = dcr

.Fields("cr1") = T2(0).Text
.Fields("cr2") = T2(1).Text
.Fields("cr3") = T2(2).Text
.Fields("cr4") = T2(3).Text
.Fields("cr5") = T2(4).Text
.Fields("cr6") = T2(5).Text
.Fields("IRC") = NI.Text

.Fields("FontName1") = cd5.FontName
.Fields("FontName2") = cd6.FontName
'.Fields("FontSize1") = cd5.FontSize
'.Fields("FontSize2") = cd6.FontSize
.Fields("FontBold1") = cd5.FontBold
.Fields("FontBold2") = cd6.FontBold
If Option3.Value = True Then
.Fields("DateOp") = 1
End If
If Option4.Value = True Then
.Fields("DateOp") = 0
End If

.Update
End With
Unload Me
Form2.Show
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
TabStrip1.Tabs(1).Selected = True
Option2.Value = True
Option4.Value = True
End Sub

Private Sub Image1_Click()
cd5.CancelError = False
cd5.Flags = 1
cd5.ShowFont
End Sub

Private Sub Image2_Click()
cd6.CancelError = False
cd6.Flags = 1
cd6.ShowFont
End Sub

Private Sub Label3_Click()
CD1.CancelError = False
CD1.Flags = 0
CD1.ShowColor
c1 = CD1.Color
Label3.BackColor = c1
End Sub

Private Sub Label4_Click()
CD2.CancelError = False
CD2.Flags = 0
CD2.ShowColor
c2 = CD2.Color
Label4.BackColor = c2
End Sub

Private Sub Label5_Click()
CD3.CancelError = False
CD3.Flags = 0
CD3.ShowColor
c3 = CD3.Color
Label5.BackColor = c3
End Sub

Private Sub Label6_Click()
CD4.CancelError = False
CD4.Flags = 0
CD4.ShowColor
c4 = CD4.Color
Label6.BackColor = c4
End Sub

Private Sub MB1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KeyDown Or KeyCode = vbKeyReturn Then
T1(0).SetFocus
End If
End Sub

Private Sub NI_GotFocus()
NI.SelStart = 0
NI.SelLength = Len(NI)
End Sub

Private Sub NI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KeyDown Or KeyCode = vbKeyReturn Then
If T2(0).Visible = True Then
T2(0).SetFocus
End If
End If
End Sub

Private Sub NI_KeyPress(KeyAscii As Integer)
If KeyAscii > 54 Or KeyAscii < 49 Then
KeyAscii = 0
End If
End Sub

Private Sub NI_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
'If NI > 0 Then
'For ixt = 0 To NI - 1
'IT(ixt).Visible = False
'T2(ixt).Visible = False
'Next
'For ixt = 0 To NI - 1
'IT(ixt).Visible = True
'T2(ixt).Visible = True
'Next
'End If
End Sub

Private Sub NI_LostFocus()
If IsNumeric(NI) Then
If NI > 6 Or NI < 1 Then
NI.SetFocus
Exit Sub
End If
Else
NI.SetFocus
Exit Sub
End If

'If IsNumeric(NI) Then
'For ixt = 0 To NI - 1
'IT(ixt).Visible = False
'T2(ixt).Visible = False
'Next
'For ixt = 0 To NI - 1
'IT(ixt).Visible = True
'T2(ixt).Visible = True
'Next
'End If

End Sub

Private Sub T1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
If Index < 19 Then
T1(Index + 1).SetFocus
End If
End If

End Sub

Private Sub T2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
If Index < 5 Then
T2(Index + 1).SetFocus
End If
End If
End Sub

Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.Index
Case 1
Picture1.Visible = True
Picture2.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Case 2
Picture2.Visible = True
Picture1.Visible = False
Picture5.Visible = False
Picture6.Visible = False
MB1.SetFocus
Case 3
Picture2.Visible = False
Picture1.Visible = False
Picture5.Visible = True
Picture6.Visible = False
Case 4
Picture1.Visible = False
Picture2.Visible = False
Picture5.Visible = False
Picture6.Visible = True
'For ixy = 0 To 5
'IT(ixy).Visible = False
'T2(ixy).Visible = False
'Next
End Select
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
If Len(Trim(Text1.Text)) > 0 Then
Text2.SetFocus
End If
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
If Len(Trim(Text2.Text)) > 0 Then
Command1.SetFocus
End If
End If
End Sub


