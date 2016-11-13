VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmSectionInfo 
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H8000000C&
      Caption         =   "New"
      Height          =   435
      Left            =   4800
      TabIndex        =   22
      Top             =   4980
      Width           =   945
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   435
      Left            =   5730
      TabIndex        =   21
      Top             =   4980
      Width           =   945
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Exit"
      Height          =   435
      Left            =   6660
      TabIndex        =   20
      Top             =   4980
      Width           =   945
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2355
      Left            =   0
      TabIndex        =   19
      Top             =   2610
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4154
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.Frame Frame3 
      Height          =   1905
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   690
      Width           =   7575
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   5
         Left            =   1290
         TabIndex        =   16
         Top             =   1530
         Width           =   6225
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   4
         Left            =   1290
         TabIndex        =   15
         Top             =   1200
         Width           =   6225
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   3
         Left            =   1290
         TabIndex        =   13
         Top             =   870
         Width           =   6225
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   2
         Left            =   5790
         TabIndex        =   12
         Top             =   510
         Width           =   1725
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   1
         Left            =   1290
         TabIndex        =   10
         Top             =   540
         Width           =   3345
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1290
         TabIndex        =   7
         Top             =   180
         Width           =   3345
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   0
         Left            =   5790
         TabIndex        =   5
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Monitor 2"
         Height          =   195
         Index           =   5
         Left            =   60
         TabIndex        =   18
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Monitor 1"
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   17
         Top             =   1230
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Teacher"
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   14
         Top             =   900
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room No"
         Height          =   195
         Index           =   2
         Left            =   4950
         TabIndex        =   11
         Top             =   540
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section Name"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   9
         Top             =   570
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section ID"
         Height          =   195
         Index           =   0
         Left            =   4920
         TabIndex        =   8
         Top             =   210
         Width           =   750
      End
      Begin VB.Label Label2 
         Caption         =   "Class"
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   210
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   705
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   30
         Left            =   0
         TabIndex        =   3
         Top             =   660
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   2
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   150
         Width           =   2235
      End
   End
End
Attribute VB_Name = "FrmSectionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rs1 As New Recordset
Set rs1 = GetData("select ClassId,ClassName from ClassInfo")
If Not rs1.EOF Then
Do Until rs1.EOF
       Combo1.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop

End If
With MSFlexGrid1
    .Rows = 1
    .Cols = 6
    .Col = 0: .Text = "               Section ID #"
    .Col = 1: .Text = " Section Name "
    .Col = 2: .Text = " Room No "
    .Col = 3: .Text = " Class Teaher "
    .Col = 4: .Text = " Monitor's Name "
    .Col = 5: .Text = " Monitor's Name(Alternative) "
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 3000
    .ColWidth(2) = 1500
    .ColWidth(3) = 4000
    .ColWidth(4) = 4000
    .ColWidth(5) = 4000
   
End With

End Sub
