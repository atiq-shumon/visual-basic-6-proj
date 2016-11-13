VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmstudentInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H8000000C&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8010
      TabIndex        =   105
      ToolTipText     =   "Click to Edit"
      Top             =   6960
      Width           =   945
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9990
      TabIndex        =   39
      ToolTipText     =   "Click to Exit"
      Top             =   6960
      Width           =   945
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H8000000C&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9000
      TabIndex        =   38
      ToolTipText     =   "Click to Delete"
      Top             =   6960
      Width           =   945
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7020
      TabIndex        =   21
      ToolTipText     =   "Click to Save"
      Top             =   6960
      Width           =   945
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H8000000C&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6030
      TabIndex        =   0
      ToolTipText     =   "Click to insert new information"
      Top             =   6960
      Width           =   945
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      FillColor       =   &H80000002&
      ForeColor       =   &H80000017&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11025
      TabIndex        =   41
      Top             =   0
      Width           =   11085
      Begin TabDlg.SSTab SSTab2 
         Height          =   300
         Left            =   30
         TabIndex        =   42
         Top             =   690
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   529
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "frmstudentInfo.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Tab 1"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Tab 2"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEF0F7&
         Height          =   285
         Left            =   3900
         TabIndex        =   84
         Top             =   150
         Width           =   2265
      End
      Begin VB.Image Image1 
         Height          =   930
         Left            =   -90
         Picture         =   "frmstudentInfo.frx":001C
         Stretch         =   -1  'True
         Top             =   -90
         Width           =   11115
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6075
      Left            =   0
      TabIndex        =   40
      Top             =   750
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   10716
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "Personal Information"
      TabPicture(0)   =   "frmstudentInfo.frx":CEC1
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture2"
      Tab(0).Control(1)=   "Frame13"
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(5)=   "Image3"
      Tab(0).Control(6)=   "Image4"
      Tab(0).Control(7)=   "Shape2"
      Tab(0).Control(8)=   "Shape1(1)"
      Tab(0).Control(9)=   "Shape1(0)"
      Tab(0).Control(10)=   "Image2"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Pysical Status"
      TabPicture(1)   =   "frmstudentInfo.frx":CEDD
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "MSFlexGrid1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Address Information"
      TabPicture(2)   =   "frmstudentInfo.frx":CEF9
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000007&
         Height          =   1125
         Left            =   -68430
         ScaleHeight     =   1065
         ScaleWidth      =   1455
         TabIndex        =   91
         Top             =   480
         Width           =   1515
      End
      Begin VB.Frame Frame13 
         Caption         =   "Previous School Information(If any)"
         ForeColor       =   &H00C00000&
         Height          =   1335
         Left            =   -74910
         TabIndex        =   85
         Top             =   3480
         Width           =   10935
         Begin VB.ComboBox cmdTobeAditted 
            Height          =   315
            ItemData        =   "frmstudentInfo.frx":CF15
            Left            =   1410
            List            =   "frmstudentInfo.frx":CF40
            TabIndex        =   95
            Top             =   630
            Width           =   1155
         End
         Begin VB.ComboBox cmdPrivClass 
            Height          =   315
            ItemData        =   "frmstudentInfo.frx":CF94
            Left            =   1410
            List            =   "frmstudentInfo.frx":CFBF
            TabIndex        =   94
            Top             =   270
            Width           =   1155
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   24
            Left            =   1410
            MaxLength       =   80
            TabIndex        =   19
            ToolTipText     =   "Insert Certificate No"
            Top             =   990
            Width           =   7215
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   23
            Left            =   9210
            MaxLength       =   20
            TabIndex        =   17
            ToolTipText     =   "Insert legal guadian's name"
            Top             =   240
            Width           =   1665
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   22
            Left            =   3660
            MaxLength       =   150
            TabIndex        =   18
            ToolTipText     =   "Insert Previous School Address"
            Top             =   630
            Width           =   7215
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   19
            Left            =   3660
            MaxLength       =   150
            TabIndex        =   16
            ToolTipText     =   "Insert Previous School  name"
            Top             =   240
            Width           =   4995
         End
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   285
            Left            =   9210
            TabIndex        =   20
            ToolTipText     =   "Insert marraige date"
            Top             =   990
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd-mmm-yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   195
            Left            =   8640
            TabIndex        =   93
            Top             =   1020
            Width           =   345
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Certificate No "
            Height          =   195
            Left            =   90
            TabIndex        =   92
            Top             =   990
            Width           =   1005
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            Height          =   195
            Left            =   8640
            TabIndex        =   90
            Top             =   270
            Width           =   465
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "School address"
            Height          =   195
            Left            =   2550
            TabIndex        =   89
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To be Admitted "
            Height          =   195
            Left            =   90
            TabIndex        =   88
            Top             =   630
            Width           =   1125
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Previous Class"
            Height          =   195
            Left            =   90
            TabIndex        =   87
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "School Name"
            Height          =   195
            Left            =   2610
            TabIndex        =   86
            Top             =   270
            Width           =   960
         End
      End
      Begin VB.Frame Frame5 
         Height          =   5145
         Left            =   -74880
         TabIndex        =   58
         Top             =   480
         Width           =   10965
         Begin VB.Frame Frame9 
            Caption         =   "Emmergency Contact"
            ForeColor       =   &H00C00000&
            Height          =   1335
            Left            =   60
            TabIndex        =   70
            Top             =   3810
            Width           =   11115
            Begin VB.TextBox txtfields 
               Height          =   285
               Index           =   18
               Left            =   7080
               MaxLength       =   50
               TabIndex        =   37
               ToolTipText     =   "Insert Mobile No"
               Top             =   870
               Width           =   3585
            End
            Begin VB.TextBox txtfields 
               Height          =   285
               Index           =   17
               Left            =   1140
               MaxLength       =   50
               TabIndex        =   36
               ToolTipText     =   "Insert Contact No"
               Top             =   870
               Width           =   3645
            End
            Begin VB.TextBox txtfields 
               Height          =   465
               Index           =   16
               Left            =   1140
               MaxLength       =   150
               TabIndex        =   35
               ToolTipText     =   "Insert Immergency Contact Address"
               Top             =   240
               Width           =   9525
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile"
               Height          =   195
               Left            =   6180
               TabIndex        =   73
               Top             =   900
               Width           =   465
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone"
               Height          =   195
               Left            =   90
               TabIndex        =   72
               Top             =   840
               Width           =   465
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               Height          =   195
               Left            =   120
               TabIndex        =   71
               Top             =   270
               Width           =   570
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Contact Information"
            ForeColor       =   &H00C00000&
            Height          =   675
            Left            =   60
            TabIndex        =   67
            Top             =   3030
            Width           =   11145
            Begin VB.TextBox txtfields 
               Height          =   285
               Index           =   15
               Left            =   7080
               MaxLength       =   50
               TabIndex        =   34
               ToolTipText     =   "Insert E-Mail Address"
               Top             =   240
               Width           =   3615
            End
            Begin VB.TextBox txtfields 
               Height          =   285
               Index           =   14
               Left            =   1140
               MaxLength       =   50
               TabIndex        =   33
               ToolTipText     =   "Insert Contact No"
               Top             =   240
               Width           =   3675
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "E-Mail"
               Height          =   195
               Left            =   6150
               TabIndex        =   69
               Top             =   300
               Width           =   435
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "Phone No"
               Height          =   195
               Left            =   120
               TabIndex        =   68
               Top             =   270
               Width           =   720
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Present"
            ForeColor       =   &H00C00000&
            Height          =   1395
            Left            =   60
            TabIndex        =   63
            Top             =   1590
            Width           =   11055
            Begin VB.TextBox txtfields 
               Height          =   465
               Index           =   13
               Left            =   1140
               MaxLength       =   150
               TabIndex        =   30
               ToolTipText     =   "Insert Streert name of Present Address"
               Top             =   210
               Width           =   9555
            End
            Begin VB.ComboBox Combo4 
               Height          =   315
               Left            =   7080
               Style           =   2  'Dropdown List
               TabIndex        =   32
               ToolTipText     =   "Select Country NAme"
               Top             =   840
               Width           =   3615
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Left            =   1140
               Style           =   2  'Dropdown List
               TabIndex        =   31
               ToolTipText     =   "Select District Name"
               Top             =   840
               Width           =   3705
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Street Name"
               Height          =   195
               Left            =   90
               TabIndex        =   66
               Top             =   255
               Width           =   885
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Country "
               Height          =   195
               Left            =   6090
               TabIndex        =   65
               Top             =   900
               Width           =   585
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "District"
               Height          =   195
               Left            =   120
               TabIndex        =   64
               Top             =   870
               Width           =   480
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Permanent"
            ForeColor       =   &H00C00000&
            Height          =   1395
            Left            =   90
            TabIndex        =   59
            Top             =   150
            Width           =   10875
            Begin VB.ComboBox Combo2 
               Height          =   315
               Left            =   1110
               Style           =   2  'Dropdown List
               TabIndex        =   28
               ToolTipText     =   "SelectDistrict Name"
               Top             =   750
               Width           =   3765
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   7110
               Style           =   2  'Dropdown List
               TabIndex        =   29
               ToolTipText     =   "SelectCountry Name"
               Top             =   750
               Width           =   3585
            End
            Begin VB.TextBox txtfields 
               Height          =   405
               Index           =   12
               Left            =   1110
               MaxLength       =   150
               TabIndex        =   27
               ToolTipText     =   "Insert Street name of Parmanent Address"
               Top             =   180
               Width           =   9555
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "District"
               Height          =   195
               Left            =   90
               TabIndex        =   62
               Top             =   810
               Width           =   480
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Country "
               Height          =   195
               Left            =   6090
               TabIndex        =   61
               Top             =   840
               Width           =   585
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Street Name"
               Height          =   195
               Left            =   90
               TabIndex        =   60
               Top             =   225
               Width           =   885
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4395
         Left            =   90
         TabIndex        =   57
         Top             =   1620
         Width           =   10965
         _ExtentX        =   19341
         _ExtentY        =   7752
         _Version        =   393216
         FixedCols       =   0
      End
      Begin VB.Frame Frame4 
         Caption         =   "Health Information"
         ForeColor       =   &H00C00000&
         Height          =   1095
         Left            =   120
         TabIndex        =   51
         Top             =   480
         Width           =   10905
         Begin VB.ComboBox cmbVaccineName 
            Height          =   315
            ItemData        =   "frmstudentInfo.frx":D013
            Left            =   1200
            List            =   "frmstudentInfo.frx":D015
            Style           =   2  'Dropdown List
            TabIndex        =   106
            ToolTipText     =   "Select Country of birth"
            Top             =   660
            Width           =   5985
         End
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   285
            Left            =   7680
            TabIndex        =   25
            ToolTipText     =   "Insert Vaccine Date"
            Top             =   660
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd-mmm-yy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   10
            Left            =   5850
            TabIndex        =   24
            ToolTipText     =   "Insert student's Blood Group"
            Top             =   240
            Width           =   1305
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   9
            Left            =   3366
            TabIndex        =   23
            ToolTipText     =   "Insert student's Weight"
            Top             =   240
            Width           =   1275
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   8
            Left            =   1200
            TabIndex        =   22
            ToolTipText     =   "Insert student's hight"
            Top             =   240
            Width           =   1275
         End
         Begin MSMask.MaskEdBox MaskEdBox4 
            Height          =   285
            Left            =   9720
            TabIndex        =   26
            ToolTipText     =   "Insert NEw VAccine Date"
            Top             =   660
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd-mmm-yy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Next Date"
            Height          =   195
            Left            =   8940
            TabIndex        =   74
            Top             =   705
            Width           =   720
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   195
            Left            =   7230
            TabIndex        =   56
            Top             =   705
            Width           =   345
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vaccine Name"
            Height          =   195
            Left            =   135
            TabIndex        =   55
            Top             =   705
            Width           =   1050
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Blood Group"
            Height          =   195
            Left            =   4803
            TabIndex        =   54
            Top             =   270
            Width           =   885
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Weight"
            Height          =   195
            Left            =   2694
            TabIndex        =   53
            Top             =   270
            Width           =   510
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Height"
            Height          =   195
            Left            =   630
            TabIndex        =   52
            Top             =   270
            Width           =   465
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Others"
         ForeColor       =   &H00C00000&
         Height          =   1305
         Left            =   -74910
         TabIndex        =   50
         Top             =   4800
         Width           =   10935
         Begin VB.Frame Frame12 
            Caption         =   "Computer/Internet"
            ForeColor       =   &H00800000&
            Height          =   555
            Left            =   0
            TabIndex        =   83
            Top             =   690
            Width           =   10935
            Begin VB.CheckBox cmdcomputer 
               Caption         =   "Computer At Home?"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   150
               TabIndex        =   14
               ToolTipText     =   "Click if student has PC"
               Top             =   180
               Width           =   2265
            End
            Begin VB.CheckBox cmdInternet 
               Caption         =   "Having Internet Connection?"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2580
               TabIndex        =   15
               ToolTipText     =   "Click if Student has Internet connection at home?"
               Top             =   150
               Width           =   3345
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Siblings Number"
            ForeColor       =   &H00800000&
            Height          =   585
            Left            =   0
            TabIndex        =   80
            Top             =   180
            Width           =   2535
            Begin VB.TextBox txtfields 
               Height          =   285
               Index           =   6
               Left            =   1800
               MaxLength       =   1
               TabIndex        =   10
               Text            =   "0"
               ToolTipText     =   "Insert No of Sisters"
               Top             =   210
               Width           =   645
            End
            Begin VB.TextBox txtfields 
               Height          =   285
               Index           =   5
               Left            =   660
               MaxLength       =   1
               TabIndex        =   9
               Text            =   "0"
               ToolTipText     =   "Insert No of Brothers"
               Top             =   210
               Width           =   645
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sister"
               Height          =   195
               Left            =   1380
               TabIndex        =   82
               Top             =   225
               Width           =   390
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Brother"
               Height          =   195
               Left            =   90
               TabIndex        =   81
               Top             =   255
               Width           =   510
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Birth Info"
            ForeColor       =   &H00800000&
            Height          =   555
            Left            =   2550
            TabIndex        =   76
            Top             =   180
            Width           =   8385
            Begin VB.ComboBox Combo5 
               Height          =   315
               ItemData        =   "frmstudentInfo.frx":D017
               Left            =   690
               List            =   "frmstudentInfo.frx":D019
               Style           =   2  'Dropdown List
               TabIndex        =   13
               ToolTipText     =   "Select Country of birth"
               Top             =   195
               Width           =   2205
            End
            Begin VB.TextBox txtfields 
               Height          =   285
               Index           =   7
               Left            =   6660
               MaxLength       =   15
               TabIndex        =   12
               ToolTipText     =   "Insert religion"
               Top             =   165
               Width           =   1665
            End
            Begin MSMask.MaskEdBox MaskEdBox2 
               Height          =   285
               Left            =   3660
               TabIndex        =   11
               ToolTipText     =   "Insert date of Birth "
               Top             =   195
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd-mmm-yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Birth Date"
               Height          =   195
               Left            =   2895
               TabIndex        =   79
               Top             =   225
               Width           =   705
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Country"
               Height          =   195
               Left            =   90
               TabIndex        =   78
               Top             =   225
               Width           =   540
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Religion"
               Height          =   195
               Left            =   6060
               TabIndex        =   77
               Top             =   195
               Width           =   570
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Parent's Status"
         ForeColor       =   &H00C00000&
         Height          =   1965
         Left            =   -74910
         TabIndex        =   46
         Top             =   1650
         Width           =   10935
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   21
            Left            =   9180
            MaxLength       =   15
            TabIndex        =   102
            Text            =   "0"
            ToolTipText     =   "Insert Monthly Average Income of Mother"
            Top             =   690
            Width           =   1665
         End
         Begin VB.ComboBox cmdProfMother 
            Height          =   315
            ItemData        =   "frmstudentInfo.frx":D01B
            Left            =   6930
            List            =   "frmstudentInfo.frx":D043
            TabIndex        =   101
            Top             =   690
            Width           =   1665
         End
         Begin VB.ComboBox cmdProfFather 
            Height          =   315
            ItemData        =   "frmstudentInfo.frx":D0CC
            Left            =   6930
            List            =   "frmstudentInfo.frx":D0F1
            TabIndex        =   100
            Top             =   270
            Width           =   1665
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   20
            Left            =   9180
            MaxLength       =   15
            TabIndex        =   97
            Text            =   "0"
            ToolTipText     =   "Insert Monthly Average Income of Father"
            Top             =   270
            Width           =   1665
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   4
            Left            =   1410
            MaxLength       =   80
            TabIndex        =   6
            ToolTipText     =   "Insert legal guadian's name"
            Top             =   1470
            Width           =   9435
         End
         Begin VB.CheckBox cmdFM 
            Caption         =   "Is Father Or Mother Late ?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3630
            TabIndex        =   8
            ToolTipText     =   "Click if Father/Mother is late"
            Top             =   1065
            Width           =   2625
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   2
            Left            =   1410
            MaxLength       =   80
            TabIndex        =   4
            ToolTipText     =   "Insert Father's Name"
            Top             =   270
            Width           =   4755
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   3
            Left            =   1410
            MaxLength       =   80
            TabIndex        =   5
            ToolTipText     =   "Insert Mother's Name"
            Top             =   690
            Width           =   4755
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   285
            Left            =   1410
            TabIndex        =   7
            ToolTipText     =   "Insert marraige date"
            Top             =   1080
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd-mmm-yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "M.A.I."
            Height          =   195
            Left            =   8760
            TabIndex        =   104
            ToolTipText     =   "Monthly Average Income"
            Top             =   720
            Width           =   420
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Profession"
            Height          =   195
            Left            =   6180
            TabIndex        =   103
            ToolTipText     =   "Monthly Average Income"
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Profession"
            Height          =   195
            Left            =   6180
            TabIndex        =   99
            ToolTipText     =   "Monthly Average Income"
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "M.A.I."
            Height          =   195
            Left            =   8760
            TabIndex        =   98
            ToolTipText     =   "Monthly Average Income"
            Top             =   300
            Width           =   420
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Legal Guardian"
            Height          =   195
            Left            =   60
            TabIndex        =   75
            Top             =   1500
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Father's Name"
            Height          =   195
            Left            =   90
            TabIndex        =   49
            Top             =   300
            Width           =   1020
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mother's Name"
            Height          =   195
            Left            =   90
            TabIndex        =   48
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marraige Date"
            Height          =   195
            Left            =   90
            TabIndex        =   47
            Top             =   1080
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Student's Status"
         ForeColor       =   &H00C00000&
         Height          =   1155
         Left            =   -74910
         TabIndex        =   43
         Top             =   390
         Width           =   6225
         Begin VB.CommandButton cmdSearch 
            Height          =   300
            Left            =   2880
            Picture         =   "frmstudentInfo.frx":D16D
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   420
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   1
            Left            =   1410
            MaxLength       =   80
            TabIndex        =   3
            ToolTipText     =   "Insert Student Name"
            Top             =   690
            Width           =   4725
         End
         Begin VB.TextBox txtfields 
            Height          =   285
            Index           =   0
            Left            =   1410
            TabIndex        =   1
            ToolTipText     =   "Press ENTER to selet Student "
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Student's Name"
            Height          =   195
            Left            =   90
            TabIndex        =   45
            Top             =   720
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Student's ID"
            Height          =   195
            Left            =   90
            TabIndex        =   44
            Top             =   300
            Width           =   870
         End
      End
      Begin VB.Image Image3 
         Height          =   270
         Left            =   -65250
         Picture         =   "frmstudentInfo.frx":D44F
         Stretch         =   -1  'True
         Top             =   900
         Width           =   1215
      End
      Begin VB.Image Image4 
         Height          =   1140
         Left            =   -65310
         Picture         =   "frmstudentInfo.frx":5F191
         Stretch         =   -1  'True
         Top             =   450
         Width           =   1320
      End
      Begin VB.Shape Shape2 
         Height          =   1215
         Left            =   -65340
         Top             =   420
         Width           =   1365
      End
      Begin VB.Shape Shape1 
         Height          =   1215
         Index           =   1
         Left            =   -66810
         Top             =   420
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         Height          =   1215
         Index           =   0
         Left            =   -68490
         Top             =   420
         Width           =   1635
      End
      Begin VB.Image Image2 
         Height          =   1170
         Left            =   -66780
         Picture         =   "frmstudentInfo.frx":701BB
         Stretch         =   -1  'True
         Top             =   450
         Width           =   1395
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   96
      ToolTipText     =   "uuu"
      Top             =   7260
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmstudentInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcomputer_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    cmdInternet.SetFocus
End If

End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorDes

Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim check As String
Dim Checkb As String
Dim check2 As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
If SSTab1.Tab = 0 Then
    Set rs1 = getdata("select StuMoorFalet,computer,Internet from StudentInfo where(StudentID = '" & Trim(txtfields(0)) & "') ")
    If Not rs1.EOF Then
        check = rs1!StuMoorFalet
        Checkb = rs1!Computer
        check2 = rs1!Internet
    End If
    Set rs = getdata("select * from vaxinInfo where(StudentID = '" & Trim(txtfields(0)) & "') ")
    If rs.EOF Then
        If (MsgBox("Are You sure to delete ?", vbYesNo + vbCritical) = vbYes) Then
                 cmd.CommandType = adCmdText
                 cmd.CommandText = "Delete from StudentInfo  where (StudentID = '" & Trim(txtfields(0)) & "') "
                 cmd.Execute
                 MsgBox "Delete successfully Student Information.", vbInformation, App.Title
                 For i = 0 To 7
                  txtfields(i) = ""
                 Next
                 For i = 8 To 18
                  txtfields(i) = ""
                 Next
                 MaskEdBox1 = Format(MaskEdBox1, "__/__/__")
                 MaskEdBox2 = Format(MaskEdBox2, "__/__/__")
                 MaskEdBox3 = Format(MaskEdBox3, "__/__/__")
                  MaskEdBox4 = Format(MaskEdBox4, "__/__/__")
                 If check = "Y" Then
                cmdFM.Value = 1
                 Else
                  cmdFM.Value = 0
                 End If
                 If Checkb = "Y" Then
                 cmdcomputer.Value = 1
                 Else
                  cmdcomputer.Value = 0
                 End If
                 If check2 = "Y" Then
                 cmdInternet.Value = 1
                 Else
                  cmdInternet.Value = 0
                 End If
                 Combo1.Text = " "
                 Combo2.Text = " "
                 Combo3.Text = " "
                 Combo4.Text = " "
                 Combo5.Text = " "
                 Call ShowFlexData
        Else
                Exit Sub
        End If
    Else
        MsgBox "Vaxin Information of this Student has to remove 1st", vbInformation, App.Title
        SSTab1.Tab = 1
               
        Exit Sub
        
    End If
   
End If
If SSTab1.Tab = 1 Then
    If txtfields(0).Text = "" Then
        MsgBox "Please put stduent Id ", vbInformation, cmp
        SSTab1.Tab = 0
        txtfields(0).SetFocus
        Exit Sub
     End If
     
     Set rs = getdata("SELECT StudentID FROM StudentAdmission where Approval='Y' and studentid='" & Trim(txtfields(0)) & "'")
     If Not rs.EOF Then
        MsgBox "Admitted Student ....You can't delete Student Information,To do this first cancel student Admission", vbInformation, cmp
        Exit Sub
     End If
    
     If (MsgBox("Are You sure to delete ?", vbYesNo + vbCritical) = vbYes) Then
        
          
        cmd.CommandType = adCmdText
        cmd.CommandText = "Delete from Vaxininfo  where (StudentId = '" & Trim(txtfields(0)) & "')and VaxinName= '" & Get_Description(Trim(cmbVaccineName)) & "'"
        cmd.Execute
        MsgBox "Delete successfully Vaxin Information.", vbInformation, App.Title
        
        If cmbVaccineName.ListCount > 0 Then
            cmbVaccineName.ListIndex = 0
        End If
        
        MaskEdBox3 = Format(MaskEdBox3, "__/__/__")
        Call ShowFlexData
     Else
        Exit Sub
    End If
End If
load_cou_dist

ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
On Error GoTo ErrorDes
    cmdSAVE_Click
Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub cmdExit_Click()
On Error GoTo ErrorDes

Unload Me

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub cmdFM_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
If KeyAscii = 13 Then
    txtfields(4).SetFocus
End If

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub cmdInternet_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
If KeyAscii = 13 Then
    cmdsave.SetFocus
End If

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub cmdnew_Click()
On Error GoTo ErrorDes

Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
If SSTab1.Tab = 0 Then
  

    For i = 0 To 4
        txtfields(i) = ""
    Next
    For i = 7 To 10
        txtfields(i) = ""
    Next
    
    For i = 12 To 18
        txtfields(i) = ""
    Next
    
    MaskEdBox1 = Format(MaskEdBox1, "__/__/__")
    MaskEdBox2 = Format(MaskEdBox2, "__/__/__")
    MaskEdBox3 = Format(MaskEdBox3, "__/__/__")
    MaskEdBox4 = Format(MaskEdBox4, "__/__/__")
    txtfields(19).Text = ""
    txtfields(22).Text = ""
    txtfields(23).Text = ""
    txtfields(24).Text = ""
    txtfields(20).Text = 0
    txtfields(21).Text = 0
    txtfields(1).SetFocus
End If
If SSTab1.Tab = 1 Then
    If cmbVaccineName.ListCount > 0 Then
        cmbVaccineName.ListIndex = 0
    End If
        
    MaskEdBox3 = Format(MaskEdBox3, "__/__/__")
    MaskEdBox4 = Format(MaskEdBox3, "__/__/__")
    cmbVaccineName.SetFocus
End If
load_cou_dist

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub cmdPrivClass_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
  If KeyAscii = 13 Then
     txtfields(19).SetFocus
  End If
Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub cmdProfFather_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
  If KeyAscii = 13 Then
     txtfields(20).SetFocus
  End If
Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
  
End Sub

Private Sub cmdProfMother_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
 If KeyAscii = 13 Then
    txtfields(21).SetFocus
 End If
Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub cmdSAVE_Click()
On Error GoTo ERR_DESC
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
If SSTab1.Tab = 0 Then
    If Len(txtfields(1)) = 0 Then
        MsgBox "Please Enter Student Name.", vbCritical, "School Management System"
        txtfields(1).SetFocus
        Exit Sub
    End If
    If Len(txtfields(2)) = 0 Then
       MsgBox "Please Enter Father's Name.", vbCritical, "School Management System"
       txtfields(2).SetFocus
       Exit Sub
    End If
    If Len(txtfields(7)) = 0 Then
       MsgBox "Please Enter Religion", vbCritical, "School Management System"
       txtfields(7).SetFocus
       Exit Sub
    End If
    If Len(txtfields(3)) = 0 Then
        MsgBox "Please Enter Mother's Name.", vbCritical, "Shool Management System"
        txtfields(3).SetFocus
        Exit Sub
    End If
'     If MaskEdBox1 = "__/__/__" Then
'         MaskEdBox1.SetFocus
'     End If
     If MaskEdBox2 = "__/__/__" Then
        MsgBox "Please Enter  Birth Date.", vbCritical, "Shool Management System"
        MaskEdBox2.SetFocus
        Exit Sub
     End If
    Dim check As String
    Dim Check1 As String
    Dim check2 As String
    If cmdFM.Value = 1 Then
        check = "Y"
    Else
        check = "N"
    End If
    If cmdcomputer.Value = 1 Then
        Check1 = "Y"
    Else
        Check1 = "N"
    End If
    If cmdInternet.Value = 1 Then
        check2 = "Y"
    Else
        check2 = "N"
    End If
    Dim cntcoutrycode As String
    Dim cntcoutrycode1 As String
    Dim cntcoutrycode2 As String
    Set rs = getdata("Select cntcoutrycode from country where cntcountryname='" & Combo5.Text & "'")
    If Not rs.EOF Then
      cntcoutrycode = rs!cntcoutrycode
    End If
    Set rs1 = getdata("Select cntcoutrycode from country where cntcountryname='" & Combo1.Text & "'")
    If Not rs1.EOF Then
        cntcoutrycode1 = rs1!cntcoutrycode
    End If
    Set rs2 = getdata("Select cntcoutrycode from country where cntcountryname='" & Combo4.Text & "'")
    If Not rs2.EOF Then
        cntcoutrycode2 = rs2!cntcoutrycode
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "StudentInformation"
    cmd(1) = IIf(Len(Trim(txtfields(0))) = 0, Null, Trim(txtfields(0)))
    cmd(2) = Trim(txtfields(1))
    cmd(3) = Trim(txtfields(2))
    cmd(4) = Trim(txtfields(3))
    cmd(5) = Trim(txtfields(4))
    'cmd(6) = iifFormat(MaskEdBox1, "dd mmm yyyy")
    'cmd(6) = IIf(MaskEdBox1 = "__/__/__", Null, Format(MaskEdBox1, "dd mmm yyyy"))
    'cmd(6) = IIf(MaskEdBox1 = "__/__/__", Null, Format(MaskEdBox1, "dd mmm yyyy"))
    If MaskEdBox1 = "__/__/__" Then
        cmd(6) = Null
    Else
        cmd(6) = Format(MaskEdBox1, "dd mmm yyyy")
    End If
    cmd(7) = check
    cmd(8) = Val(txtfields(5))
    cmd(9) = Val(txtfields(6))
    cmd(10) = cntcoutrycode
    cmd(11) = Trim(txtfields(7).Text)
    cmd(12) = Format(MaskEdBox2, "dd mmm yyyy")
    cmd(13) = Check1
    cmd(14) = check2
    cmd(15) = Date
    cmd(16) = soft_user
    cmd(17) = Trim(cmdProfFather)
    cmd(18) = Trim(cmdProfMother)
    cmd(19) = Val(txtfields(20))
    cmd(20) = Val(txtfields(21))
    cmd(21) = Trim(cmdPrivClass)
    cmd(22) = txtfields(19)
    cmd(23) = txtfields(23)
    cmd(24) = txtfields(22)
    cmd(25) = txtfields(24)
    cmd(26) = IIf(MaskEdBox5.Text = "__/__/__", Null, Format(MaskEdBox5, "dd mmm yyyy"))
    cmd(27) = Null '''student Photo
    cmd(28) = Trim(cmdTobeAditted)
    For i = 1 To 16
        ss = ss & cmd(i) & ","
    Next
    txtfields(19).Text = ""
    txtfields(22).Text = ""
    txtfields(23).Text = ""
    txtfields(24).Text = ""
    
    Debug.Print ss
    
    cmd.Execute
    MsgBox "Saved Successfully.", vbInformation, "Student Management System"
    get_Maximum
    If MsgBox("More Info....", vbInformation + vbYesNo + vbDefaultButton1, cmp) = vbYes Then
          SSTab1.Tab = 1
    Else
        cmdnew.SetFocus
        If txtfields(0).Text = "" Then get_Maximum
    End If
End If
If SSTab1.Tab = 1 Then
    If Len(txtfields(8)) = 0 Then
        MsgBox "Please Enter Student's Hight .", vbCritical, "School Mangement System"
        txtfields(8).SetFocus
        Exit Sub
    End If
    If Len(txtfields(9)) = 0 Then
       MsgBox "Please Enter Student's Weight.", vbCritical, "School Management System"
       txtfields(9).SetFocus
       Exit Sub
    End If
    If Len(txtfields(10)) = 0 Then
       MsgBox "Please Enter Student's Blood Group.", vbCritical, "School Management System"
       txtfields(10).SetFocus
       Exit Sub
    End If
    
    If MaskEdBox3 = "__/__/__" Then
       MsgBox "Please Enter a valid date", vbInformation, cmp
       MaskEdBox3.SetFocus
       Exit Sub
    End If
    
     If MaskEdBox4 = "__/__/__" Then
       MsgBox "Please Enter a valid date", vbInformation, cmp
       MaskEdBox4.SetFocus
       Exit Sub
    End If
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "StudentInformation1"
    cmd(1) = Trim(txtfields(0))
    cmd(2) = Trim(txtfields(8))
    cmd(3) = Trim(txtfields(9))
    cmd(4) = Trim(txtfields(10))
    cmd(5) = Format(MaskEdBox4, "dd mmm yyyy")
    
    cmd.Execute
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "VaxinInformation"
    cmd(1) = Trim(txtfields(0))
    cmd(2) = Get_Code(Trim(cmbVaccineName))
    cmd(3) = Format(MaskEdBox3, "dd mmm yyyy")
    cmd.Execute
    MsgBox "Save Successfully.", vbInformation, "Student Management System"

    Call ShowFlexData
    cmdnew.SetFocus
End If
If SSTab1.Tab = 2 Then
    If Len(txtfields(12)) = 0 Then
        MsgBox "Please Enter Street Name Of Present Address .", vbCritical, "School Management Sysytem"
        txtfields(12).SetFocus
        Exit Sub
    End If
    If Len(txtfields(13)) = 0 Then
       MsgBox "Please Enter Street Name Of Permanant Address .", vbCritical, "School Management System"
       txtfields(13).SetFocus
       Exit Sub
    End If
    If Combo1.Text = "Bangladesh" Then
    If Len(Combo2) = 0 Then
        MsgBox "Please Enter District Of Present Address .", vbCritical, "School Management Sysytem"
        Combo2.SetFocus
        Exit Sub
    End If
End If
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "StudentInformation2"
Dim DisDistrictCode As String
Set rs = getdata("Select DisDistrictCode from District where DisDistrictName='" & Combo2.Text & "'")
DisDistrictCode = rs!DisDistrictCode
cmd(1) = Trim(txtfields(0))
cmd(2) = Trim(txtfields(12))
cmd(3) = DisDistrictCode
cmd(4) = cntcoutrycode1
cmd(5) = Trim(txtfields(13))
cmd(6) = DisDistrictCode
cmd(7) = cntcoutrycode2
cmd(8) = Trim(txtfields(14))
cmd(9) = Trim(txtfields(15))
cmd(10) = Trim(txtfields(16))
cmd(11) = Trim(txtfields(17))
cmd(12) = Trim(txtfields(18))

cmd.Execute
MsgBox "Save Successfully.", vbInformation, "Student Management System"
SSTab1.Tab = 0
cmdnew.SetFocus
load_cou_dist

End If
Exit Sub
ERR_DESC:
 MsgBox Err.Description, vbInformation, cmp
End Sub

Private Sub cmdsearch_Click()
On Error GoTo ErrorDes

Dim f As New frmFind
Set f.OwnerForm = Me
    f.intInputsel = 0
    f.SQLString = "Select StudentId, Studentname from StudentInfo order by StudentId"
    f.Show 1
    txtfields(0).SetFocus
    
Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
    
End Sub



Private Sub cmdTobeAditted_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
  If KeyAscii = 13 Then
    txtfields(22).SetFocus
  End If
  
Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
  
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
If KeyAscii = 13 Then
    Combo2.SetFocus
End If
Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
If KeyAscii = 13 Then
    txtfields(13).SetFocus
End If

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
If KeyAscii = 13 Then
    txtfields(14).SetFocus
End If

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub



Private Sub Combo4_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
If KeyAscii = 13 Then
    Combo3.SetFocus
End If

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
If KeyAscii = 13 Then
    MaskEdBox2.SetFocus

End If

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub load_cou_dist()
On Error GoTo ErrorDes
If Combo5.ListCount > 45 Then
    Combo5.ListIndex = 17
    Combo2.ListIndex = 42
    Combo1.ListIndex = 17
    Combo3.ListIndex = 42
    Combo4.ListIndex = 17
End If

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
  If KeyAscii = 13 Then
    txtfields(19).SetFocus
  End If
Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub Combo7_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
  If KeyAscii = 13 Then
    txtfields(22).SetFocus
  End If
Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
  If KeyAscii = 27 Then
    Unload Me
  ElseIf KeyAscii = 1 Then
    SSTab1.Tab = 2
  ElseIf KeyAscii = 16 Then
    SSTab1.Tab = 0
  ElseIf KeyAscii = 20 Then
    SSTab1.Tab = 1
    
  End If
Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrorDes

SSTab1.Tab = 0
With MSFlexGrid1
    .Rows = 1
    .Cols = 3
    .Col = 0: .Text = " Vaxin Name  #"
    .Col = 1: .Text = "Date"
    .Col = 2: .Text = "Next Vaxin Date"
    
    .ColWidth(0) = 5000
    .ColWidth(1) = 2550
    .ColWidth(2) = 2550
    
    
End With
Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("Select cntcoutrycode,cntcountryname from country order by cntcountryname")
If Not rs1.EOF Then
    Do Until rs1.EOF

          Combo1.AddItem rs1!CntCountryName

         Combo4.AddItem rs1!CntCountryName
         Combo5.AddItem rs1!CntCountryName
         
        rs1.MoveNext
        
    Loop
    Combo1.AddItem (" ")
    Combo4.AddItem (" ")
    Combo5.AddItem (" ")
    
'    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
End If

Set rs1 = getdata("SELECT     VaccineID, VaccineName FROM         VaccineInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
        cmbVaccineName.AddItem rs1(1) & "~" & rs1(0)
        rs1.MoveNext
    Loop
End If

If cmbVaccineName.ListCount > 0 Then
    cmbVaccineName.ListIndex = 0
End If

Set rs1 = getdata("Select disdistrictname from District order by disdistrictname")
If Not rs1.EOF Then
    Do Until rs1.EOF
        Combo2.AddItem rs1(0)
        Combo3.AddItem rs1(0)
        rs1.MoveNext
        
    Loop
        Combo2.AddItem (" ")
        Combo3.AddItem (" ")
End If
'Call ShowFlexData
load_cou_dist

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description

End Sub





Private Sub Image5_Click()

End Sub

Private Sub MaskEdBox1_GotFocus()
On Error GoTo ErrorDes
  MaskEdBox1.SelStart = 0
  MaskEdBox1.SelLength = Len(MaskEdBox1)
Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
If KeyAscii = 13 Then
    If MaskEdBox1 <> "__/__/__" Then
            If Check_ValidDate(MaskEdBox1) = False Then
                MaskEdBox1.SetFocus
                Exit Sub
            End If
    End If
    cmdFM.SetFocus
End If

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub MaskEdBox2_GotFocus()
On Error GoTo ErrorDes

  MaskEdBox2.SelStart = 0
  MaskEdBox2.SelLength = Len(MaskEdBox2)
  
Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes

If KeyAscii = 13 Then
    If MaskEdBox2 <> "__/__/__" Then
            If Check_ValidDate(MaskEdBox2) = False Then
                MaskEdBox2.SetFocus
                Exit Sub
            End If
    End If
    txtfields(7).SetFocus
End If
Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub MaskEdBox3_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes

If KeyAscii = 13 Then
    If MaskEdBox3 <> "__/__/__" Then
            If Check_ValidDate(MaskEdBox3) = False Then
                MaskEdBox3.SetFocus
                Exit Sub
            End If
    End If
MaskEdBox4.SetFocus
End If

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub MaskEdBox4_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorDes
If KeyAscii = 13 Then
    If MaskEdBox4 <> "__/__/__" Then
            If Check_ValidDate(MaskEdBox4) = False Then
                MaskEdBox4.SetFocus
                Exit Sub
            End If
    End If
cmdsave.SetFocus
End If

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub



Private Sub MaskEdBox5_GotFocus()
  MaskEdBox5.SelStart = 0
  MaskEdBox5.SelLength = Len(MaskEdBox5)
End Sub

Private Sub MaskEdBox5_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     txtfields(5).SetFocus
  End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo ErrorDes
 If SSTab1.Tab = 0 Then
     StatusBar1.Panels(1).Text = "Ctrl+P"
  ElseIf SSTab1.Tab = 1 Then
    StatusBar1.Panels(1).Text = "Ctrl+T"
  ElseIf SSTab1.Tab = 2 Then
    StatusBar1.Panels(1).Text = "Ctrl+A"
  End If
  
  
  
  
If SSTab1.Tab = 1 Then
    Call ShowFlexData
    txtfields(8).SetFocus
End If

If SSTab1.Tab = 2 Then
    txtfields(12).SetFocus

End If
Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub
Private Sub txtfields_Change(Index As Integer)
On Error GoTo ErrorDes

            StatusBar1.Panels(1).Text = ""
            Select Case Index
                    Case 5, 6
                        If Not IsNumeric(txtfields(5).Text) Then
                            txtfields(5).Text = ""
                            txtfields(6).Text = ""
                        End If
               
                    Case 7
                        If txtfields(7).Text = "i" Or txtfields(7).Text = "I" Then
                           txtfields(7).Text = "Islam"
                        ElseIf txtfields(7).Text = "c" Or txtfields(7).Text = "C" Then
                           txtfields(7).Text = "Christian"
                        ElseIf txtfields(7).Text = "b" Or txtfields(7).Text = "B" Then
                           txtfields(7).Text = "Buddist"
                        ElseIf txtfields(7).Text = "h" Or txtfields(7).Text = "H" Then
                           txtfields(7).Text = "Hindu"
                        ElseIf txtfields(7).Text = "o" Or txtfields(7).Text = "O" Then
                           txtfields(7).Text = "Others"
                           
                       End If
                    Case 20, 21
                      If Not IsNumeric(txtfields(Index)) Then
                            txtfields(Index) = ""
                      End If
                    Case 12
                            txtfields(13).Text = txtfields(12).Text
                            txtfields(16).Text = txtfields(12).Text
                            
                            
                End Select
Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub

Private Sub txtfields_DblClick(Index As Integer)
On Error GoTo ErrorDes
Select Case Index
Case 2
    txtfields(4).Text = ""
    txtfields(4).Text = txtfields(2).Text
Case 3
    txtfields(4).Text = ""
    txtfields(4).Text = txtfields(3).Text
    
    
End Select

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub


Private Sub txtfields_GotFocus(Index As Integer)
  
  txtfields(Index).SelStart = 0
  txtfields(Index).SelLength = Len(txtfields(Index))
  
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ErrorDes

StatusBar1.Panels(1).Text = ""
If KeyAscii = 13 Then
    Select Case Index
        Case 0
            
                  If Len(txtfields(0)) <> 0 Then
                       If Mid(txtfields(0), 1, 3) <> "STI" Then
                           txtfields(0).Text = Format(txtfields(0), "000000")
                           txtfields(0).Text = "STI-" + txtfields(0).Text
                   End If
                End If
            txtfields(1).SetFocus
        Case 1
            txtfields(2).SetFocus
            
        Case 2
            cmdProfFather.SetFocus
        Case 3
            cmdProfMother.SetFocus
        Case 4
             cmdPrivClass.SetFocus
       Case 5
            If IsNumeric(txtfields(5)) = True Then
            txtfields(6).SetFocus
        Else
            MsgBox "Enter Numeric Value.", vbInformation, App.Title
            txtfields(5) = ""
            txtfields(5).SetFocus
            Exit Sub
        End If
            
        Case 6
            If IsNumeric(txtfields(6)) = True Then
               Combo5.SetFocus
            Else
                
                MsgBox "Enter Numeric Value.", vbInformation, App.Title
                
                txtfields(6) = ""
                txtfields(6).SetFocus
                Exit Sub
            End If
        Case 7
            cmdcomputer.SetFocus
                   
        Case 8
            txtfields(9).SetFocus
        Case 9
            txtfields(10).SetFocus
        Case 10
            cmbVaccineName.SetFocus
        Case 11
            MaskEdBox3.SetFocus
                 
        Case 12
            Combo2.SetFocus
            
        Case 13
            Combo3.SetFocus
        Case 14
            txtfields(15).SetFocus
        Case 15
            txtfields(16).SetFocus
        Case 16
            txtfields(17).SetFocus
        Case 17
            txtfields(18).SetFocus
        Case 18
            cmdsave.SetFocus
        Case 20
          txtfields(3).SetFocus
       Case 21
         MaskEdBox1.SetFocus
       Case 19
          txtfields(23).SetFocus
       Case 23
         cmdTobeAditted.SetFocus
       Case 22
          txtfields(24).SetFocus
      Case 24
          MaskEdBox5.SetFocus
                  
    End Select
End If

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub
Public Function getdata(SQLString As String) As ADODB.Recordset

Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = SQLString
    Set rs = cmd.Execute
Set getdata = rs

End Function

Private Sub ShowFlexData()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT a.VaxinName,a.vaxindate,b.NextvaccineDate From VaxinInfo a  ,StudentInfo b where  a.studentID = b.studentID and a.studentID='" & Trim(txtfields(0)) & "'")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .TextMatrix(i, 0) = rs(0)
                .TextMatrix(i, 1) = Format(rs(1), "dd/mm/yy")
                .TextMatrix(i, 2) = Format(rs(2), "dd/mm/yy")
                
                i = i + 1
            rs.MoveNext
        Loop
    End With
Else
    MSFlexGrid1.Rows = 1
 End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub MSFlexGrid1_Click()

On Error GoTo errdes

Dim intcmbItem As Integer

For intcmbItem = 0 To cmbVaccineName.ListCount
    If Get_Code(cmbVaccineName.List(intcmbItem)) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) Then
        cmbVaccineName.ListIndex = intcmbItem
        Exit For
    End If
Next

MaskEdBox3 = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
MaskEdBox4 = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title


End Sub
Private Sub get_Maximum()
On Error GoTo ErrorDes

Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT isnull(max(StudentID),0) FROM StudentInfo")
If Not rs.EOF Then
        txtfields(0) = rs.Fields(0)
Else
    txtfields(0) = "STI-000001"
End If

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub


Private Sub txtfields_LostFocus(Index As Integer)
On Error GoTo ErrorDes

Select Case Index
    Case 0
           Dim rs As New ADODB.Recordset
           Dim rs1 As New ADODB.Recordset
           Dim check As String
           Dim Check1 As String
           Dim check2 As String
             Set rs = getdata("SELECT  * from  StudentInfo  where StudentInfo.studentid='" & Trim(txtfields(0)) & "'")
            
                If Not rs.EOF Then
                        txtfields(1) = "" & rs!StudentName
                        txtfields(2) = "" & rs!StuFatherName
                        txtfields(3) = "" & rs!StuMotherName
                        txtfields(4) = "" & rs!LegalGerdian
                        txtfields(5) = "" & rs!StuBroNo
                        txtfields(6) = "" & rs!StuSisNo
                        txtfields(7) = "" & rs!StuReligion
                        cmdProfFather = "" & rs!Father_Profession
                        cmdProfMother = "" & rs!Mother_profession
                        txtfields(20) = "" & rs!Father_Avg_Income
                        txtfields(21) = "" & rs!Mother_Avg_Income
                        cmdPrivClass = "" & rs!Previous_Class
                        cmdTobeAditted = "" & rs!TobeAdmittedClass
                        txtfields(19) = "" & rs!Previous_School_name
                        txtfields(23) = "" & rs!Previous_School_phone
                        txtfields(22) = "" & rs!Previous_School_Address
                        txtfields(24) = "" & rs!Certificate_no
                        MaskEdBox5 = IIf(IsNull(rs!Certificate_date) = True, "__/__/__", Format(rs!Certificate_date, "dd/mm/yy"))
                        ''''picture2.Image  = "" & rs!Std_Photo
                        
                        check = rs!StuMoorFalet
                        Check1 = rs!Computer
                        check2 = rs!Internet
                        If check = "Y" Then
                            cmdFM.Value = 1
                        Else
                            cmdFM.Value = 0
                        End If
                        If Check1 = "Y" Then
                            cmdcomputer.Value = 1
                        Else
                            cmdcomputer.Value = 0
                        End If
                        If check2 = "Y" Then
                            cmdInternet.Value = 1
                        Else
                            cmdInternet.Value = 0
                        End If
                         
                        MaskEdBox1 = IIf(IsNull(rs!StuMarraigeDate) = True, "__/__/__", Format(rs!StuMarraigeDate, "DD/mm/yy"))
                        MaskEdBox2 = IIf(IsNull(rs!StuDateOfBirth) = True, "__/__/__", Format(rs!StuDateOfBirth, "DD/mm/yy"))
                        MaskEdBox4 = IIf(IsNull(rs!NextvaccineDate) = True, "__/__/__", Format(rs!NextvaccineDate, "DD/mm/yy"))
                         
                         Set rs1 = getdata("select CntCountryName from Country where CntCoutryCode ='" & rs!StuCountryOfBirth & "'")
                         
                         Dim strCountryName As String
                         Dim i As Integer
                         
                         strCountryName = "" & rs1(0)
                         
                         For i = 0 To Combo5.ListIndex
                            If Combo5.List(i) = strCountryName Then
                                Combo5.ListIndex = i
                            Exit For
                            End If
                         Next
                         
                                                  
                         txtfields(8) = IIf(IsNull(rs!StuHight), "", rs!StuHight)
                         txtfields(9) = "" & rs!StuWeight
                         txtfields(10) = "" & rs!StuBloodGroup
                         txtfields(12) = "" & rs!StuStreetPAddress
                         txtfields(13) = "" & rs!StuCStreetAddress
                         txtfields(14) = "" & rs!StuPhone
                         txtfields(15) = "" & rs!StuEmail
                         txtfields(16) = "" & rs!ImmAddress
                         txtfields(17) = "" & rs!ImmPhone
                         txtfields(18) = "" & rs!ImmMob
                         
                         
                            
                         If IsNull(rs!StuPCountry) = False Then
                            Set rs1 = Nothing
                            Set rs1 = getdata("select CntCountryName from Country where CntCoutryCode ='" & rs!StuCCountry & "'")
                           If Not rs1.EOF Then
                                strCountryName = rs1(0)
                           End If
                            
                            For i = 0 To Combo1.ListIndex
                               If Combo1.List(i) = strCountryName Then
                                   Combo1.ListIndex = i
                                   Exit For
                               End If
                            Next
                                                  
                         End If
                         
                          If IsNull(rs!StuCCountry) = False Then
                            Set rs1 = Nothing
                            Set rs1 = getdata("select CntCountryName from Country where CntCoutryCode ='" & rs!StuCCountry & "'")
                            If Not rs1.EOF Then
                                strCountryName = rs1(0)
                            End If
                            
                            For i = 0 To Combo4.ListIndex
                               If Combo4.List(i) = strCountryName Then
                                   Combo4.ListIndex = i
                                   Exit For
                               End If
                            Next
                                                  
                         End If
                         
                        Set rs1 = getdata("SELECT DisDistrictName FROM   District where DisDistrictCode ='" & rs!StuCDistrict & "'")


                          If IsNull(rs!StuCDistrict) = False Then
                            Set rs1 = Nothing
                            Set rs1 = getdata("SELECT DisDistrictName FROM   District where DisDistrictCode ='" & rs!StuCDistrict & "'")
                            If Not rs1.EOF Then
                              strCountryName = rs1(0)
                            End If

                            For i = 0 To Combo3.ListIndex
                               If Combo3.List(i) = strCountryName Then
                                   Combo3.ListIndex = i
                                   Exit For
                               End If
                            Next

                         End If
                         
                         If IsNull(rs!StuCDistrict) = False Then
                            Set rs1 = Nothing
                            Set rs1 = getdata("SELECT DisDistrictName FROM   District where DisDistrictCode ='" & rs!StuPDistrict & "'")
                            If Not rs1.EOF Then
                              strCountryName = rs1(0)
                            End If

                            For i = 0 To Combo2.ListIndex
                               If Combo2.List(i) = strCountryName Then
                                   Combo2.ListIndex = i
                                   Exit For
                               End If
                            Next

                         End If
                          Set rs1 = Nothing
                          Set rs1 = getdata("SELECT VaxinName , vaxinDate FROM  VaxinInfo where StudentId ='" & Trim(txtfields(0)) & "'")
                         
                         If Not rs1.EOF Then
                             Dim intcmbItem As Integer

                            For intcmbItem = 0 To cmbVaccineName.ListCount
                                If Get_Code(cmbVaccineName.List(intcmbItem)) = rs1(0) Then
                                    cmbVaccineName.ListIndex = intcmbItem
                                    Exit For
                                End If
                            Next

                             MaskEdBox3 = IIf(IsNull(rs1!vaxinDate) = True, "__/__/__", Format(rs1!vaxinDate, "DD/mm/yy"))
                          End If
               Else
                     cmdnew_Click
                                           
            End If

   Case 15
          txtfields(15).Text = LCase(txtfields(15).Text)
          
        
End Select

Exit Sub
ErrorDes:   If Err Then MsgBox Err.Description
End Sub
