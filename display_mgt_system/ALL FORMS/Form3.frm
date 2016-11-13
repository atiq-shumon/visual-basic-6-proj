VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form2 
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8595
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   5220
      Left            =   8010
      ScaleHeight     =   5220
      ScaleWidth      =   6855
      TabIndex        =   46
      Top             =   2880
      Width           =   6855
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   510
         Left            =   0
         ScaleHeight     =   510
         ScaleWidth      =   6855
         TabIndex        =   47
         Top             =   0
         Width           =   6855
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exchange Rate"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   18
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   2070
            TabIndex        =   48
            Top             =   45
            Width           =   2610
         End
      End
      Begin MSMask.MaskEdBox IR 
         Height          =   255
         Index           =   0
         Left            =   5085
         TabIndex        =   49
         Top             =   1485
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         Format          =   "#,##0.00;(#,##0.00)"
         Mask            =   "##.##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox IR 
         Height          =   255
         Index           =   1
         Left            =   5085
         TabIndex        =   50
         Top             =   1875
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         Mask            =   "##.##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox IR 
         Height          =   255
         Index           =   2
         Left            =   5085
         TabIndex        =   51
         Top             =   2250
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         Format          =   "#,##0.00;(#,##0.00)"
         Mask            =   "##.##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox IR 
         Height          =   255
         Index           =   3
         Left            =   5085
         TabIndex        =   52
         Top             =   2610
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         Mask            =   "##.##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox IR 
         Height          =   255
         Index           =   4
         Left            =   5085
         TabIndex        =   53
         Top             =   2955
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         Format          =   "#,##0.00;(#,##0.00)"
         Mask            =   "##.##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox IR 
         Height          =   255
         Index           =   5
         Left            =   5085
         TabIndex        =   54
         Top             =   3315
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         Format          =   "#,##0.00;(#,##0.00)"
         Mask            =   "##.##"
         PromptChar      =   " "
      End
      Begin VB.Label Label100 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   62
         Top             =   1485
         Width           =   135
      End
      Begin VB.Label Label100 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   5
         Left            =   495
         TabIndex        =   61
         Top             =   3330
         Width           =   135
      End
      Begin VB.Label Label100 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   4
         Left            =   495
         TabIndex        =   60
         Top             =   2955
         Width           =   135
      End
      Begin VB.Label Label100 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   495
         TabIndex        =   59
         Top             =   2610
         Width           =   135
      End
      Begin VB.Label Label100 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   58
         Top             =   2250
         Width           =   135
      End
      Begin VB.Label Label100 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   495
         TabIndex        =   57
         Top             =   1875
         Width           =   135
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interest"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Index           =   1
         Left            =   5145
         TabIndex        =   56
         Top             =   795
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Index           =   1
         Left            =   495
         TabIndex        =   55
         Top             =   765
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   5220
      Left            =   720
      ScaleHeight     =   5220
      ScaleWidth      =   6855
      TabIndex        =   10
      Top             =   2880
      Width           =   6855
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   510
         Left            =   0
         ScaleHeight     =   510
         ScaleWidth      =   6855
         TabIndex        =   44
         Top             =   0
         Width           =   6855
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exchange Rate"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   18
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   2070
            TabIndex        =   45
            Top             =   45
            Width           =   2610
         End
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   11
         Top             =   1425
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   12
         Top             =   1785
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   3
         Left            =   5040
         TabIndex        =   13
         Top             =   1785
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   14
         Top             =   2145
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   5
         Left            =   5040
         TabIndex        =   15
         Top             =   2145
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   6
         Left            =   2760
         TabIndex        =   16
         Top             =   2505
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   7
         Left            =   5040
         TabIndex        =   17
         Top             =   2505
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   8
         Left            =   2760
         TabIndex        =   18
         Top             =   2865
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   9
         Left            =   5040
         TabIndex        =   19
         Top             =   2865
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   10
         Left            =   2760
         TabIndex        =   20
         Top             =   3225
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   11
         Left            =   5040
         TabIndex        =   21
         Top             =   3225
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   18
         Left            =   2760
         TabIndex        =   22
         Top             =   4680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   19
         Left            =   5040
         TabIndex        =   23
         Top             =   4665
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   24
         Top             =   1425
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   12
         Left            =   2760
         TabIndex        =   38
         Top             =   3555
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   14
         Left            =   2760
         TabIndex        =   39
         Top             =   3915
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   16
         Left            =   2760
         TabIndex        =   40
         Top             =   4275
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   13
         Left            =   5040
         TabIndex        =   41
         Top             =   3555
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   15
         Left            =   5040
         TabIndex        =   42
         Top             =   3915
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ME1 
         Height          =   255
         Index           =   17
         Left            =   5040
         TabIndex        =   43
         Top             =   4275
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "#,##0.000;(#,##0.000)"
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Index           =   0
         Left            =   495
         TabIndex        =   37
         Top             =   765
         Width           =   1260
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buying"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Index           =   0
         Left            =   2910
         TabIndex        =   36
         Top             =   795
         Width           =   930
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selling"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Index           =   0
         Left            =   5145
         TabIndex        =   35
         Top             =   795
         Width           =   915
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00001"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   495
         TabIndex        =   34
         Top             =   1830
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00001"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   33
         Top             =   2190
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00001"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   495
         TabIndex        =   32
         Top             =   2565
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00001"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   4
         Left            =   495
         TabIndex        =   31
         Top             =   2910
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00001"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   5
         Left            =   495
         TabIndex        =   30
         Top             =   3270
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00001"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   6
         Left            =   495
         TabIndex        =   29
         Top             =   3630
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00001"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   7
         Left            =   495
         TabIndex        =   28
         Top             =   3990
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00001"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   8
         Left            =   495
         TabIndex        =   27
         Top             =   4365
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00001"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   9
         Left            =   495
         TabIndex        =   26
         Top             =   4710
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00001"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   25
         Top             =   1485
         Width           =   675
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   13440
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   9240
      Width           =   1455
   End
   Begin VB.TextBox DT 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox TM 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2880
      Top             =   4200
   End
   Begin MSCommLib.MSComm MC1 
      Left            =   10800
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   15240
      TabIndex        =   2
      Top             =   7980
      Width           =   15240
      Begin VB.CommandButton Command2 
         Caption         =   "E&XIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13320
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&SEND"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   12720
      TabIndex        =   8
      Top             =   9240
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Developed By:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   420
      Index           =   2
      Left            =   6345
      TabIndex        =   7
      Top             =   8520
      Width           =   2550
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   4200
      Picture         =   "Form3.frx":0442
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   6870
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Slogan/Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6368
      TabIndex        =   1
      Top             =   1770
      Width           =   2505
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAME OF THE BANK"
      BeginProperty Font 
         Name            =   "Penguin"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   840
      Left            =   4200
      TabIndex        =   0
      Top             =   960
      Width           =   6855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error Resume Next
Dim Dd, MM, YY
If Ddd = 1 Then
MC1.PortOpen = True
MC1.Settings = "1200,N,8,1"
Dd = Day(Date)
MM = Month(Date)
YY = Year(Date)
If Len(Trim(Dd)) = 1 Then
Dd = "0" + Trim(Str(Dd))
Else
Dd = Str(Dd)
End If
If Len(Trim(MM)) = 1 Then
MM = "0" + Trim(Str(MM))
Else
MM = Trim(Str(MM))
End If
dts = Dd + MM + Right(Trim(Str(Year(Date))), 2)

'Date Display
sm = dts
'--------------------------

For i = 0 To Opnrs.Fields(6) * 2 - 1
sm = sm + Left(ME1(i), 3) + Right(ME1(i), 2)
Next
sm = Trim(sm)



MC1.PortOpen = False
Close #1
Open Combo1.Text For Output As #1
Print #1, Chr$(160);
For X = 1 To Len(sm)
b$ = Mid$(sm, X, 1)
Print #1, b$;
Next X
Print #1, Chr$(13)
Close #1
End If

'===========================

If Ddd = 0 Then
MC1.PortOpen = True
MC1.Settings = "1200,N,8,1"
Dd = Day(Date)
MM = Month(Date)
YY = Year(Date)
If Len(Trim(Dd)) = 1 Then
Dd = "0" + Trim(Str(Dd))
Else
Dd = Str(Dd)
End If
If Len(Trim(MM)) = 1 Then
MM = "0" + Trim(Str(MM))
Else
MM = Trim(Str(MM))
End If
dts = Dd + MM + Right(Trim(Str(Year(Date))), 2)

'Date Display
sm = ""
'--------------------------

For i = 0 To Opnrs.Fields(6) * 2 - 1
sm = sm + Left(ME1(i), 3) + Right(ME1(i), 2)
Next
sm = Trim(sm)

MC1.PortOpen = False
Close #1
Open Combo1.Text For Output As #1
Print #1, Chr$(160);
For X = 1 To Len(sm)
b$ = Mid$(sm, X, 1)
Print #1, b$;
Next X
Print #1, Chr$(13)
Close #1
End If

MsgBox sm

If Err.Number > 0 Then
MsgBox Err.Description, vbOKOnly + vbInformation, "ERDS"
End If

End Sub

Private Sub Command2_Click()
'Me.Hide
Close #1
Open App.Path + "\Store.dat" For Random As #1 Len = Len(Dbf)
For i = 0 To ME1.UBound
If IsNumeric(ME1(i)) Then
Dbf.Dg = ME1(i).Text
Put #1, i + 1, Dbf.Dg
End If
Next
Close #1
End
End Sub

Private Sub Form_Activate()
Open App.Path + "\Store.dat" For Random As #1 Len = Len(Dbf)
If LOF(1) / Len(Dbf) > 0 Then
rn = LOF(1) / Len(Dbf)
For i = 1 To rn
Get #1, i, Dbf.Dg
If Len(Trim(Str(Dbf.Dg))) < 3 Then
DGT = String(3 - Len(Trim(Str(Dbf.Dg))), "0") + Trim(Str(Dbf.Dg))
Else
DGT = Dbf.Dg
End If
ME1(i - 1).Text = Format(DGT, "000.000")
Next
End If
Close #1

End Sub

Private Sub Form_Load()
Combo1.AddItem "COM1"
Combo1.AddItem "COM2"
Combo1.AddItem "COM3"
Combo1.AddItem "COM4"
Combo1.ListIndex = 0
Tmr = 0
Set Opnrs = Opndb.OpenRecordset("Option")
Label3.Caption = Opnrs.Fields(0)
'Label4.Caption = opnrs.Fields(0)
Label3.ForeColor = Opnrs.Fields(2)
'Label4.ForeColor = opnrs.Fields(3)


Label5.Caption = Opnrs.Fields(1)
'Label6.Caption = opnrs.Fields(1)
Label5.ForeColor = Opnrs.Fields(4)
'Label6.ForeColor = opnrs.Fields(5)
Label3.FontName = Opnrs.Fields("FontName1")
Label5.FontName = Opnrs.Fields("FontName2")
Label3.FontBold = Opnrs.Fields("Fontbold1")
Label5.FontBold = Opnrs.Fields("Fontbold2")

For i = 0 To 9
Label10(i).Visible = False
Next

For i = 0 To Opnrs.Fields(6) - 1
Label10(i).Visible = True
Next
Ddd = Opnrs.Fields("Dateop")
'MsgBox Screen.Width / 15

If Screen.Width / 15 = 800 Then
For i = 0 To 19
ME1(i).Visible = False
ME1(i).Font.Size = 8
ME1(i).Font = ME1(0).Font
ME1(i).FontSize = ME1(0).FontSize
ME1(i).Height = 255
'ME1(i).FontBold = True
Next
End If

If Screen.Width / 15 = 1024 Then
For i = 0 To 19
ME1(i).Visible = False
ME1(i).Font.Size = 10
ME1(i).Font = ME1(0).Font
ME1(i).FontSize = ME1(0).FontSize
ME1(i).Height = 300
'ME1(i).FontBold = True
Next
End If


For i = 0 To Opnrs.Fields(6) * 2 - 1
ME1(i).Visible = True
Next

For i = 0 To 5
Label100(i).Visible = False
IR(i).Visible = False
Next

For i = 0 To Opnrs.Fields("IRC") - 1
Label100(i).Visible = True
IR(i).Visible = True
Next

For i = 0 To Opnrs.Fields(6) - 1
Label10(i).Caption = Opnrs.Fields(7 + i)
Next

Label100(0).Caption = Opnrs.Fields("CR1")
Label100(1).Caption = Opnrs.Fields("CR2")
Label100(2).Caption = Opnrs.Fields("CR3")
Label100(3).Caption = Opnrs.Fields("CR4")
Label100(4).Caption = Opnrs.Fields("CR5")
Label100(5).Caption = Opnrs.Fields("CR6")

'If opnrs.Fields(28) = 1 And opnrs.Fields(6) > 5 Then
'Label7(1).Visible = True
'Label8(1).Visible = True
'Label9(1).Visible = True
'Shape1(1).Visible = True
'Line2(1).Visible = True
'Line1(2).Visible = True
'Line1(3).Visible = True
'Line1(6).Visible = True
'Else
'Label7(1).Visible = False
'Label8(1).Visible = False
'Label9(1).Visible = False
'Shape1(1).Visible = False
'Line2(1).Visible = False
'Line1(2).Visible = False
'Line1(3).Visible = False
'Line1(6).Visible = False

'Shape1(0).Left = Shape1(0).Left + 2500
'Label7(0).Left = Label7(0).Left + 2500
'Label8(0).Left = Label8(0).Left + 2500
'Label9(0).Left = Label9(0).Left + 2500
'Line1(4).X1 = Line1(4).X1 + 2500
'Line1(4).X2 = Line1(4).X2 + 2500

'Line1(0).X1 = Line1(0).X1 + 2500
'Line1(0).X2 = Line1(0).X2 + 2500

'Line1(1).X1 = Line1(1).X1 + 2500
'Line1(1).X2 = Line1(1).X2 + 2500

'Line1(5).X1 = Line1(5).X1 + 2500
'Line1(5).X2 = Line1(5).X2 + 2500

'Line2(0).X1 = Line2(0).X1 + 2500
'Line2(0).X2 = Line2(0).X2 + 2500

'For i = 0 To 9
'Label10(i).Left = Label10(i).Left + 2500
'Next

'For i = 0 To 19
'ME1(i).Left = ME1(i).Left + 2500
'Next

'End If



End Sub

Private Sub Form_Resize()
On Error Resume Next
DT.Left = Me.Width - 2000
Command2.Left = Me.Width - 2000
Label2.Left = (Me.Width - Label2.Width) / 2
Label3.Left = (Me.Width - Label3.Width) / 2
Label5.Left = (Me.Width - Label5.Width) / 2
If Picture3.Visible = False Then
Picture2.Left = (Me.Width - Picture2.Width) / 2
End If
Label7(2).Left = (Me.Width - Label7(2).Width) / 2

If Screen.Width / 15 = 800 And Tmr = 0 Then
Tmr = 1
Label7(2).Top = Me.Height - 2350
Image1.Top = Me.Height - 1800
Image1.Height = Image1.Height - (Image1.Height * 20) / 100
Image1.Width = Image1.Width - (Image1.Width * 20) / 100
End If

If Screen.Width / 15 = 800 Then
Label3.Top = Label2.Top + 500
Label5.Top = Label3.Top + 800
Picture2.Top = Label5.Top + 400
End If

If Screen.Width / 15 = 1024 Then
Combo1.Left = Me.Width - 2000
Combo1.Top = Image1.Top
'Label1.Top = Image1.Top
'Label1.Left = Me.Width - 2600
End If

If Screen.Width / 15 = 800 Then
Combo1.Left = Me.Width - 1800
Combo1.Top = Image1.Top
'Label1.Top = Image1.Top
'Label1.Left = Me.Width - 2400
End If


Image1.Left = (Me.Width - Image1.Width) / 2

'If Me.WindowState <> vbMinimized Or Me.WindowState <> vbMaximized Then
'Me.Left = 0
'Me.Top = 0
'Me.Width = Screen.Width
'Me.Height = Screen.Height
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #1
Open App.Path + "\Store.dat" For Random As #1 Len = Len(Dbf)
For i = 0 To ME1.UBound
If IsNumeric(ME1(i)) Then
Dbf.Dg = ME1(i).Text
Put #1, i + 1, Dbf.Dg
End If
Next
Close #1
End Sub

Private Sub ME1_GotFocus(Index As Integer)
ME1(Index).SelStart = 0
ME1(Index).SelLength = 7
End Sub

Private Sub ME1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
Dim dgf As String
MM = Val(ME1(Index).Text)
If Len(MM) < 3 Then
dgf = String(3 - Len(MM), "0") + Trim(ME1(Index).Text)
ME1(Index).Text = Format(Val(dgf), "000.000")
End If
If Index < Opnrs.Fields(6) * 2 - 1 Then
ME1(Index + 1).SelStart = 0
ME1(Index + 1).SelLength = Len(ME1(Index + 1))
ME1(Index + 1).SetFocus
End If
End If
End Sub

Private Sub ME1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
If Index > 0 Then
ME1(Index - 1).SetFocus
End If
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
TM.Text = Format(Time, "hh:mm:ss AMPM")
DT.Text = Format(Date, "dd/mm/yyyy")
End Sub

