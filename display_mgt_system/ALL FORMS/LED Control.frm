VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Computer Control Status Monitoring System"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13800
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   13800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   9480
      TabIndex        =   52
      Top             =   8760
      Width           =   1095
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   375
      Left            =   5520
      TabIndex        =   51
      Top             =   8760
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   375
      Left            =   2640
      TabIndex        =   19
      Top             =   8760
      Width           =   375
   End
   Begin VB.TextBox Text3 
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
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Text            =   "OFF"
      Top             =   8760
      Width           =   975
   End
   Begin VB.TextBox Text2 
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
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Text            =   "85"
      Top             =   8760
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   12360
      TabIndex        =   2
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEND"
      Height          =   375
      Left            =   10920
      TabIndex        =   1
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12960
      TabIndex        =   50
      Top             =   8160
      Width           =   495
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   49
      Top             =   8160
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   199
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   48
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   198
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   47
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   197
      Left            =   8880
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   46
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   196
      Left            =   9480
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9360
      TabIndex        =   45
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   195
      Left            =   10080
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   44
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   194
      Left            =   10680
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   43
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   193
      Left            =   11280
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11160
      TabIndex        =   42
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   192
      Left            =   11880
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11760
      TabIndex        =   41
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   191
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12360
      TabIndex        =   40
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   190
      Left            =   13080
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12960
      TabIndex        =   39
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      Caption         =   "101"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   38
      Top             =   2280
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   189
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   188
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   187
      Left            =   8880
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   186
      Left            =   9480
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   185
      Left            =   10080
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   184
      Left            =   10680
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   183
      Left            =   11280
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   182
      Left            =   11880
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   181
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   180
      Left            =   13080
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   179
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   178
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   177
      Left            =   8880
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   176
      Left            =   9480
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   175
      Left            =   10080
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   174
      Left            =   10680
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   173
      Left            =   11280
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   172
      Left            =   11880
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   171
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   170
      Left            =   13080
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   169
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   168
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   167
      Left            =   8880
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   166
      Left            =   9480
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   165
      Left            =   10080
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   164
      Left            =   10680
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   163
      Left            =   11280
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   162
      Left            =   11880
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   161
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   160
      Left            =   13080
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   159
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   158
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   157
      Left            =   8880
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   156
      Left            =   9480
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   155
      Left            =   10080
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   154
      Left            =   10680
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   153
      Left            =   11280
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   152
      Left            =   11880
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   151
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   150
      Left            =   13080
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   149
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   148
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   147
      Left            =   8880
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   146
      Left            =   9480
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   145
      Left            =   10080
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   144
      Left            =   10680
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   143
      Left            =   11280
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   142
      Left            =   11880
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   141
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   140
      Left            =   13080
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   139
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   138
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   137
      Left            =   8880
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   136
      Left            =   9480
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   135
      Left            =   10080
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   134
      Left            =   10680
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   133
      Left            =   11280
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   132
      Left            =   11880
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   131
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   130
      Left            =   13080
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   129
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   128
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   127
      Left            =   8880
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   126
      Left            =   9480
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   125
      Left            =   10080
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   124
      Left            =   10680
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   123
      Left            =   11280
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   122
      Left            =   11880
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   121
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   120
      Left            =   13080
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   119
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   118
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   117
      Left            =   8880
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   116
      Left            =   9480
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   115
      Left            =   10080
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   114
      Left            =   10680
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   113
      Left            =   11280
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   112
      Left            =   11880
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   111
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   110
      Left            =   13080
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   109
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   108
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   107
      Left            =   8880
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   106
      Left            =   9480
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   105
      Left            =   10080
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   104
      Left            =   10680
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   103
      Left            =   11280
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   102
      Left            =   11880
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   101
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   100
      Left            =   13080
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      Caption         =   "111"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   37
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      Caption         =   "121"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   36
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      Caption         =   "131"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   35
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Caption         =   "141"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   34
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Caption         =   "151"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   33
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Caption         =   "161"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   32
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Caption         =   "171"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   31
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Caption         =   "181"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   30
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Caption         =   "191"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   29
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Caption         =   "91"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Caption         =   "81"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "71"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "61"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "51"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "41"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   99
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   98
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   97
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   96
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   95
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   94
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   93
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   92
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   91
      Left            =   840
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   90
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   89
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   88
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   87
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   86
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   85
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   84
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   83
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   82
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   81
      Left            =   840
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   80
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   79
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   78
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   77
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   76
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   75
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   74
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   73
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   72
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   71
      Left            =   840
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   70
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   69
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   68
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   67
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   66
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   65
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   64
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   63
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   62
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   61
      Left            =   840
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   60
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   59
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   58
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   57
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   56
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   55
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   54
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   53
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   52
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   51
      Left            =   840
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   50
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   49
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   48
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   47
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   46
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   45
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   44
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   43
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   42
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   41
      Left            =   840
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   40
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   39
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   38
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   37
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   36
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   35
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   34
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   33
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   32
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   31
      Left            =   840
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   30
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   29
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   28
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   27
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   26
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   25
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   24
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   23
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   22
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   21
      Left            =   840
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   20
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   19
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   18
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   17
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   16
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   15
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   14
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   13
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   12
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   11
      Left            =   840
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   10
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   9
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   16
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   8
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   7
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   6
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   5
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   4
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "LOCATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   8760
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Developed By: Digilog Systems"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   840
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   840
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Computer Control Status Monitoring System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   10335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

End Sub
