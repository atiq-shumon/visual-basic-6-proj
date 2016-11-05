VERSION 5.00
Begin VB.Form frmRptItemStatements 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Item Statements"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7095
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Statement"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1950
         TabIndex        =   11
         Top             =   150
         Width           =   2595
      End
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      Caption         =   "Type && Group Wise"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   3300
      TabIndex        =   8
      Top             =   960
      Width           =   3105
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmRptItemStatements.frx":0000
      Left            =   1590
      List            =   "frmRptItemStatements.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2160
      Width           =   4785
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      Caption         =   "Category Wise"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   1050
      TabIndex        =   5
      Top             =   960
      Width           =   2715
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   960
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmRptItemStatements.frx":0004
      Left            =   1590
      List            =   "frmRptItemStatements.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   4785
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5280
      TabIndex        =   1
      ToolTipText     =   "Click to Save"
      Top             =   3210
      Width           =   1065
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H8000000C&
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4200
      TabIndex        =   0
      ToolTipText     =   "Click to View Report"
      Top             =   3210
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   -120
      TabIndex        =   9
      Top             =   3030
      Width           =   7095
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   1575
      Left            =   -60
      Top             =   1410
      Width           =   7155
   End
   Begin VB.Shape Shape1 
      Height          =   435
      Left            =   4410
      Top             =   3180
      Width           =   2025
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name"
      Height          =   345
      Left            =   330
      TabIndex        =   7
      Top             =   2250
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category Name"
      Enabled         =   0   'False
      Height          =   345
      Left            =   330
      TabIndex        =   3
      Top             =   1650
      Width           =   1095
   End
End
Attribute VB_Name = "frmRptItemStatements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objcom As New DSLComFram.clsCommon
Dim objRs As New ADODB.Recordset

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
If Option1(0).value = True Then
    optionMode = 1
    rptmode = 1
    rptViewer.Show 1
ElseIf Option1(1).value = True Then  '''''''''category wise

   If Len(Combo1) = 0 Then
    MsgBox "Category Required", vbInformation, App.title
    Combo1.SetFocus
    Exit Sub
  End If
    optionMode = 2
   rptmode = 1
   rptViewer.Show 1
ElseIf Option1(2).value = True Then  '''''''''''category and groupwise
   If Len(Combo1) = 0 Then
      MsgBox "Category Required", vbInformation, App.title
      Combo1.SetFocus
      Exit Sub
    End If
  If Len(Combo2) = 0 Then
     MsgBox "Group Required", vbInformation, App.title
     Combo2.SetFocus
     Exit Sub
  End If
   optionMode = 3
   rptmode = 1
   rptViewer.Show 1

End If
  

End Sub



Private Sub Combo1_Click()
   load_group
End Sub
Private Sub load_group()
  Set objRs = objcom.Get_RS("SELECT group_code,group_name from item_group_info where type_code='" & Get_Code(Combo1) & "'", objmyCon)
  Combo2.Clear
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       Combo2.AddItem objRs(1) + "~" + objRs(0)
       objRs.MoveNext
    Loop
End If

End Sub
Private Sub Form_Load()
Set objRs = objcom.Get_RS("SELECT type_code,type_name from item_type_info where cate_code='" & CategoryCode & "'", objmyCon)
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       Combo1.AddItem objRs(1) + "~" + objRs(0)
       objRs.MoveNext
    Loop
End If
End Sub

Private Sub Option1_Click(Index As Integer)
  Select Case Index
         Case 0
               Combo1.Enabled = False
               Combo2.Enabled = False
               Label1.Enabled = False
               Label2.Enabled = False
         Case 1
               Combo1.Enabled = True
               Combo2.Enabled = False
               Label1.Enabled = True
               Label2.Enabled = False
         Case 2
               Combo1.Enabled = True
               Combo2.Enabled = True
               Label1.Enabled = True
               Label2.Enabled = True
         End Select
  End Sub
