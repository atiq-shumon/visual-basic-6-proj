VERSION 5.00
Begin VB.Form frmReport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inventory Statements"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7260
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H80000003&
      Height          =   675
      Left            =   -120
      TabIndex        =   13
      Top             =   6750
      Width           =   9225
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H8000000C&
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   10
         ToolTipText     =   " Click to View Report"
         Top             =   210
         Width           =   1455
      End
      Begin VB.CommandButton cmdExit 
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
         Height          =   375
         Left            =   5715
         TabIndex        =   0
         ToolTipText     =   "Click to Close"
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000003&
      Height          =   795
      Left            =   -30
      TabIndex        =   12
      Top             =   -90
      Width           =   9015
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REPORT MANAGER"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   2340
         TabIndex        =   16
         Top             =   270
         Width           =   2685
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6225
      Left            =   -120
      TabIndex        =   11
      Top             =   600
      Width           =   7425
      Begin VB.OptionButton RptOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Yearly Requisition Statement"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   10
         Left            =   1200
         TabIndex        =   15
         Top             =   4950
         Width           =   4515
      End
      Begin VB.OptionButton RptOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Closing Stock Valuation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   9
         Left            =   1200
         TabIndex        =   14
         Top             =   4470
         Width           =   4515
      End
      Begin VB.OptionButton RptOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Minimum Stock Statements"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   8
         Left            =   1215
         TabIndex        =   9
         Top             =   3990
         Width           =   4485
      End
      Begin VB.OptionButton RptOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Item Ledger"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   1230
         TabIndex        =   7
         Top             =   3030
         Width           =   4455
      End
      Begin VB.OptionButton RptOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Adjustment Statements"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   1215
         TabIndex        =   5
         Top             =   2100
         Width           =   4485
      End
      Begin VB.OptionButton RptOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Stock/Vaule Statements"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   5
         Left            =   1230
         TabIndex        =   8
         Top             =   3510
         Width           =   4455
      End
      Begin VB.OptionButton RptOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Expire Date Statements"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   4
         Left            =   1230
         TabIndex        =   6
         Top             =   2550
         Width           =   4455
      End
      Begin VB.OptionButton RptOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Issue Statements"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   1215
         TabIndex        =   4
         Top             =   1650
         Width           =   4485
      End
      Begin VB.OptionButton RptOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Purchase Statements"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   1230
         TabIndex        =   3
         Top             =   1200
         Width           =   4455
      End
      Begin VB.OptionButton RptOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Item List"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   1230
         TabIndex        =   1
         Top             =   315
         Width           =   4455
      End
      Begin VB.OptionButton RptOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Opening Balance Statements"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   1260
         TabIndex        =   2
         Top             =   750
         Width           =   4425
      End
      Begin VB.Shape Shape1 
         Height          =   405
         Index           =   10
         Left            =   1170
         Top             =   4920
         Width           =   4575
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Index           =   9
         Left            =   1170
         Top             =   4440
         Width           =   4575
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Index           =   8
         Left            =   1170
         Top             =   3960
         Width           =   4545
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Index           =   7
         Left            =   1200
         Top             =   3480
         Width           =   4515
      End
      Begin VB.Shape Shape1 
         Height          =   345
         Index           =   6
         Left            =   1200
         Top             =   3000
         Width           =   4515
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   5
         Left            =   1200
         Top             =   2520
         Width           =   4515
      End
      Begin VB.Shape Shape1 
         Height          =   345
         Index           =   4
         Left            =   1200
         Top             =   2070
         Width           =   4515
      End
      Begin VB.Shape Shape1 
         Height          =   345
         Index           =   3
         Left            =   1200
         Top             =   1620
         Width           =   4515
      End
      Begin VB.Shape Shape1 
         Height          =   345
         Index           =   2
         Left            =   1200
         Top             =   1170
         Width           =   4515
      End
      Begin VB.Shape Shape1 
         Height          =   345
         Index           =   1
         Left            =   1200
         Top             =   720
         Width           =   4515
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   0
         Left            =   1200
         Top             =   270
         Width           =   4515
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
If RptOption(0).value = True Then
    rptmode = 1
    rptViewer.Show 1
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
    Unload Me
 End If
End Sub

Private Sub RptOption_Click(Index As Integer)
Select Case Index
    Case 0
       frmRptItemStatements.Show 1
       RptOption(0).value = False
    Case 1
        Dim f As New frmRptOpeningBal
        f.Show 1
        RptOption(1).value = False
    Case 2
        Dim f1 As New frmRptPurchaseStatements
        f1.Show 1
        RptOption(2).value = False
    Case 3
        Dim f2 As New frmRptIssueStatements
        f2.Show 1
        RptOption(3).value = False
    Case 4
        Dim f3 As New frmRptExpireDateStatements
        f3.Show 1
        RptOption(4).value = False
    Case 5
        stkFormMOde = 1
        frmRptStockStatements.Show 1
        RptOption(5).value = False
    Case 6
        Dim f5 As New frmRptAdjStatements
        f5.Show 1
        RptOption(6).value = False
   Case 7
        frmRptItemLedger.Show 1
        RptOption(7).value = False
   Case 8
       frmMinStockBal.Show 1
       RptOption(8).value = False
   Case 9
       frmRptClosingStockValuation.Show 1
       RptOption(9).value = False
   Case 10
       stkFormMOde = 2
       frmRptStockStatements.Caption = "Yearly Requisition Statement"
       frmRptStockStatements.Show 1
       RptOption(9).value = False
    
End Select
End Sub

Private Sub RptOption_GotFocus(Index As Integer)
  RptOption(Index).ForeColor = vbRed
End Sub

Private Sub RptOption_LostFocus(Index As Integer)
   RptOption(Index).ForeColor = vbBlack
End Sub
